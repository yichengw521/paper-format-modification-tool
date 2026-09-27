package com.paperformat.server.task;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.paperformat.core.ProcessingReport;
import com.paperformat.core.FormatPlan;
import com.paperformat.core.WordFormatProcessor;
import jakarta.annotation.PostConstruct;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.OffsetDateTime;
import java.time.ZoneOffset;
import java.util.Map;
import java.util.List;
import java.util.LinkedHashSet;
import java.util.NoSuchElementException;
import java.util.Optional;
import java.util.UUID;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.Executor;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * 任务编排服务。
 *
 * <p>负责接收上传文件、保存任务状态、异步调用核心处理器，并提供结果文件路径。</p>
 */
@Service
public class TaskService {
    private static final String RESULT_FILE = "formatted-document.docx";
    private static final String REPORT_FILE = "format-report.json";
    private static final String ANALYSIS_FILE = "format-plan.json";
    private static final String CONFIRMATION_FILE = "confirmation.json";

    private final Path taskRoot;
    private final ObjectMapper objectMapper;
    private final Executor executor;
    // 当前版本用内存 Map 做快速查询，同时把 task.json 写到磁盘支持重启后恢复历史任务。
    private final Map<String, TaskView> tasks = new ConcurrentHashMap<>();

    public TaskService(
            @Value("${paper-format.storage-root:./storage}") String storageRoot,
            ObjectMapper objectMapper,
            @Qualifier("formatTaskExecutor") Executor executor
    ) {
        this.taskRoot = Path.of(storageRoot).toAbsolutePath().normalize().resolve("tasks");
        this.objectMapper = objectMapper;
        this.executor = executor;
    }

    @PostConstruct
    void initialize() throws IOException {
        Files.createDirectories(taskRoot);
        try (var directories = Files.list(taskRoot)) {
            // 服务重启时加载历史 task.json；未完成任务会标记为失败，避免永远卡在处理中。
            directories.filter(Files::isDirectory).forEach(this::loadExistingTask);
        }
    }

    /**
     * 创建分析任务。此阶段只提取规则，不生成修改后的文档。
     */
    public TaskView submit(MultipartFile template, MultipartFile document) throws IOException {
        validateDocx(template, "template");
        validateDocx(document, "document");

        String id = UUID.randomUUID().toString();
        Path taskDirectory = taskDirectory(id);
        Path inputDirectory = taskDirectory.resolve("input");
        Path outputDirectory = taskDirectory.resolve("output");
        Files.createDirectories(inputDirectory);
        Files.createDirectories(outputDirectory);

        Path templatePath = inputDirectory.resolve("template.docx");
        Path documentPath = inputDirectory.resolve("document.docx");
        copyUpload(template, templatePath);
        copyUpload(document, documentPath);

        OffsetDateTime now = OffsetDateTime.now(ZoneOffset.UTC);
        TaskView task = new TaskView(
                id,
                TaskStatus.ANALYZING,
                now,
                now,
                safeOriginalName(template, "template.docx"),
                safeOriginalName(document, "document.docx"),
                RESULT_FILE,
                REPORT_FILE,
                ANALYSIS_FILE,
                null,
                null,
                null,
                null,
                links(id)
        );
        tasks.put(id, task);
        persist(task);
        // 先分析；必须等用户确认后才会进入真正的修改线程。
        executor.execute(() -> analyze(id, templatePath, documentPath, outputDirectory));
        return task;
    }

    public FormatPlan analysis(String id) throws IOException {
        TaskView task = get(id);
        if (task.status() != TaskStatus.AWAITING_CONFIRMATION
                && task.status() != TaskStatus.QUEUED
                && task.status() != TaskStatus.PROCESSING
                && task.status() != TaskStatus.COMPLETED) {
            throw new IllegalStateException("Format plan is not ready: " + task.status());
        }
        return objectMapper.readValue(taskDirectory(id).resolve("output").resolve(ANALYSIS_FILE).toFile(), FormatPlan.class);
    }

    /** Confirm the selected rules and start the irreversible processing stage. */
    public synchronized TaskView confirm(
            String id,
            List<String> enabledRuleKeys,
            List<String> acceptedIssueKeys
    ) throws IOException {
        TaskView current = get(id);
        if (current.status() != TaskStatus.AWAITING_CONFIRMATION) {
            throw new IllegalStateException("Task is not awaiting confirmation: " + current.status());
        }
        FormatPlan plan = analysis(id);
        LinkedHashSet<String> allowed = new LinkedHashSet<>();
        for (FormatPlan.Rule rule : plan.rules()) allowed.add(rule.key());
        LinkedHashSet<String> selected = new LinkedHashSet<>();
        if (enabledRuleKeys != null) {
            for (String key : enabledRuleKeys) if (allowed.contains(key)) selected.add(key);
        }
        if (selected.isEmpty()) {
            throw new IllegalArgumentException("Please select at least one format rule.");
        }
        LinkedHashSet<String> allowedIssues = new LinkedHashSet<>();
        if (plan.issues() != null) {
            for (FormatPlan.Issue issue : plan.issues()) allowedIssues.add(issue.key());
        }
        LinkedHashSet<String> accepted = new LinkedHashSet<>();
        if (acceptedIssueKeys != null) {
            for (String key : acceptedIssueKeys) if (allowedIssues.contains(key)) accepted.add(key);
        }
        Path directory = taskDirectory(id);
        objectMapper.writerWithDefaultPrettyPrinter().writeValue(
                directory.resolve("output").resolve(CONFIRMATION_FILE).toFile(),
                Map.of(
                        "confirmedAt", OffsetDateTime.now(ZoneOffset.UTC).toString(),
                        "enabledRuleKeys", selected,
                        "acceptedIssueKeys", accepted));
        update(id, TaskStatus.QUEUED, null, null, null, null);
        executor.execute(() -> process(
                id,
                directory.resolve("input").resolve("template.docx"),
                directory.resolve("input").resolve("document.docx"),
                directory.resolve("output"),
                selected,
                accepted
        ));
        return get(id);
    }

    /**
     * 根据任务 ID 查询当前状态。
     */
    public TaskView get(String id) {
        return Optional.ofNullable(tasks.get(id))
                .orElseThrow(() -> new NoSuchElementException("Task not found: " + id));
    }

    /**
     * 获取处理完成后的 DOCX 路径。任务未完成时会抛出冲突异常。
     */
    public Path resultPath(String id) {
        TaskView task = requireCompleted(id);
        return taskDirectory(task.id()).resolve("output").resolve(RESULT_FILE);
    }

    /**
     * 获取处理完成后的 JSON 报告路径。任务未完成时会抛出冲突异常。
     */
    public Path reportPath(String id) {
        TaskView task = requireCompleted(id);
        return taskDirectory(task.id()).resolve("output").resolve(REPORT_FILE);
    }

    /**
     * 后台执行实际 DOCX 处理，并把结果摘要写回任务状态。
     */
    private void analyze(String id, Path templatePath, Path documentPath, Path outputDirectory) {
        try {
            FormatPlan plan = new WordFormatProcessor().analyze(templatePath, documentPath);
            objectMapper.writerWithDefaultPrettyPrinter().writeValue(
                    outputDirectory.resolve(ANALYSIS_FILE).toFile(), plan);
            update(id, TaskStatus.AWAITING_CONFIRMATION, null,
                    plan.documentSummary().tables(), plan.documentSummary().sections(), null);
        } catch (Exception exception) {
            update(id, TaskStatus.FAILED, null, null, null, rootMessage(exception));
        }
    }

    private void process(
            String id,
            Path templatePath,
            Path documentPath,
            Path outputDirectory,
            java.util.Set<String> enabledRuleKeys,
            java.util.Set<String> acceptedIssueKeys
    ) {
        update(id, TaskStatus.PROCESSING, null, null, null, null);
        try {
            ProcessingReport report = new WordFormatProcessor().process(
                    templatePath,
                    documentPath,
                    outputDirectory.resolve(RESULT_FILE),
                    outputDirectory.resolve(REPORT_FILE),
                    enabledRuleKeys,
                    acceptedIssueKeys
            );
            update(
                    id,
                    TaskStatus.COMPLETED,
                    report.modifications().size(),
                    report.documentSummary().tables(),
                    report.documentSummary().sections(),
                    null
            );
        } catch (Exception exception) {
            update(id, TaskStatus.FAILED, null, null, null, rootMessage(exception));
        }
    }

/**
 * 更新任务信息的方法
 * 这是一个同步方法，确保线程安全
 * @param id 任务ID
 * @param status 任务状态
 * @param modifications 修改次数
 * @param tables 表格数量
 * @param sections 章节数量
 * @param error 错误信息
 */
    private synchronized void update(
            String id,           // 任务唯一标识符
            TaskStatus status,   // 任务状态枚举
            Integer modifications, // 修改次数计数
            Integer tables,      // 表格数量计数
            Integer sections,    // 章节数量计数
            String error         // 错误信息描述
    ) {
        TaskView current = get(id);            // 获取当前任务视图
        TaskView updated = new TaskView(        // 创建更新后的任务视图对象
                current.id(),                  // 保持原任务ID
                status,                        // 更新任务状态
                current.createdAt(),           // 保持创建时间不变
                OffsetDateTime.now(ZoneOffset.UTC), // 更新最后修改时间为当前UTC时间
                current.templateFileName(),    // 保持模板文件名不变
                current.documentFileName(),    // 保持文档文件名不变
                current.resultFileName(),     // 保持结果文件名不变
                current.reportFileName(),     // 保持报告文件名不变
                current.analysisFileName(),
                modifications,                // 更新修改次数
                tables,                       // 更新表格数量
                sections,                     // 更新章节数量
                error,                        // 更新错误信息
                current.links()               // 保持链接不变
        );
        tasks.put(id, updated);                // 将更新后的任务视图存储到tasks映射中
        persistQuietly(updated);              // 静默持久化更新后的任务视图
    }

/**
 * 根据任务ID获取已完成的任务视图
 * @param id 任务ID
 * @return 已完成的任务视图
 * @throws IllegalStateException 如果任务未完成
 */
    private TaskView requireCompleted(String id) {
    // 根据ID获取任务
        TaskView task = get(id);
    // 检查任务状态是否为已完成
        if (task.status() != TaskStatus.COMPLETED) {
        // 如果任务未完成，抛出非法状态异常
            throw new IllegalStateException("Task is not completed: " + task.status());
        }
    // 返回已完成的任务
        return task;
    }

/**
 * 验证上传的文件是否为有效的DOCX文档
 * @param file 上传的文件对象
 * @param fieldName 字段名称，用于错误提示
 * @throws IOException 如果文件读取过程中发生错误
 */
    private void validateDocx(MultipartFile file, String fieldName) throws IOException {
    // 检查文件是否为空
        if (file == null || file.isEmpty()) {
            throw new IllegalArgumentException(fieldName + " must not be empty.");
        }
    // 获取安全的原始文件名，如果为空则使用默认值
        String originalName = safeOriginalName(file, "");
    // 检查文件扩展名是否为.docx（不区分大小写）
        if (!originalName.toLowerCase().endsWith(".docx")) {
            throw new IllegalArgumentException(fieldName + " must be a .docx file.");
        }



    // 标志位，用于验证DOCX文档必需的部件
        boolean contentTypes = false;  // [Content_Types].xml 文件标志
        boolean documentPart = false;  // word/document.xml 文件标志
        try (InputStream input = file.getInputStream(); ZipInputStream zip = new ZipInputStream(input)) {
            ZipEntry entry;
        // 遍历ZIP包中的所有条目
            while ((entry = zip.getNextEntry()) != null) {
                // DOCX 本质是 ZIP 包；这两个部件是最基础的 Word 文档结构标志。
                contentTypes |= "[Content_Types].xml".equals(entry.getName());
                documentPart |= "word/document.xml".equals(entry.getName());
            // 如果两个必需部件都找到，则提前终止循环
                if (contentTypes && documentPart) {
                    break;
                }
            }
        } catch (IOException exception) {
        // 如果读取过程中发生错误，抛出异常说明文件不是一个可读取的DOCX包
            throw new IllegalArgumentException(fieldName + " is not a readable DOCX package.", exception);
        }
    // 验证是否包含DOCX文档必需的两个核心部件
        if (!contentTypes || !documentPart) {
            throw new IllegalArgumentException(fieldName + " is not a valid DOCX package.");
        }
    }

    private void copyUpload(MultipartFile file, Path destination) throws IOException {
        try (InputStream input = file.getInputStream()) {
            Files.copy(input, destination, StandardCopyOption.REPLACE_EXISTING);
        }
    }

    private void loadExistingTask(Path directory) {
        Path stateFile = directory.resolve("task.json");
        if (!Files.isRegularFile(stateFile)) {
            return;
        }
        try {
            TaskView task = objectMapper.readValue(stateFile.toFile(), TaskView.class);
            if (task.status() == TaskStatus.ANALYZING || task.status() == TaskStatus.QUEUED || task.status() == TaskStatus.PROCESSING) {
                task = new TaskView(
                        task.id(), TaskStatus.FAILED, task.createdAt(), OffsetDateTime.now(ZoneOffset.UTC),
                        task.templateFileName(), task.documentFileName(), task.resultFileName(), task.reportFileName(), task.analysisFileName(),
                        task.modifications(), task.tables(), task.sections(),
                        "The server restarted before this task completed.", task.links()
                );
                persist(task);
            }
            tasks.put(task.id(), task);
        } catch (Exception ignored) {
            // A corrupt task state must not prevent the service from starting.
        }
    }

    private void persist(TaskView task) throws IOException {
        Path destination = taskDirectory(task.id()).resolve("task.json");
        Path temporary = destination.resolveSibling("task.json.tmp");
        objectMapper.writerWithDefaultPrettyPrinter().writeValue(temporary.toFile(), task);
        // 先写临时文件再替换，降低进程中断时留下半个 JSON 的概率。
        Files.move(temporary, destination, StandardCopyOption.REPLACE_EXISTING);
    }

    private void persistQuietly(TaskView task) {
        try {
            persist(task);
        } catch (IOException ignored) {
            // Processing results remain available in memory; the API status is still accurate.
        }
    }

    private Path taskDirectory(String id) {
        return taskRoot.resolve(id);
    }

    private static String safeOriginalName(MultipartFile file, String fallback) {
        String name = file.getOriginalFilename();
        if (name == null || name.isBlank()) {
            return fallback;
        }
        return Path.of(name).getFileName().toString();
    }

    private static Map<String, String> links(String id) {
        String base = "/api/v1/tasks/" + id;
        return Map.of(
                "status", base,
                "analysis", base + "/analysis",
                "confirm", base + "/confirm",
                "result", base + "/result",
                "report", base + "/report"
        );
    }

/**
 * 获取异常的根异常信息
 * 该方法会遍历异常链，找到最初的异常（即没有cause的异常），
 * 然后返回该异常的信息。如果异常信息为空或空白字符串，则返回异常类名
 *
 * @param throwable 要处理的异常对象
 * @return 根异常的信息，如果信息为空则返回异常类名
 */
    private static String rootMessage(Throwable throwable) {
    // 设置当前异常为传入的异常
        Throwable current = throwable;
    // 循环遍历异常链，直到找到最底层的异常（没有cause的异常）
        while (current.getCause() != null) {
            current = current.getCause();
        }
    // 获取最底层异常的信息
        String message = current.getMessage();
    // 判断异常信息是否为空或空白字符串，如果是则返回异常类名，否则返回异常信息
        return message == null || message.isBlank() ? current.getClass().getSimpleName() : message;
    }
}
