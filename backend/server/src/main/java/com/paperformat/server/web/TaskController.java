package com.paperformat.server.web;

import com.paperformat.core.FormatPlan;
import com.paperformat.server.task.TaskService;
import com.paperformat.server.task.TaskView;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.Map;
import java.util.List;

/**
 * 前端调用的文件处理 API。
 */
@RestController
@RequestMapping("/api/v1")
public class TaskController {
    private final TaskService taskService;

    public TaskController(TaskService taskService) {
        this.taskService = taskService;
    }

    @GetMapping("/health")
    public Map<String, Object> health() {
        return Map.of(
                "status", "UP",
                "service", "paper-format-server"
        );
    }

    /**
     * 接收模板和待处理文档，创建异步处理任务。
     */
    @PostMapping(path = "/tasks", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<TaskView> createTask(
            @RequestPart("template") MultipartFile template,
            @RequestPart("document") MultipartFile document
    ) throws IOException {
        return ResponseEntity.accepted().body(taskService.submit(template, document));
    }

    @GetMapping("/tasks/{taskId}")
    public TaskView task(@PathVariable String taskId) {
        return taskService.get(taskId);
    }

    @GetMapping("/tasks/{taskId}/analysis")
    public FormatPlan analysis(@PathVariable String taskId) throws IOException {
        return taskService.analysis(taskId);
    }

    @PostMapping(path = "/tasks/{taskId}/confirm", consumes = MediaType.APPLICATION_JSON_VALUE)
    public ResponseEntity<TaskView> confirm(
            @PathVariable String taskId,
            @RequestBody ConfirmationRequest request
    ) throws IOException {
        return ResponseEntity.accepted().body(taskService.confirm(
                taskId, request.enabledRuleKeys(), request.acceptedIssueKeys()));
    }

    /**
     * 下载生成后的 DOCX。只有 COMPLETED 任务能下载。
     */
    @GetMapping("/tasks/{taskId}/result")
    public ResponseEntity<Resource> result(@PathVariable String taskId) throws IOException {
        Path path = taskService.resultPath(taskId);
        return attachment(path, MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
    }

    /**
     * 下载格式处理 JSON 报告。
     */
    @GetMapping("/tasks/{taskId}/report")
    public ResponseEntity<Resource> report(@PathVariable String taskId) throws IOException {
        return attachment(taskService.reportPath(taskId), MediaType.APPLICATION_JSON);
    }

    private ResponseEntity<Resource> attachment(Path path, MediaType mediaType) throws IOException {
        Resource resource = new FileSystemResource(path);
        // 使用 UTF-8 文件名，避免中文文件名在浏览器下载时乱码。
        ContentDisposition disposition = ContentDisposition.attachment()
                .filename(path.getFileName().toString(), StandardCharsets.UTF_8)
                .build();
        return ResponseEntity.ok()
                .contentType(mediaType)
                .contentLength(resource.contentLength())
                .header(HttpHeaders.CONTENT_DISPOSITION, disposition.toString())
                .body(resource);
    }

    public record ConfirmationRequest(List<String> enabledRuleKeys, List<String> acceptedIssueKeys) {
    }
}
