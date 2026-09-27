package com.paperformat.core;

import java.nio.file.Path;
import java.util.Arrays;
import java.util.LinkedHashSet;
import java.util.Set;

/**
 * 命令行入口，用于不启动 Spring Boot 服务时直接处理一组 DOCX 文件。
 */
public final class PaperFormatCli {
    private PaperFormatCli() {
    }

    /**
     * 参数顺序固定为：模板、源文档、输出文档、JSON 报告，以及可选的逗号分隔结构修正键。
     */
    public static void main(String[] args) throws Exception {
        if (args.length < 4 || args.length > 5) {
            System.err.println("Usage: PaperFormatCli <template.docx> <source.docx> <output.docx> <report.json> [acceptedIssueKeys]");
            System.exit(2);
        }

        // 统一转成绝对规范路径，避免工作目录变化导致输出位置不明确。
        Path template = Path.of(args[0]).toAbsolutePath().normalize();
        Path source = Path.of(args[1]).toAbsolutePath().normalize();
        Path output = Path.of(args[2]).toAbsolutePath().normalize();
        Path report = Path.of(args[3]).toAbsolutePath().normalize();

        Set<String> acceptedIssueKeys = args.length == 5 && !args[4].isBlank()
                ? new LinkedHashSet<>(Arrays.asList(args[4].split(",")))
                : Set.of();
        WordFormatProcessor processor = new WordFormatProcessor();
        FormatPlan plan = processor.analyze(template, source);
        if (!plan.issues().isEmpty()) {
            System.out.println("Detected structure issues:");
            for (FormatPlan.Issue issue : plan.issues()) {
                System.out.printf("- %s: %s -> %s%n", issue.key(), issue.originalText(), issue.suggestedText());
            }
        }
        ProcessingReport result = processor.process(
                template, source, output, report, null, acceptedIssueKeys);
        System.out.printf(
                "Completed: %d paragraph style fixes, %d section layout checks, %d tables inspected.%nOutput: %s%nReport: %s%n",
                result.modifications().size(),
                result.sectionSummary().sectionsInspected(),
                result.documentSummary().tables(),
                output,
                report
        );
    }
}
