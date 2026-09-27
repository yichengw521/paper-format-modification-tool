package com.paperformat.core;

import java.nio.file.Path;

/**
 * 命令行入口，用于不启动 Spring Boot 服务时直接处理一组 DOCX 文件。
 */
public final class PaperFormatCli {
    private PaperFormatCli() {
    }

    /**
     * 参数顺序固定为：模板、源文档、输出文档、JSON 报告。
     */
    public static void main(String[] args) throws Exception {
        if (args.length != 4) {
            System.err.println("Usage: PaperFormatCli <template.docx> <source.docx> <output.docx> <report.json>");
            System.exit(2);
        }

        // 统一转成绝对规范路径，避免工作目录变化导致输出位置不明确。
        Path template = Path.of(args[0]).toAbsolutePath().normalize();
        Path source = Path.of(args[1]).toAbsolutePath().normalize();
        Path output = Path.of(args[2]).toAbsolutePath().normalize();
        Path report = Path.of(args[3]).toAbsolutePath().normalize();

        ProcessingReport result = new WordFormatProcessor().process(template, source, output, report);
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
