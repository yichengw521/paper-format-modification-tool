package com.paperformat.core;

import java.util.List;

/**
 * 上传后、真正修改前展示给用户确认的格式计划。
 */
public record FormatPlan(
        String generatedAt,
        ProcessingReport.InputFile template,
        ProcessingReport.InputFile source,
        ProcessingReport.TemplateRules templateRules,
        ProcessingReport.DocumentSummary documentSummary,
        List<Rule> rules,
        List<String> warnings
) {
    public record Rule(
            String key,
            String group,
            String label,
            boolean enabled,
            String font,
            String size,
            String paragraphFormatting,
            String sourceType,
            String evidence,
            int confidence
    ) {
    }
}
