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
        List<Issue> issues,
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

    /**
     * 文档内容结构中需要用户明确确认后才能修正的问题。
     */
    public record Issue(
            String key,
            String type,
            int paragraphIndex,
            String originalText,
            String suggestedText,
            String reason,
            int confidence,
            boolean selected
    ) {
    }
}
