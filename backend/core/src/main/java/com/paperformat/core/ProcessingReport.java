package com.paperformat.core;

import java.util.List;
import java.util.Map;

/**
 * 单次格式处理的结构化报告。
 *
 * <p>该 record 会被 Jackson 序列化为 JSON，供前端展示和用户下载。</p>
 */
public record ProcessingReport(
        String tool,
        String generatedAt,
        InputFile template,
        InputFile source,
        OutputFile output,
        TemplateRules templateRules,
        DocumentSummary documentSummary,
        SectionSummary sectionSummary,
        List<Modification> modifications,
        List<String> preservedFeatures,
        List<String> limitations,
        List<String> manualReviewItems,
        Verification verification
) {
    /** 输入文件快照，用于证明处理前后的原文件哈希未变化。 */
    public record InputFile(String path, long bytes, String sha256) {
    }

    /** 输出文件快照，用于下载核对和问题追踪。 */
    public record OutputFile(String path, long bytes, String sha256) {
    }

    /** 从模板中抽象出的页面和样式规则说明。 */
    public record TemplateRules(
            String mode,
            String modeDescription,
            String pageSize,
            Map<String, Double> marginsCm,
            Map<String, Double> headerFooterDistanceCm,
            Map<String, StyleRule> styles,
            List<String> extractedInstructions
    ) {
    }

    /** 单个模板样式的人类可读说明。 */
    public record StyleRule(
            String role,
            String styleName,
            String font,
            String size,
            String paragraphFormatting,
            String sourceType,
            String evidence
    ) {
    }

    /** 源文档结构统计，前端结果页会读取这些字段。 */
    public record DocumentSummary(
            int bodyParagraphs,
            int tables,
            int sections,
            int images,
            Map<String, Integer> detectedRoles
    ) {
    }

    /** 分节页面设置的检查和同步结果。 */
    public record SectionSummary(
            int sectionsInspected,
            int sectionsAdjusted,
            String action
    ) {
    }

    /** 单个段落样式修改记录。 */
    public record Modification(
            int bodyParagraphIndex,
            String role,
            String beforeStyle,
            String afterStyle,
            String textPreview
    ) {
    }

    /** 输出前后的安全校验结果。 */
    public record Verification(
            boolean originalTemplateUnchanged,
            boolean originalSourceUnchanged,
            boolean outputReloadedWithDocx4j,
            boolean zipPackageReadable,
            int outputBodyParagraphs,
            int outputTables,
            int outputSections
    ) {
    }
}
