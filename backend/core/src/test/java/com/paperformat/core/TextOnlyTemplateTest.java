package com.paperformat.core;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.TextUtils;
import org.docx4j.XmlUtils;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Style;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class TextOnlyTemplateTest {
    @TempDir
    Path temporaryDirectory;

    @Test
    void recognizesTextOnlyRequirementsAndAnalyzesTheSourceDocumentFirst() throws Exception {
        Path templatePath = temporaryDirectory.resolve("requirements.docx");
        WordprocessingMLPackage template = WordprocessingMLPackage.createPackage();
        template.getMainDocumentPart().addParagraphOfText("正文使用小四号宋体，固定值18磅行距，首行缩进2个字符，两端对齐。");
        template.getMainDocumentPart().addParagraphOfText("一级标题使用三号黑体，加粗，居中。");
        template.getMainDocumentPart().addParagraphOfText("表格内文字使用五号宋体，固定值18磅行距，居中。");
        template.save(templatePath.toFile());

        Path sourcePath = temporaryDirectory.resolve("source.docx");
        WordprocessingMLPackage source = WordprocessingMLPackage.createPackage();
        for (String paragraph : new String[]{
                "毕业设计说明书", "招聘管理系统", "招聘管理系统",
                "摘要：本文说明系统的设计与实现。", "关键词：招聘；管理系统",
                "Recruitment Management System", "Abstract: This paper describes the system.",
                "Keywords: recruitment; management", "目录", "1 绪论 1", "招聘管理系统",
                "1 绪论", "这是待修改文档中的正文段落，用于验证系统先识别源文档结构。",
                "参考文献", "[1]AlamM,KhanS.AMicroservicesBasedSystem[J].2024.",
                "致谢", "感谢指导教师。", "附录", "附录内容"
        }) {
            source.getMainDocumentPart().addParagraphOfText(paragraph);
        }
        source.save(sourcePath.toFile());

        FormatPlan plan = new WordFormatProcessor().analyze(templatePath, sourcePath);

        assertEquals("TEXT_INSTRUCTIONS_ONLY", plan.templateRules().mode());
        assertTrue(plan.templateRules().modeDescription().contains("先识别待修改文档结构"));
        assertEquals("宋体", plan.templateRules().styles().get("body").font());
        assertTrue(plan.templateRules().styles().get("body").paragraphFormatting().contains("18 pt"));
        assertTrue(plan.documentSummary().detectedRoles().getOrDefault("heading-1", 0) >= 1);
        assertTrue(plan.documentSummary().detectedRoles().getOrDefault("body", 0) >= 1);
        assertTrue(plan.rules().stream().anyMatch(rule -> rule.key().equals("tables")));

        Path outputPath = temporaryDirectory.resolve("formatted.docx");
        Path reportPath = temporaryDirectory.resolve("report.json");
        ProcessingReport report = new WordFormatProcessor().process(
                templatePath, sourcePath, outputPath, reportPath, null, Set.of());
        assertEquals("TEXT_INSTRUCTIONS_ONLY", report.templateRules().mode());
        assertTrue(report.verification().outputReloadedWithDocx4j());

        WordprocessingMLPackage formatted = WordprocessingMLPackage.load(outputPath.toFile());
        P bodyParagraph = formatted.getMainDocumentPart().getContent().stream()
                .map(XmlUtils::unwrap)
                .filter(P.class::isInstance)
                .map(P.class::cast)
                .filter(paragraph -> TextUtils.getText(paragraph).contains("用于验证系统先识别源文档结构"))
                .findFirst()
                .orElseThrow();
        String styleId = bodyParagraph.getPPr().getPStyle().getVal();
        Style bodyStyle = formatted.getMainDocumentPart().getStyleDefinitionsPart().getJaxbElement().getStyle().stream()
                .filter(style -> styleId.equals(style.getStyleId()))
                .findFirst()
                .orElseThrow();
        RPr effectiveRunProperties = bodyStyle.getRPr();
        PPr effectiveParagraphProperties = bodyStyle.getPPr();
        assertEquals("宋体", effectiveRunProperties.getRFonts().getEastAsia(),
                () -> "Unexpected effective font for style " + styleId);
        assertEquals(24, effectiveRunProperties.getSz().getVal().intValue());
        assertEquals(360, effectiveParagraphProperties.getSpacing().getLine().intValue());
        assertEquals(200, effectiveParagraphProperties.getInd().getFirstLineChars().intValue());
    }
}
