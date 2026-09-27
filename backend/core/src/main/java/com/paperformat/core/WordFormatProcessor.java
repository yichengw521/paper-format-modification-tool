package com.paperformat.core;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import org.docx4j.TextUtils;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Body;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Styles;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tr;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.time.OffsetDateTime;
import java.time.ZoneOffset;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipFile;

/**
 * DOCX 格式处理核心。
 *
 * <p>处理流程是：读取模板和源文档，复制模板中受管样式，识别论文结构，再把源文档中
 * 可明确分类的段落切换到模板样式。这里尽量只改格式，不改正文内容。</p>
 */
public final class WordFormatProcessor {
    // 正文标题和题注的识别规则。它们只用于结构分类，真正的格式来自模板里的段落样式。
    private static final Pattern LEVEL_4 = Pattern.compile("^\\d+\\.\\d+\\.\\d+\\.\\d+\\s*\\S.*$");
    private static final Pattern LEVEL_3 = Pattern.compile("^\\d+\\.\\d+\\.\\d+\\s*\\S.*$");
    private static final Pattern LEVEL_2 = Pattern.compile("^\\d+\\.\\d+\\s*\\S.*$");
    private static final Pattern LEVEL_1 = Pattern.compile("^\\d+\\s*[^\\d.\\s].*$");
    private static final Pattern TABLE_CAPTION = Pattern.compile("^表\\s*\\d+[-－]\\d+.*$");
    private static final Pattern FIGURE_CAPTION = Pattern.compile("^图\\s*\\d+[-－]\\d+.*$");
    private static final Pattern CONTINUED_TABLE = Pattern.compile("^续表\\s*\\d+[-－]\\d+.*$");
    private static final Pattern HEADING_PREFIX = Pattern.compile("^(\\d+(?:\\.\\d+){0,3})\\s*(\\S.*)$");
    private static final Pattern COVER_DATE = Pattern.compile("^(\\d{4})年(\\d{1,2})月(\\d{1,2})日$");
    private static final Pattern PAGE_DISTANCE = Pattern.compile(
            "(上|下|左|右|页眉|页脚)(?:页边距|边距|距边界|距离)?\\s*(?:为|是|[:：])?\\s*(\\d+(?:\\.\\d+)?)\\s*(?:厘米|cm)",
            Pattern.CASE_INSENSITIVE);

    // 工具会管理这些模板样式。缺少任意一个样式时，说明模板不符合当前处理器的假设。
    private static final Set<String> TEMPLATE_STYLE_NAMES = Set.of(
            "论文一级标题", "论文二级标题", "论文三级标题", "论文四级标题", "论文正文",
            "论文图号图名", "论文表格序号及题目", "论文续表", "论文表格后面段落正文", "论文参考文献",
            "论文封面课题名称", "论文摘要中课题名称", "论文摘要正文", "论文表格正文", "参考文献及致谢",
            "toc 1", "toc 2"
    );

    /**
     * 根据模板修正源文档格式，并输出新的 DOCX 和 JSON 报告。
     *
     * @param templatePath 学校格式模板
     * @param sourcePath 待修改的论文文档
     * @param outputPath 输出文档路径，不能与输入文件相同
     * @param reportPath JSON 报告路径
     */
    public ProcessingReport process(Path templatePath, Path sourcePath, Path outputPath, Path reportPath) throws Exception {
        return process(templatePath, sourcePath, outputPath, reportPath, null);
    }

    /**
     * 只分析、不修改，用于生成用户二次确认页面。
     */
    public FormatPlan analyze(Path templatePath, Path sourcePath) throws Exception {
        requireDocx(templatePath, "template");
        requireDocx(sourcePath, "source");
        WordprocessingMLPackage template = WordprocessingMLPackage.load(templatePath.toFile());
        WordprocessingMLPackage source = WordprocessingMLPackage.load(sourcePath.toFile());
        String templateMode = TemplateRuleExtractor.detectMode(template);
        TemplateRuleExtractor.ensureOperationalStyles(template, source);
        ProcessingReport.TemplateRules templateRules = TemplateRuleExtractor.extract(template, templateMode);
        List<P> paragraphs = directBodyParagraphs(source.getMainDocumentPart());
        DocumentStructure structure = analyzeStructure(source, paragraphs);
        Map<String, Integer> detected = new LinkedHashMap<>();
        for (int index = 0; index < paragraphs.size(); index++) {
            String text = normalizeText(TextUtils.getText(paragraphs.get(index)));
            if (!text.isBlank()) {
                detected.merge(classify(text, index, structure,
                        styleNameById(source, getStyleId(paragraphs.get(index)))), 1, Integer::sum);
            }
        }
        ProcessingReport.DocumentSummary summary = new ProcessingReport.DocumentSummary(
                paragraphs.size(), allTables(source.getMainDocumentPart()).size(),
                allSections(source.getMainDocumentPart()).size(), countImageRelationships(source.getMainDocumentPart()), detected);
        List<FormatPlan.Rule> rules = confirmationRules(templateRules);
        List<FormatPlan.Issue> issues = detectStructureIssues(template, source, paragraphs, structure);
        List<String> warnings = new ArrayList<>();
        warnings.add(templateRules.modeDescription());
        if ("TEXT_INSTRUCTIONS_ONLY".equals(templateRules.mode())) {
            warnings.add("文字要求未明确的页面或格式属性会保留待修改文档原值，不会采用要求文档自身的默认版式。");
        }
        warnings.add("封面采用模板中的同类封面段落和信息表格式，同时保留用户填写的文字。");
        warnings.add("目录保留为 Word 自动目录，并在本机装有 Microsoft Word 时刷新页码和点引导符。");
        warnings.add("公式、浮动文本框和无法可靠匹配的复杂对象只保留，不自动重建。");
        if (!issues.isEmpty()) {
            warnings.add("检测到 " + issues.size() + " 个章节编号或标题文字问题；只有在确认页勾选后才会修改文字。");
        }
        return new FormatPlan(
                OffsetDateTime.now(ZoneOffset.UTC).toString(),
                new ProcessingReport.InputFile(templatePath.toString(), Files.size(templatePath), sha256(templatePath)),
                new ProcessingReport.InputFile(sourcePath.toString(), Files.size(sourcePath), sha256(sourcePath)),
                templateRules,
                summary,
                rules,
                issues,
                warnings
        );
    }

    /**
     * 按用户确认后的规则执行修改。enabledRuleKeys 为 null 时表示全部启用，兼容命令行 Demo。
     */
    public ProcessingReport process(
            Path templatePath,
            Path sourcePath,
            Path outputPath,
            Path reportPath,
            Set<String> enabledRuleKeys
    ) throws Exception {
        return process(templatePath, sourcePath, outputPath, reportPath, enabledRuleKeys, Set.of());
    }

    /**
     * 按确认后的格式规则和内容修正项执行处理。
     */
    public ProcessingReport process(
            Path templatePath,
            Path sourcePath,
            Path outputPath,
            Path reportPath,
            Set<String> enabledRuleKeys,
            Set<String> acceptedIssueKeys
    ) throws Exception {
        requireDocx(templatePath, "template");
        requireDocx(sourcePath, "source");
        if (templatePath.equals(outputPath)
                || sourcePath.equals(outputPath)
                || (Files.exists(outputPath)
                && (Files.isSameFile(templatePath, outputPath) || Files.isSameFile(sourcePath, outputPath)))) {
            throw new IllegalArgumentException("Output must be different from both input files.");
        }

        Files.createDirectories(outputPath.getParent());
        Files.createDirectories(reportPath.getParent());

        String templateHashBefore = sha256(templatePath);
        String sourceHashBefore = sha256(sourcePath);
        WordprocessingMLPackage template = WordprocessingMLPackage.load(templatePath.toFile());
        WordprocessingMLPackage target = WordprocessingMLPackage.load(sourcePath.toFile());

        // 先从模板命名样式、示例段落和文字说明中生成可执行样式，再同步到目标文档。
        String templateMode = TemplateRuleExtractor.detectMode(template);
        TemplateRuleExtractor.ensureOperationalStyles(template, target);
        P referenceParagraphTemplate = findReferenceParagraphTemplate(template);
        P figureCaptionTemplate = findMatchingParagraph(template, FIGURE_CAPTION);
        P chineseAbstractTemplate = "TEXT_INSTRUCTIONS_ONLY".equals(templateMode)
                ? null : findPrefixParagraph(template, "摘要", "摘 要");
        P englishAbstractTemplate = "TEXT_INSTRUCTIONS_ONLY".equals(templateMode)
                ? null : findPrefixParagraph(template, "Abstract");
        P tocLevel1Template = findParagraphByStyleName(template, "toc 1");
        P tocLevel2Template = findParagraphByStyleName(template, "toc 2");
        P tableCellTemplate = findTemplateTableCellParagraph(template);
        P figureContainerTemplate = findTemplateFigureContainerParagraph(template);
        Map<String, String> templateStyleIds = styleIdsByName(template);
        ensureRequiredStyles(templateStyleIds);

        boolean stylesChanged = synchronizeTemplateStyles(template, target);
        Map<String, String> targetStyleIds = styleIdsByName(target);
        ensureRequiredStyles(targetStyleIds);
        SectionResult sectionResult = enabled(enabledRuleKeys, "page")
                ? synchronizeSectionLayout(template, target, templateMode,
                        TemplateRuleExtractor.extract(template, templateMode).extractedInstructions())
                : new SectionResult(0, 0);

        MainDocumentPart targetMain = target.getMainDocumentPart();
        List<P> bodyParagraphs = directBodyParagraphs(targetMain);
        DocumentStructure structure = analyzeStructure(target, bodyParagraphs);
        List<FormatPlan.Issue> structureIssues = detectStructureIssues(template, target, bodyParagraphs, structure);
        List<ProcessingReport.Modification> modifications = new ArrayList<>();
        applyConfirmedStructureFixes(bodyParagraphs, structureIssues, acceptedIssueKeys, modifications);
        // 内容确认项可能把普通段落改成真正的标题，必须重新计算区域和角色。
        structure = analyzeStructure(target, bodyParagraphs);

        if (enabled(enabledRuleKeys, "cover")) {
            applyCoverFormatting(template, target, structure);
        }

        // LinkedHashMap 保持角色出现顺序，报告阅读时更接近文档顺序。
        Map<String, Integer> detectedRoles = new LinkedHashMap<>();

        for (int i = 0; i < bodyParagraphs.size(); i++) {
            P paragraph = bodyParagraphs.get(i);
            String text = normalizeText(TextUtils.getText(paragraph));
            if (text.isBlank()) {
                continue;
            }

            String beforeStyleId = getStyleId(paragraph);
            String beforeStyleName = styleNameById(target, beforeStyleId);
            String role = classify(text, i, structure, beforeStyleName);
            detectedRoles.merge(role, 1, Integer::sum);
            if (!enabled(enabledRuleKeys, ruleKeyForRole(role))) {
                continue;
            }
            if (role.equals("toc-title")) {
                // 目录标题本身不应该作为一级目录项出现在自动目录里。
                excludeTocTitleFromGeneratedContents(paragraph);
            }
            String desiredStyleName = styleForRole(role);
            if (desiredStyleName == null) {
                continue;
            }

            String desiredStyleId = targetStyleIds.get(desiredStyleName);
            if (desiredStyleId == null) {
                continue;
            }

            boolean shouldApply = shouldApplyStyle(role, beforeStyleName);
            boolean sameStyle = Objects.equals(beforeStyleId, desiredStyleId);
            boolean normalizeOverrides = shouldNormalizeDirectOverrides(role);
            if (!shouldApply || (sameStyle && !normalizeOverrides)) {
                continue;
            }

            applyParagraphStyle(paragraph, desiredStyleId);
            applyRoleFormatting(paragraph, role);
            if (role.equals("keywords-en")) {
                P sample = findPrefixParagraph(template, "Key words", "Keywords");
                if (sample != null) applyParagraphPropertiesFromSample(sample, paragraph);
            }
            if (role.equals("abstract-zh") && compact(text).startsWith("摘要") && chineseAbstractTemplate != null) {
                applyParagraphPropertiesFromSample(chineseAbstractTemplate, paragraph);
                setFirstLineCharacters(paragraph, 0);
            }
            if (role.equals("abstract-en") && compact(text).toLowerCase(Locale.ROOT).startsWith("abstract")
                    && englishAbstractTemplate != null) {
                applyParagraphPropertiesFromSample(englishAbstractTemplate, paragraph);
                setFirstLineCharacters(paragraph, 0);
            }
            if (role.equals("toc-level-1") && tocLevel1Template != null) {
                applyParagraphPropertiesFromSample(tocLevel1Template, paragraph);
            }
            if (role.equals("toc-level-2") && tocLevel2Template != null) {
                applyParagraphPropertiesFromSample(tocLevel2Template, paragraph);
            }
            if (role.equals("reference-item") && referenceParagraphTemplate != null) {
                applyParagraphPropertiesFromSample(referenceParagraphTemplate, paragraph);
            }
            if (role.equals("figure-caption") && figureCaptionTemplate != null) {
                // 图片题注通常包含命名样式之外的居中、段前/段后等直接段落属性。
                applyParagraphPropertiesFromSample(figureCaptionTemplate, paragraph);
                applyRunPropertiesFromSample(figureCaptionTemplate, paragraph);
            }
            if (role.equals("toc-title") || role.equals("section-title")) {
                applyTemplateDisplayText(template, paragraph, text);
            }
            if (role.equals("figure-caption")) {
                keepFigureWithCaption(targetMain, paragraph);
            } else if (role.equals("table-caption") || role.equals("continued-table-caption")) {
                setKeepNext(paragraph);
            }
            if (role.equals("reference-item")) {
                String normalizedReference = normalizeReferenceText(text);
                if (!normalizedReference.equals(text)) {
                    replaceParagraphText(paragraph, normalizedReference);
                    text = normalizedReference;
                }
            }
            modifications.add(new ProcessingReport.Modification(
                    i,
                    role,
                    sameStyle
                            ? desiredStyleName + "（清除直接格式覆盖）"
                            : (beforeStyleName == null ? beforeStyleId : beforeStyleName),
                    desiredStyleName,
                    preview(text)
            ));
        }

        if (enabled(enabledRuleKeys, "captions")) {
            String figureCaptionStyleId = targetStyleIds.get(styleForRole("figure-caption"));
            formatFigureCaptionsInsideTables(
                    targetMain,
                    figureCaptionTemplate,
                    figureCaptionStyleId,
                    detectedRoles,
                    modifications
            );
            formatFigureContainers(targetMain, figureContainerTemplate, modifications);
        }

        if (enabled(enabledRuleKeys, "tables")) {
            String tableCellStyleId = targetStyleIds.get("论文表格正文");
            formatBodyTables(targetMain, tableCellTemplate, tableCellStyleId, modifications);
        }

        requestWordFieldRefresh(target);

        target.save(outputPath.toFile());
        boolean wordFieldsRefreshed = refreshFieldsWithMicrosoftWord(outputPath);
        boolean zipReadable = isReadableZip(outputPath);
        WordprocessingMLPackage reloaded = WordprocessingMLPackage.load(outputPath.toFile());

        String templateHashAfter = sha256(templatePath);
        String sourceHashAfter = sha256(sourcePath);
        String outputHash = sha256(outputPath);

        int tables = allTables(targetMain).size();
        int images = countImageRelationships(targetMain);
        int sections = allSections(targetMain).size();

        ProcessingReport result = new ProcessingReport(
                "Paper Format Modification Tool Core 0.8.0",
                OffsetDateTime.now(ZoneOffset.UTC).toString(),
                new ProcessingReport.InputFile(templatePath.toString(), Files.size(templatePath), templateHashBefore),
                new ProcessingReport.InputFile(sourcePath.toString(), Files.size(sourcePath), sourceHashBefore),
                new ProcessingReport.OutputFile(outputPath.toString(), Files.size(outputPath), outputHash),
                TemplateRuleExtractor.extract(template, templateMode),
                new ProcessingReport.DocumentSummary(bodyParagraphs.size(), tables, sections, images, detectedRoles),
                new ProcessingReport.SectionSummary(
                        sectionResult.inspected(),
                        sectionResult.adjusted(),
                        "TEXT_INSTRUCTIONS_ONLY".equals(templateMode)
                                ? "仅应用文字要求中明确给出的纸张和边距参数；未说明的页面属性及原页眉页脚保持不变。"
                                : "同步模板的页面尺寸、页边距、分栏与文档网格；按摘要、目录和正文分节连续计算页码，保留原页眉页脚。"
                ),
                modifications,
                List.of(
                        "保留原说明书中的全部文字、表格、图片、超链接、书签和各分节页眉页脚。",
                        "先识别封面、中英文摘要、目录、正文、参考文献、致谢和附录，再按各区域分别应用模板样式。",
                        "封面说明文字、课题名称和封面信息表已按模板对应组件复制格式，同时保留原文内容。",
                        "中文/英文摘要与关键词分别处理，标签和正文保留不同字体、字号与加粗规则。",
                        "正文表格单元格、图片布局表格和目录条目的直接段落格式已纳入处理。",
                        wordFieldsRefreshed
                                ? "已调用本机 Microsoft Word 刷新自动目录、页码和域结果。"
                                : "已标记 Word 在打开文档时刷新自动目录和域结果。",
                        stylesChanged
                                ? "已将受管段落样式定义同步为模板版本，并补充缺少的模板样式。"
                                : "原说明书的受管段落样式定义已与模板一致。"
                ),
                List.of(
                        "目录保留为 Word 的 TOC 域；若本机没有可用的 Microsoft Word，需要打开文件后右键目录并选择“更新整个目录”。",
                        "公式、浮动文本框、复杂编号和交叉引用会被保留，不会重新构建。",
                        "涉及论文内容含义的问题只会列入人工复核，不会自动改写。"
                ),
                manualReviewItems(structureIssues, acceptedIssueKeys),
                new ProcessingReport.Verification(
                        templateHashBefore.equals(templateHashAfter),
                        sourceHashBefore.equals(sourceHashAfter),
                        reloaded != null,
                        zipReadable,
                        directBodyParagraphs(reloaded.getMainDocumentPart()).size(),
                        allTables(reloaded.getMainDocumentPart()).size(),
                        allSections(reloaded.getMainDocumentPart()).size()
                )
        );

        ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
        mapper.writeValue(reportPath.toFile(), result);
        return result;
    }

    private static List<FormatPlan.Rule> confirmationRules(ProcessingReport.TemplateRules rules) {
        Map<String, ProcessingReport.StyleRule> styles = rules.styles();
        List<FormatPlan.Rule> result = new ArrayList<>();
        boolean textOnly = "TEXT_INSTRUCTIONS_ONLY".equals(rules.mode());
        result.add(new FormatPlan.Rule("page", "页面", "页面尺寸、页边距和分节版式", true,
                "—", rules.pageSize(), "页边距 " + rules.marginsCm(), "SECTION_PROPERTIES",
                textOnly ? "仅应用文字中明确写出的页面参数，其余保留原文档" : "模板页面设置", textOnly ? 92 : 100));
        result.add(new FormatPlan.Rule("cover", "封面", "封面全部格式", true,
                "按模板逐组件匹配", "校名、说明书标题、课题名称和信息表分别匹配",
                "复制对应段落、行距、对齐方式、表格行列与单元格格式；保留用户文字",
                textOnly ? "SOURCE_STRUCTURE+TEXT_INSTRUCTION" : "MATCHED_COMPONENT",
                textOnly ? "从待修改文档识别封面组件，只应用文字要求中明确的属性" : "模板封面中的同位置段落和封面信息表",
                textOnly ? 88 : 96));
        ProcessingReport.StyleRule abstractRule = styles.get("abstract");
        result.add(new FormatPlan.Rule("abstract-zh", "摘要", "中文题目、摘要与关键词", true,
                textOnly && abstractRule != null ? abstractRule.font() : "题目/标签黑体，正文宋体；西文 Times New Roman",
                textOnly && abstractRule != null ? abstractRule.size() : "题目 16 pt；标签 14 pt；正文 12 pt",
                textOnly && abstractRule != null ? abstractRule.paragraphFormatting()
                        : "题目居中；摘要标签左顶格；后续正文首行缩进 2 字符；关键词不缩进",
                textOnly ? "SOURCE_STRUCTURE+TEXT_INSTRUCTION" : "TEXT_INSTRUCTION+MATCHED_SAMPLE",
                textOnly && abstractRule != null ? abstractRule.evidence() : "模板摘要样例及其下方格式说明", textOnly ? 90 : 96));
        result.add(new FormatPlan.Rule("abstract-en", "摘要", "英文题目、Abstract 与 Keywords", true,
                textOnly && abstractRule != null ? abstractRule.font() : "Times New Roman",
                textOnly && abstractRule != null ? abstractRule.size() : "题目 16 pt；标签 14 pt；正文 12 pt",
                textOnly && abstractRule != null ? abstractRule.paragraphFormatting()
                        : "题目居中；Abstract 标签左顶格；后续正文首行缩进 2 字符；Keywords 不缩进",
                textOnly ? "SOURCE_STRUCTURE+TEXT_INSTRUCTION" : "MATCHED_SAMPLE",
                textOnly && abstractRule != null ? abstractRule.evidence() : "模板英文摘要样例；标签和正文按字符范围分别处理",
                textOnly ? 90 : 95));
        result.add(new FormatPlan.Rule("toc", "目录", "Word 自动目录、点引导符与页码", true,
                "目录标题黑体，条目宋体/Times New Roman", "标题 16 pt；条目 10.5 pt",
                "保留 TOC 域，只显示一、二级标题；点引导符；页码右对齐；目录标题不进入目录",
                "WORD_FIELD+NAMED_STYLE", "模板 TOC 域、TOC1/TOC2 样式及目录文字要求", 98));
        addRule(result, styles, "headings", "正文", "一至四级标题", "heading1", 96);
        addRule(result, styles, "body", "正文", "正文段落", "body", 98);
        addRule(result, styles, "captions", "图表", "图题、表题、续表与图片布局", "tableCaption", 93);
        addRule(result, styles, "tables", "图表", "正文表格单元格", "tableCell", 94);
        addRule(result, styles, "references", "后置部分", "参考文献标题、条目格式与安全空格规范", "references", 96);
        addRule(result, styles, "thanks", "后置部分", "致谢标题与正文", "thanksBody", 95);
        return result;
    }

    private static void addRule(
            List<FormatPlan.Rule> target,
            Map<String, ProcessingReport.StyleRule> styles,
            String key,
            String group,
            String label,
            String styleKey,
            int confidence
    ) {
        ProcessingReport.StyleRule style = styles.get(styleKey);
        target.add(new FormatPlan.Rule(
                key, group, label, true,
                style == null ? "按模板对应示例" : style.font(),
                style == null ? "按模板对应示例" : style.size(),
                style == null ? "按模板对应示例" : style.paragraphFormatting(),
                style == null ? "MATCHED_COMPONENT" : style.sourceType(),
                style == null ? "模板中的同类组件或文字要求" : style.evidence(),
                confidence
        ));
    }

    private static boolean enabled(Set<String> enabledRuleKeys, String key) {
        return enabledRuleKeys == null || key == null || enabledRuleKeys.contains(key);
    }

    private static String ruleKeyForRole(String role) {
        return switch (role) {
            case "cover-title" -> "cover";
            case "thesis-title", "abstract-zh", "keywords-zh" -> "abstract-zh";
            case "abstract-en", "keywords-en" -> "abstract-en";
            case "toc-title", "toc-level-1", "toc-level-2" -> "toc";
            case "heading-1", "heading-2", "heading-3", "heading-4" -> "headings";
            case "figure-caption", "table-caption", "continued-table-caption", "table-following-body" -> "captions";
            case "table-cell" -> "tables";
            case "reference-item" -> "references";
            case "thanks-body" -> "thanks";
            case "section-title" -> "references";
            case "body", "appendix-body" -> "body";
            default -> null;
        };
    }

    /** Copy the complete cover component formatting while preserving target text. */
    private static void applyCoverFormatting(
            WordprocessingMLPackage template,
            WordprocessingMLPackage target,
            DocumentStructure targetStructure
    ) {
        List<P> templateParagraphs = directBodyParagraphs(template.getMainDocumentPart());
        int templateMarker = findFirstCompact(templateParagraphs, 0, templateParagraphs.size(), "毕业设计说明书");
        if (templateMarker >= 0 && targetStructure.coverMarker() >= 0) {
            copyParagraphFormatting(templateParagraphs.get(templateMarker),
                    directBodyParagraphs(target.getMainDocumentPart()).get(targetStructure.coverMarker()));
        }
        List<Tbl> templateTables = allTables(template.getMainDocumentPart());
        List<Tbl> targetTables = allTables(target.getMainDocumentPart());
        Tbl templateTable = findCoverTable(template.getMainDocumentPart());
        Tbl targetTable = findCoverTable(target.getMainDocumentPart());
        if (templateTable != null && targetTable != null) {
            int templateCoverIndex = templateTables.indexOf(templateTable);
            int targetCoverIndex = targetTables.indexOf(targetTable);
            int prefixCount = Math.min(templateCoverIndex, targetCoverIndex);
            // 封面顶部的校徽/校名字样经常放在一个无边框布局表格中；它和信息表都要复制格式。
            for (int i = 0; i <= prefixCount; i++) {
                copyTableFormatting(templateTables.get(i), targetTables.get(i));
            }
            if (templateCoverIndex != targetCoverIndex) {
                copyTableFormatting(templateTable, targetTable);
            }
        }
    }

    private static Tbl findCoverTable(MainDocumentPart main) {
        for (Tbl table : allTables(main)) {
            String text = compact(TextUtils.getText(table));
            if (text.contains("专业") && (text.contains("学生姓名") || text.contains("指导教师"))) {
                return table;
            }
        }
        return null;
    }

    private static void copyTableFormatting(Tbl from, Tbl to) {
        to.setTblPr(copy(from.getTblPr()));
        to.setTblGrid(copy(from.getTblGrid()));
        List<Tr> sourceRows = childRows(from);
        List<Tr> targetRows = childRows(to);
        for (int rowIndex = 0; rowIndex < Math.min(sourceRows.size(), targetRows.size()); rowIndex++) {
            Tr sourceRow = sourceRows.get(rowIndex);
            Tr targetRow = targetRows.get(rowIndex);
            targetRow.setTrPr(copy(sourceRow.getTrPr()));
            List<Tc> sourceCells = childCells(sourceRow);
            List<Tc> targetCells = childCells(targetRow);
            for (int cellIndex = 0; cellIndex < Math.min(sourceCells.size(), targetCells.size()); cellIndex++) {
                Tc sourceCell = sourceCells.get(cellIndex);
                Tc targetCell = targetCells.get(cellIndex);
                targetCell.setTcPr(copy(sourceCell.getTcPr()));
                List<P> sourceParagraphs = childParagraphs(sourceCell);
                List<P> targetParagraphs = childParagraphs(targetCell);
                for (int paragraphIndex = 0; paragraphIndex < Math.min(sourceParagraphs.size(), targetParagraphs.size()); paragraphIndex++) {
                    copyParagraphFormatting(sourceParagraphs.get(paragraphIndex), targetParagraphs.get(paragraphIndex));
                }
            }
        }
    }

    private static List<Tr> childRows(Tbl table) {
        List<Tr> result = new ArrayList<>();
        for (Object item : table.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (value instanceof Tr row) result.add(row);
        }
        return result;
    }

    private static List<Tc> childCells(Tr row) {
        List<Tc> result = new ArrayList<>();
        for (Object item : row.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (value instanceof Tc cell) result.add(cell);
        }
        return result;
    }

    private static List<P> childParagraphs(Tc cell) {
        List<P> result = new ArrayList<>();
        for (Object item : cell.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (value instanceof P paragraph) result.add(paragraph);
        }
        return result;
    }

    private static void copyParagraphFormatting(P from, P to) {
        SectPr preservedSection = to.getPPr() == null ? null : to.getPPr().getSectPr();
        PPr pPr = from.getPPr() == null ? new PPr() : copy(from.getPPr());
        pPr.setSectPr(preservedSection);
        to.setPPr(pPr);
        if (copySemanticDateRunFormatting(from, to)) {
            return;
        }
        List<RPr> patterns = new ArrayList<>();
        for (Object item : from.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (value instanceof R run && !runText(run).isBlank()) {
                patterns.add(run.getRPr() == null ? null : copy(run.getRPr()));
            }
        }
        int runIndex = 0;
        for (Object item : to.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (value instanceof R run && !runText(run).isBlank()) {
                RPr pattern = patterns.isEmpty() ? null : patterns.get(Math.min(runIndex, patterns.size() - 1));
                run.setRPr(pattern == null ? null : copy(pattern));
                runIndex++;
            }
        }
    }

    /**
     * 封面日期不能按 run 序号复制格式：日期位数变化后，红色等局部格式会落到“日”字上。
     * 这里按“年值、年、月值、月、日值、日”六个语义片段映射模板格式。
     */
    private static boolean copySemanticDateRunFormatting(P from, P to) {
        String sourceText = normalizeText(TextUtils.getText(from)).replace(" ", "");
        String targetText = normalizeText(TextUtils.getText(to)).replace(" ", "");
        Matcher sourceDate = COVER_DATE.matcher(sourceText);
        Matcher targetDate = COVER_DATE.matcher(targetText);
        if (!sourceDate.matches() || !targetDate.matches()) return false;

        List<RunSpan> sourceRuns = runSpans(from);
        List<RunSpan> targetRuns = runSpans(to);
        int[] sourcePositions = {
                sourceDate.start(1), sourceDate.end(1),
                sourceDate.start(2), sourceDate.end(2),
                sourceDate.start(3), sourceDate.end(3), sourceDate.end(3) + 1
        };
        for (RunSpan targetRun : targetRuns) {
            int semantic = dateSemanticIndex(targetDate, targetRun.start());
            int sourcePosition = sourcePositions[Math.min(semantic, sourcePositions.length - 1)];
            RPr pattern = runPropertiesAt(sourceRuns, sourcePosition);
            targetRun.run().setRPr(pattern == null ? null : copy(pattern));
        }
        return true;
    }

    private static int dateSemanticIndex(Matcher date, int position) {
        if (position < date.end(1)) return 0;       // 年份数字
        if (position < date.start(2)) return 1;     // 年
        if (position < date.end(2)) return 2;       // 月份数字
        if (position < date.start(3)) return 3;     // 月
        if (position < date.end(3)) return 4;       // 日期数字
        return 5;                                   // 日
    }

    private static List<RunSpan> runSpans(P paragraph) {
        List<RunSpan> result = new ArrayList<>();
        int offset = 0;
        for (Object item : paragraph.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (!(value instanceof R run)) continue;
            String text = runText(run).replace(" ", "");
            if (text.isEmpty()) continue;
            result.add(new RunSpan(run, offset, offset + text.length()));
            offset += text.length();
        }
        return result;
    }

    private static RPr runPropertiesAt(List<RunSpan> runs, int position) {
        for (RunSpan run : runs) {
            if (position >= run.start() && position < run.end()) return run.run().getRPr();
        }
        return runs.isEmpty() ? null : runs.get(runs.size() - 1).run().getRPr();
    }

    private static String runText(R run) {
        StringBuilder result = new StringBuilder();
        for (Object item : run.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (value instanceof Text text) result.append(text.getValue());
        }
        return result.toString();
    }

    private static P findReferenceParagraphTemplate(WordprocessingMLPackage template) {
        List<P> paragraphs = directBodyParagraphs(template.getMainDocumentPart());
        int title = findFirstCompact(paragraphs, 0, paragraphs.size(), "参考文献");
        int sample = nextNonEmpty(paragraphs, title + 1, paragraphs.size());
        return title >= 0 && sample >= 0 ? paragraphs.get(sample) : null;
    }

    private static P findPrefixParagraph(WordprocessingMLPackage document, String... prefixes) {
        List<P> paragraphs = directBodyParagraphs(document.getMainDocumentPart());
        int index = findPrefixCompact(paragraphs, 0, paragraphs.size(), prefixes);
        return index < 0 ? null : paragraphs.get(index);
    }

    private static P findMatchingParagraph(WordprocessingMLPackage document, Pattern pattern) {
        for (P paragraph : directBodyParagraphs(document.getMainDocumentPart())) {
            if (pattern.matcher(normalizeText(TextUtils.getText(paragraph))).matches()) {
                return paragraph;
            }
        }
        for (Tbl table : allTables(document.getMainDocumentPart())) {
            for (P paragraph : paragraphsInTable(table)) {
                if (pattern.matcher(normalizeText(TextUtils.getText(paragraph))).matches()) {
                    return paragraph;
                }
            }
        }
        return null;
    }

    private static P findParagraphByStyleName(
            WordprocessingMLPackage document,
            String expectedStyleName
    ) throws Docx4JException {
        for (P paragraph : directBodyParagraphs(document.getMainDocumentPart())) {
            String styleName = styleNameById(document, getStyleId(paragraph));
            if (expectedStyleName.equalsIgnoreCase(styleName)) return paragraph;
        }
        return null;
    }

    private static P findTemplateTableCellParagraph(WordprocessingMLPackage template) {
        Tbl cover = findCoverTable(template.getMainDocumentPart());
        for (Tbl table : allTables(template.getMainDocumentPart())) {
            if (table == cover || containsDrawing(table)) continue;
            List<Tr> rows = childRows(table);
            if (rows.size() < 2 || childCells(rows.get(0)).size() < 2) continue;
            for (P paragraph : paragraphsInTable(table)) {
                if (!normalizeText(TextUtils.getText(paragraph)).isBlank()) return paragraph;
            }
        }
        return null;
    }

    private static P findTemplateFigureContainerParagraph(WordprocessingMLPackage template) {
        List<Tbl> tables = allTables(template.getMainDocumentPart());
        Tbl cover = findCoverTable(template.getMainDocumentPart());
        int coverIndex = cover == null ? -1 : tables.indexOf(cover);
        for (int i = coverIndex + 1; i < tables.size(); i++) {
            for (P paragraph : paragraphsInTable(tables.get(i))) {
                if (containsDrawing(paragraph)) return paragraph;
            }
        }
        return null;
    }

    private static void formatFigureContainers(
            MainDocumentPart main,
            P sample,
            List<ProcessingReport.Modification> modifications
    ) {
        if (sample == null) return;
        List<Tbl> tables = allTables(main);
        Tbl cover = findCoverTable(main);
        int coverIndex = cover == null ? -1 : tables.indexOf(cover);
        int reportIndex = -10_000;
        for (int tableIndex = coverIndex + 1; tableIndex < tables.size(); tableIndex++) {
            Tbl table = tables.get(tableIndex);
            boolean changed = false;
            for (P paragraph : paragraphsInTable(table)) {
                if (!containsDrawing(paragraph)) continue;
                applyParagraphPropertiesFromSample(sample, paragraph);
                applyRunPropertiesFromSample(sample, paragraph);
                setKeepNext(paragraph);
                setKeepLines(paragraph);
                changed = true;
            }
            if (changed) {
                modifications.add(new ProcessingReport.Modification(
                        reportIndex--, "figure-container", "原图片容器段落格式",
                        "模板图片容器段落格式", "图片布局表格 " + (tableIndex + 1)));
            }
        }
    }

    private static void formatBodyTables(
            MainDocumentPart main,
            P sample,
            String styleId,
            List<ProcessingReport.Modification> modifications
    ) {
        List<Tbl> tables = allTables(main);
        Tbl cover = findCoverTable(main);
        int coverIndex = cover == null ? -1 : tables.indexOf(cover);
        int reportIndex = -20_000;
        for (int tableIndex = coverIndex + 1; tableIndex < tables.size(); tableIndex++) {
            Tbl table = tables.get(tableIndex);
            if (containsDrawing(table)) continue;
            List<Tr> rows = childRows(table);
            if (rows.isEmpty() || childCells(rows.get(0)).size() < 2) continue;

            int formattedParagraphs = 0;
            for (Tr row : rows) {
                for (Tc cell : childCells(row)) {
                    for (P paragraph : childParagraphs(cell)) {
                        if (normalizeText(TextUtils.getText(paragraph)).isBlank()) continue;
                        if (styleId != null) applyParagraphStyle(paragraph, styleId);
                        if (sample != null) {
                            applyParagraphPropertiesFromSample(sample, paragraph);
                            applyRunPropertiesFromSample(sample, paragraph);
                        }
                        formattedParagraphs++;
                    }
                }
            }
            if (formattedParagraphs > 0) {
                modifications.add(new ProcessingReport.Modification(
                        reportIndex--, "table-cell", "原表格单元格格式",
                        sample == null ? "文字要求生成的表格正文样式" : "模板表格样例格式",
                        "表格 " + (tableIndex + 1) + "，处理 " + formattedParagraphs + " 个段落"));
            }
        }
    }

    /**
     * Applies only high-confidence bibliography whitespace repairs.  It deliberately leaves names,
     * title wording and punctuation symbols intact; questionable bibliographic content still belongs
     * in manual review.
     */
    private static String normalizeReferenceText(String text) {
        String value = text.replace('\u00A0', ' ').replaceAll("\\s+", " ").trim();
        value = value.replaceFirst("^\\s*(\\[?\\d+\\]?[.]?)\\s*", "$1 ");

        int protectedStart = value.length();
        String lower = value.toLowerCase(Locale.ROOT);
        for (String marker : List.of("http://", "https://", "doi:")) {
            int index = lower.indexOf(marker);
            if (index >= 0) protectedStart = Math.min(protectedStart, index);
        }
        String editable = value.substring(0, protectedStart);
        String protectedTail = value.substring(protectedStart);
        if (editable.matches(".*[A-Za-z].*")) {
            editable = editable.replaceAll("(?<=[,;:])(?=[A-Za-z])", " ");
            editable = editable.replaceAll("(?<=[,;:])(?=[0-9])", " ");
            editable = editable.replaceAll("(?<=[.])(?=[A-Za-z])", " ");
            editable = editable.replaceAll("(?<=[a-z])(?=[A-Z])", " ");
            editable = editable.replaceAll("\\b([A-Z])(?=[A-Z][a-z])", "$1 ");
            editable = editable.replaceAll("(?<=[A-Z])(?=[A-Z][a-z])", " ");
            editable = editable.replaceAll("\\b(Boot|Java|Vue|React|System|Platform|Recruitment|Management)and\\b", "$1 and");
            editable = editable.replaceAll("\\b(js)(for|and|using)\\b", "$1 $2");
            // Restore common technology names that intentionally contain a lower-to-upper boundary.
            editable = editable.replace("My SQL", "MySQL").replace("My Batis", "MyBatis")
                    .replace("Vue. js", "Vue.js");
            editable = editable.replaceAll(" {2,}", " ");
        }
        return (editable + protectedTail).trim();
    }

    /** Keep the exact visible spacing used by short template titles such as “致    谢”. */
    private static void applyTemplateDisplayText(
            WordprocessingMLPackage template,
            P target,
            String targetText
    ) {
        String expectedCompact = compact(targetText);
        for (P candidate : directBodyParagraphs(template.getMainDocumentPart())) {
            String candidateText = displayText(TextUtils.getText(candidate));
            if (compact(candidateText).equals(expectedCompact) && !candidateText.equals(targetText)) {
                replaceParagraphText(target, candidateText);
                return;
            }
        }
    }

    private static void setKeepNext(P paragraph) {
        PPr pPr = paragraph.getPPr();
        if (pPr == null) {
            pPr = new PPr();
            paragraph.setPPr(pPr);
        }
        BooleanDefaultTrue keepNext = new BooleanDefaultTrue();
        keepNext.setVal(true);
        pPr.setKeepNext(keepNext);
    }

    private static void setKeepLines(P paragraph) {
        PPr pPr = paragraph.getPPr();
        if (pPr == null) {
            pPr = new PPr();
            paragraph.setPPr(pPr);
        }
        BooleanDefaultTrue keepLines = new BooleanDefaultTrue();
        keepLines.setVal(true);
        pPr.setKeepLines(keepLines);
    }

    /**
     * Keep an inline image, any blank spacer paragraphs immediately after it, and its caption
     * on the same page. Real documents often contain one blank paragraph between an image and
     * “图x-x ...”, so binding only the immediately preceding paragraph is insufficient.
     */
    private static void keepFigureWithCaption(MainDocumentPart main, P caption) {
        setKeepLines(caption);

        List<Object> bodyContent = main.getContent();
        int captionPosition = -1;
        for (int i = 0; i < bodyContent.size(); i++) {
            if (XmlUtils.unwrap(bodyContent.get(i)) == caption) {
                captionPosition = i;
                break;
            }
        }
        if (captionPosition < 0) return;

        int minimum = Math.max(0, captionPosition - 4);
        List<P> trailingBlankParagraphs = new ArrayList<>();
        for (int cursor = captionPosition - 1; cursor >= minimum; cursor--) {
            Object candidate = XmlUtils.unwrap(bodyContent.get(cursor));
            if (candidate instanceof P paragraph) {
                if (containsDrawing(paragraph)) {
                    setKeepNext(paragraph);
                    trailingBlankParagraphs.forEach(WordFormatProcessor::setKeepNext);
                    return;
                }
                if (normalizeText(TextUtils.getText(paragraph)).isBlank()) {
                    trailingBlankParagraphs.add(paragraph);
                    continue;
                }
                return;
            }

            if (candidate instanceof Tbl table && containsDrawing(table)) {
                List<P> tableParagraphs = paragraphsInTable(table);
                // 图片常被 Word 放进一行一列的无边框表格；把表格最后的段落链到题注。
                for (P tableParagraph : tableParagraphs) {
                    setKeepNext(tableParagraph);
                    setKeepLines(tableParagraph);
                }
                trailingBlankParagraphs.forEach(WordFormatProcessor::setKeepNext);
                return;
            }
            return;
        }
    }

    private static boolean containsDrawing(P paragraph) {
        return containsDrawing((Object) paragraph);
    }

    private static boolean containsDrawing(Object value) {
        String xml = XmlUtils.marshaltoString(value, true, true).toLowerCase(Locale.ROOT);
        return xml.contains(":drawing") || xml.contains(":pict") || xml.contains(":imagedata");
    }

    private static void formatFigureCaptionsInsideTables(
            MainDocumentPart main,
            P sample,
            String styleId,
            Map<String, Integer> detectedRoles,
            List<ProcessingReport.Modification> modifications
    ) {
        if (styleId == null) return;
        int nestedIndex = -1;
        for (Tbl table : allTables(main)) {
            List<P> paragraphs = paragraphsInTable(table);
            for (int i = 0; i < paragraphs.size(); i++) {
                P paragraph = paragraphs.get(i);
                String text = normalizeText(TextUtils.getText(paragraph));
                if (!FIGURE_CAPTION.matcher(text).matches()) continue;

                String beforeStyleId = getStyleId(paragraph);
                applyParagraphStyle(paragraph, styleId);
                applyRoleFormatting(paragraph, "figure-caption");
                if (sample != null) {
                    applyParagraphPropertiesFromSample(sample, paragraph);
                    applyRunPropertiesFromSample(sample, paragraph);
                }
                setKeepLines(paragraph);

                for (int previous = i - 1; previous >= Math.max(0, i - 4); previous--) {
                    P candidate = paragraphs.get(previous);
                    if (containsDrawing(candidate)) {
                        setKeepNext(candidate);
                        for (int spacer = previous + 1; spacer < i; spacer++) {
                            setKeepNext(paragraphs.get(spacer));
                        }
                        break;
                    }
                    if (!normalizeText(TextUtils.getText(candidate)).isBlank()) break;
                }

                detectedRoles.merge("figure-caption", 1, Integer::sum);
                modifications.add(new ProcessingReport.Modification(
                        nestedIndex--,
                        "figure-caption",
                        beforeStyleId,
                        "论文图号图名（表格内题注）",
                        preview(text)
                ));
            }
        }
    }

    private static List<P> paragraphsInTable(Tbl table) {
        List<P> result = new ArrayList<>();
        collectTableParagraphs(table, result);
        return result;
    }

    private static void collectTableParagraphs(Object node, List<P> result) {
        Object value = XmlUtils.unwrap(node);
        if (value instanceof P paragraph) {
            result.add(paragraph);
        } else if (value instanceof Tbl table) {
            table.getContent().forEach(item -> collectTableParagraphs(item, result));
        } else if (value instanceof Tr row) {
            row.getContent().forEach(item -> collectTableParagraphs(item, result));
        } else if (value instanceof Tc cell) {
            cell.getContent().forEach(item -> collectTableParagraphs(item, result));
        }
    }

    private static List<FormatPlan.Issue> detectStructureIssues(
            WordprocessingMLPackage template,
            WordprocessingMLPackage source,
            List<P> paragraphs,
            DocumentStructure structure
    ) throws Exception {
        Map<Integer, String> templateChapterTitles = new LinkedHashMap<>();
        for (P paragraph : directBodyParagraphs(template.getMainDocumentPart())) {
            String text = normalizeText(TextUtils.getText(paragraph));
            Matcher matcher = HEADING_PREFIX.matcher(text);
            if (!matcher.matches() || matcher.group(1).contains(".")) continue;
            if (!"论文一级标题".equals(styleNameById(template, getStyleId(paragraph)))) continue;
            templateChapterTitles.putIfAbsent(Integer.parseInt(matcher.group(1)), text);
        }

        Set<String> allPrefixes = new LinkedHashSet<>();
        for (int i = structure.mainStart(); i < structure.referencesStart(); i++) {
            Matcher matcher = HEADING_PREFIX.matcher(normalizeText(TextUtils.getText(paragraphs.get(i))));
            if (matcher.matches()) allPrefixes.add(matcher.group(1));
        }

        List<FormatPlan.Issue> result = new ArrayList<>();
        Set<String> seenPrefixes = new HashSet<>();
        Set<Integer> seenChapters = new HashSet<>();
        Set<Integer> reportedMissingChapters = new HashSet<>();
        for (int i = structure.mainStart(); i < structure.referencesStart(); i++) {
            String text = normalizeText(TextUtils.getText(paragraphs.get(i)));
            Matcher matcher = HEADING_PREFIX.matcher(text);
            if (!matcher.matches()) continue;
            String prefix = matcher.group(1);
            String[] parts = prefix.split("\\.");
            int chapter = Integer.parseInt(parts[0]);

            if (parts.length > 1 && !seenChapters.contains(chapter)
                    && reportedMissingChapters.add(chapter)
                    && templateChapterTitles.containsKey(chapter)) {
                int candidateIndex = previousNonEmpty(paragraphs, i - 1, structure.mainStart());
                if (candidateIndex >= structure.mainStart()) {
                    String candidate = normalizeText(TextUtils.getText(paragraphs.get(candidateIndex)));
                    if (!candidate.isBlank() && candidate.length() <= 30
                            && !HEADING_PREFIX.matcher(candidate).matches()
                            && !candidate.matches(".*[。；;，,：:]$")) {
                        String suggestion = templateChapterTitles.get(chapter);
                        result.add(new FormatPlan.Issue(
                                "missing-heading-" + candidateIndex,
                                "MISSING_HEADING_NUMBER",
                                candidateIndex,
                                candidate,
                                suggestion,
                                "后续出现“" + prefix + "”，但此前没有第 " + chapter
                                        + " 章一级标题；模板中的对应标题是“" + suggestion + "”。",
                                98,
                                true
                        ));
                        seenChapters.add(chapter);
                    }
                }
            }

            if (parts.length == 1) seenChapters.add(chapter);
            if (!seenPrefixes.add(prefix)) {
                String replacementPrefix = nextAvailablePrefix(prefix, allPrefixes);
                String suggestion = replacementPrefix + " " + matcher.group(2);
                result.add(new FormatPlan.Issue(
                        "duplicate-heading-" + i,
                        "DUPLICATE_HEADING_NUMBER",
                        i,
                        text,
                        suggestion,
                        "章节编号“" + prefix + "”在正文中重复，建议使用下一个未占用编号“"
                                + replacementPrefix + "”。",
                        96,
                        true
                ));
                allPrefixes.add(replacementPrefix);
            }
        }
        return result;
    }

    private static String nextAvailablePrefix(String prefix, Set<String> used) {
        String[] parts = prefix.split("\\.");
        int last = Integer.parseInt(parts[parts.length - 1]);
        String parent = parts.length == 1 ? "" : String.join(".", java.util.Arrays.copyOf(parts, parts.length - 1)) + ".";
        String candidate;
        do {
            candidate = parent + (++last);
        } while (used.contains(candidate));
        return candidate;
    }

    private static void applyConfirmedStructureFixes(
            List<P> paragraphs,
            List<FormatPlan.Issue> issues,
            Set<String> acceptedIssueKeys,
            List<ProcessingReport.Modification> modifications
    ) {
        Set<String> accepted = acceptedIssueKeys == null ? Set.of() : acceptedIssueKeys;
        for (FormatPlan.Issue issue : issues) {
            if (!accepted.contains(issue.key())) continue;
            if (issue.paragraphIndex() < 0 || issue.paragraphIndex() >= paragraphs.size()) continue;
            P paragraph = paragraphs.get(issue.paragraphIndex());
            String current = normalizeText(TextUtils.getText(paragraph));
            if (!current.equals(issue.originalText())) continue;
            replaceParagraphText(paragraph, issue.suggestedText());
            modifications.add(new ProcessingReport.Modification(
                    issue.paragraphIndex(), "confirmed-content-fix", issue.originalText(),
                    issue.suggestedText(), issue.suggestedText()));
            if ("DUPLICATE_HEADING_NUMBER".equals(issue.type())) {
                cascadeConfirmedHeadingPrefix(paragraphs, issue, modifications);
            }
        }
    }

    private static void cascadeConfirmedHeadingPrefix(
            List<P> paragraphs,
            FormatPlan.Issue issue,
            List<ProcessingReport.Modification> modifications
    ) {
        Matcher original = HEADING_PREFIX.matcher(issue.originalText());
        Matcher suggested = HEADING_PREFIX.matcher(issue.suggestedText());
        if (!original.matches() || !suggested.matches()) return;

        String oldPrefix = original.group(1);
        String newPrefix = suggested.group(1);
        int parentDepth = oldPrefix.split("\\.").length;
        for (int i = issue.paragraphIndex() + 1; i < paragraphs.size(); i++) {
            P paragraph = paragraphs.get(i);
            String current = normalizeText(TextUtils.getText(paragraph));
            Matcher child = HEADING_PREFIX.matcher(current);
            if (!child.matches()) continue;

            String childPrefix = child.group(1);
            int childDepth = childPrefix.split("\\.").length;
            if (childDepth <= parentDepth) break;
            if (!childPrefix.startsWith(oldPrefix + ".")) continue;

            String replacement = newPrefix + childPrefix.substring(oldPrefix.length())
                    + " " + child.group(2);
            replaceParagraphText(paragraph, replacement);
            modifications.add(new ProcessingReport.Modification(
                    i,
                    "confirmed-content-fix-child",
                    current,
                    replacement,
                    replacement
            ));
        }
    }

    private static List<String> manualReviewItems(
            List<FormatPlan.Issue> issues,
            Set<String> acceptedIssueKeys
    ) {
        Set<String> accepted = acceptedIssueKeys == null ? Set.of() : acceptedIssueKeys;
        List<String> result = new ArrayList<>();
        for (FormatPlan.Issue issue : issues) {
            if (!accepted.contains(issue.key())) {
                result.add("未采用结构修正：“" + issue.originalText() + "” → “" + issue.suggestedText() + "”。");
            }
        }
        result.add("英文摘要较长时，关键词可能自然延续到下一页；程序保持模板字号和行距，不压缩正文内容。");
        return result;
    }

    private static void replaceParagraphText(P paragraph, String replacement) {
        boolean written = false;
        for (Object item : paragraph.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (!(value instanceof R run)) continue;
            for (Object runItem : run.getContent()) {
                Object runValue = XmlUtils.unwrap(runItem);
                if (runValue instanceof Text text) {
                    text.setValue(written ? "" : replacement);
                    text.setSpace("preserve");
                    written = true;
                }
            }
        }
    }

    /**
     * Copy sample paragraph properties that live outside the named style (notably reference
     * hanging indents), while retaining the target style id and any section break.
     */
    private static void applyParagraphPropertiesFromSample(P sample, P target) {
        if (sample.getPPr() == null) return;
        PPr current = target.getPPr();
        PPrBase.PStyle targetStyle = current == null ? null : copy(current.getPStyle());
        SectPr targetSection = current == null ? null : current.getSectPr();
        PPr replacement = copy(sample.getPPr());
        replacement.setPStyle(targetStyle);
        replacement.setSectPr(targetSection);
        target.setPPr(replacement);
    }

    private static void applyRunPropertiesFromSample(P sample, P target) {
        List<RPr> patterns = new ArrayList<>();
        for (Object item : sample.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (value instanceof R run && !runText(run).isBlank()) {
                patterns.add(run.getRPr() == null ? null : copy(run.getRPr()));
            }
        }
        if (patterns.isEmpty()) return;
        int runIndex = 0;
        for (Object item : target.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (value instanceof R run && !runText(run).isBlank()) {
                RPr pattern = patterns.get(Math.min(runIndex, patterns.size() - 1));
                run.setRPr(pattern == null ? null : copy(pattern));
                runIndex++;
            }
        }
    }

    /**
     * 报告中展示的模板规则说明。
     *
     * <p>这些文字用于帮助用户理解处理依据，不直接驱动 docx4j 修改。</p>
     */
    /**
     * 将模板中的受管样式复制到目标文档。
     *
     * <p>如果目标文档已有同名样式，保留原 styleId，只替换样式定义。这样已有段落引用不会断。</p>
     */
    private static boolean synchronizeTemplateStyles(WordprocessingMLPackage template, WordprocessingMLPackage target) throws Exception {
        StyleDefinitionsPart templatePart = template.getMainDocumentPart().getStyleDefinitionsPart();
        StyleDefinitionsPart targetPart = target.getMainDocumentPart().getStyleDefinitionsPart();
        if (templatePart == null || targetPart == null) {
            throw new IllegalStateException("Both documents must contain a styles part.");
        }
        Styles templateStyles = templatePart.getJaxbElement();
        Styles targetStyles = targetPart.getJaxbElement();
        boolean changed = false;

        for (Style templateStyle : templateStyles.getStyle()) {
            String templateName = templateStyle.getName() == null ? null : templateStyle.getName().getVal();
            if (templateName == null || TEMPLATE_STYLE_NAMES.stream().noneMatch(name -> name.equalsIgnoreCase(templateName))) {
                continue;
            }
            boolean exists = false;
            for (int i = 0; i < targetStyles.getStyle().size(); i++) {
                Style targetStyle = targetStyles.getStyle().get(i);
                String targetName = targetStyle.getName() == null ? null : targetStyle.getName().getVal();
                if (templateName != null && targetName != null && templateName.equalsIgnoreCase(targetName)) {
                    exists = true;
                    // 深拷贝模板样式，避免同一个 JAXB 对象同时挂到两个文档包中。
                    Style replacement = (Style) XmlUtils.deepCopy(templateStyle);
                    replacement.setStyleId(targetStyle.getStyleId());
                    String before = XmlUtils.marshaltoString(targetStyle, true, true);
                    String after = XmlUtils.marshaltoString(replacement, true, true);
                    if (!before.equals(after)) {
                        targetStyles.getStyle().set(i, replacement);
                        changed = true;
                    }
                    break;
                }
            }

            if (!exists) {
                targetStyles.getStyle().add((Style) XmlUtils.deepCopy(templateStyle));
                changed = true;
            }
        }
        return changed;
    }

    /**
     * 同步页面尺寸、页边距、分栏和文档网格，并修正论文前置部分的连续页码。
     *
     * <p>页眉页脚关系仍保留目标文档；摘要从 I 开始、目录继续摘要页码、正文从 1 开始。</p>
     */
    private static SectionResult synchronizeSectionLayout(
            WordprocessingMLPackage template,
            WordprocessingMLPackage target,
            String templateMode,
            List<String> instructions
    ) throws Exception {
        if ("TEXT_INSTRUCTIONS_ONLY".equals(templateMode)) {
            return applyTextPageRequirements(target, instructions);
        }
        List<SectPr> templateSections = allSections(template.getMainDocumentPart());
        List<SectPr> targetSections = allSections(target.getMainDocumentPart());
        int inspected = Math.min(templateSections.size(), targetSections.size());
        int adjusted = 0;
        for (int i = 0; i < inspected; i++) {
            SectPr from = templateSections.get(i);
            SectPr to = targetSections.get(i);
            String before = sectionLayoutSignature(to);

            to.setPgSz(copy(from.getPgSz()));
            to.setPgMar(copy(from.getPgMar()));
            to.setCols(copy(from.getCols()));
            to.setDocGrid(copy(from.getDocGrid()));

            if (!before.equals(sectionLayoutSignature(to))) {
                adjusted++;
            }
        }
        if (targetSections.size() >= 4 && templateSections.size() >= 4) {
            // 第 2 节是中英文摘要，从罗马数字 I 开始。
            targetSections.get(1).setPgNumType(copy(templateSections.get(1).getPgNumType()));
            if (targetSections.get(1).getPgNumType() != null) {
                targetSections.get(1).getPgNumType().setStart(BigInteger.ONE);
            }
            // 第 3 节是目录，不能使用模板里固定的起始值；摘要篇幅变化时必须连续编号。
            targetSections.get(2).setPgNumType(copy(templateSections.get(2).getPgNumType()));
            if (targetSections.get(2).getPgNumType() != null) {
                targetSections.get(2).getPgNumType().setStart(null);
            }
            // 最后一节是正文，按模板从阿拉伯数字 1 重新开始。
            int lastTarget = targetSections.size() - 1;
            int lastTemplate = templateSections.size() - 1;
            targetSections.get(lastTarget).setPgNumType(copy(templateSections.get(lastTemplate).getPgNumType()));
            if (targetSections.get(lastTarget).getPgNumType() != null) {
                targetSections.get(lastTarget).getPgNumType().setStart(BigInteger.ONE);
            }
        }
        return new SectionResult(inspected, adjusted);
    }

    /**
     * A prose-only requirements file is not a layout template.  Only explicitly stated page
     * measurements are applied; unspecified values remain those of the user's source document.
     */
    private static SectionResult applyTextPageRequirements(
            WordprocessingMLPackage target,
            List<String> instructions
    ) {
        String text = String.join(" ", instructions).replace(" ", "");
        Map<String, BigInteger> distances = new HashMap<>();
        Matcher matcher = PAGE_DISTANCE.matcher(text);
        while (matcher.find()) {
            double cm = Double.parseDouble(matcher.group(2));
            distances.put(matcher.group(1), BigInteger.valueOf(Math.round(cm * 1440.0 / 2.54)));
        }
        boolean a4 = text.toUpperCase(Locale.ROOT).contains("A4");
        List<SectPr> sections = allSections(target.getMainDocumentPart());
        int adjusted = 0;
        for (SectPr section : sections) {
            String before = sectionLayoutSignature(section);
            if (a4) {
                SectPr.PgSz size = section.getPgSz() == null ? new SectPr.PgSz() : section.getPgSz();
                size.setW(BigInteger.valueOf(11906));
                size.setH(BigInteger.valueOf(16838));
                section.setPgSz(size);
            }
            if (!distances.isEmpty()) {
                SectPr.PgMar margins = section.getPgMar() == null ? new SectPr.PgMar() : section.getPgMar();
                if (distances.containsKey("上")) margins.setTop(distances.get("上"));
                if (distances.containsKey("下")) margins.setBottom(distances.get("下"));
                if (distances.containsKey("左")) margins.setLeft(distances.get("左"));
                if (distances.containsKey("右")) margins.setRight(distances.get("右"));
                if (distances.containsKey("页眉")) margins.setHeader(distances.get("页眉"));
                if (distances.containsKey("页脚")) margins.setFooter(distances.get("页脚"));
                section.setPgMar(margins);
            }
            if (!before.equals(sectionLayoutSignature(section))) adjusted++;
        }
        return new SectionResult(sections.size(), adjusted);
    }

    /**
     * 把分节页面设置压成字符串，用于判断同步前后是否真的发生变化。
     */
    private static String sectionLayoutSignature(SectPr section) {
        var size = section.getPgSz();
        var margins = section.getPgMar();
        var columns = section.getCols();
        var grid = section.getDocGrid();
        var pageNumbers = section.getPgNumType();
        return List.of(
                value(size == null ? null : size.getW()),
                value(size == null ? null : size.getH()),
                value(size == null ? null : size.getOrient()),
                value(size == null ? null : size.getCode()),
                value(margins == null ? null : margins.getTop()),
                value(margins == null ? null : margins.getRight()),
                value(margins == null ? null : margins.getBottom()),
                value(margins == null ? null : margins.getLeft()),
                value(margins == null ? null : margins.getHeader()),
                value(margins == null ? null : margins.getFooter()),
                value(margins == null ? null : margins.getGutter()),
                value(columns == null ? null : columns.getNum()),
                value(columns == null ? null : columns.getSpace()),
                value(columns == null ? null : columns.isEqualWidth()),
                value(columns == null ? null : columns.isSep()),
                value(grid == null ? null : grid.getType()),
                value(grid == null ? null : grid.getLinePitch()),
                value(grid == null ? null : grid.getCharSpace()),
                value(pageNumbers == null ? null : pageNumbers.getFmt()),
                value(pageNumbers == null ? null : pageNumbers.getStart())
        ).toString();
    }

    private static String value(Object value) {
        return value == null ? "<null>" : value.toString();
    }

    @SuppressWarnings("unchecked")
    private static <T> T copy(T value) {
        return value == null ? null : (T) XmlUtils.deepCopy(value);
    }

    /**
     * 只取 body 直属段落。
     *
     * <p>表格单元格内段落不在这里处理，避免误改复杂表格排版。</p>
     */
    private static List<P> directBodyParagraphs(MainDocumentPart main) {
        Body body = main.getJaxbElement().getBody();
        List<P> result = new ArrayList<>();
        for (Object item : body.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (value instanceof P paragraph) {
                result.add(paragraph);
            }
        }
        return result;
    }

    private static List<Tbl> allTables(MainDocumentPart main) {
        List<Tbl> result = new ArrayList<>();
        for (Object node : main.getJaxbElement().getBody().getContent()) {
            Object value = XmlUtils.unwrap(node);
            if (value instanceof Tbl table) {
                result.add(table);
            }
        }
        return result;
    }

    /**
     * 收集所有分节设置。Word 的最后一个分节通常挂在 body 上，其余分节挂在段落属性里。
     */
    private static List<SectPr> allSections(MainDocumentPart main) {
        List<SectPr> result = new ArrayList<>();
        Body body = main.getJaxbElement().getBody();
        for (Object node : body.getContent()) {
            Object value = XmlUtils.unwrap(node);
            if (value instanceof P paragraph
                    && paragraph.getPPr() != null
                    && paragraph.getPPr().getSectPr() != null) {
                result.add(paragraph.getPPr().getSectPr());
            }
        }
        if (body.getSectPr() != null) {
            result.add(body.getSectPr());
        }
        return result;
    }

    private static int countImageRelationships(MainDocumentPart main) {
        int count = 0;
        for (Relationship relationship : main.getRelationshipsPart().getRelationships().getRelationship()) {
            if (relationship.getType() != null && relationship.getType().endsWith("/image")) {
                count++;
            }
        }
        return count;
    }

    /**
     * 根据典型毕业设计文档结构定位各区域边界。
     *
     * <p>这个方法只生成边界索引，不直接修改文档。后续分类会结合这些索引和段落样式判断角色。</p>
     */
    private static DocumentStructure analyzeStructure(WordprocessingMLPackage document, List<P> paragraphs) throws Exception {
        int size = paragraphs.size();
        int coverMarker = findFirstCompact(paragraphs, 0, size, "毕业设计说明书");
        int searchStart = Math.max(0, coverMarker + 1);
        int tocTitle = findFirstCompact(paragraphs, searchStart, size, "目录");

        int chineseAbstract = findPrefixCompact(paragraphs, searchStart, tocTitle < 0 ? size : tocTitle,
                "摘要", "摘要：", "摘要:");
        int chineseKeywordsHint = findPrefixCompact(paragraphs, Math.max(0, chineseAbstract + 1),
                tocTitle < 0 ? size : tocTitle, "关键词", "关键词：", "关键词:");
        int englishAbstractHint = findPrefixCompact(paragraphs, Math.max(0, chineseKeywordsHint + 1),
                tocTitle < 0 ? size : tocTitle, "Abstract", "Abstract：", "Abstract:");

        // 封面题名通常紧跟“毕业设计说明书”，摘要题名通常与封面题名文本相同。
        int coverTitle = nextNonEmpty(paragraphs, searchStart,
                chineseAbstract >= 0 ? chineseAbstract : (tocTitle < 0 ? size : tocTitle));
        String titleText = coverTitle >= 0 ? normalizeText(TextUtils.getText(paragraphs.get(coverTitle))) : "";
        int chineseTitle = titleText.isBlank() ? -1
                : findSameText(paragraphs, coverTitle + 1, tocTitle < 0 ? size : tocTitle, titleText);
        if (chineseTitle < 0) {
            // 文本匹配失败时，用模板样式作为兜底信号。
            chineseTitle = findStyle(paragraphs, document, Math.max(0, coverTitle + 1),
                    tocTitle < 0 ? size : tocTitle, "论文摘要中课题名称");
        }
        if (chineseTitle < 0 && chineseAbstract >= 0) {
            chineseTitle = previousNonEmpty(paragraphs, chineseAbstract - 1, Math.max(0, coverTitle + 1));
        }
        if (coverTitle < 0 && chineseTitle >= 0) {
            coverTitle = chineseTitle;
            titleText = normalizeText(TextUtils.getText(paragraphs.get(chineseTitle)));
        }

        int chineseKeywords = chineseKeywordsHint >= 0 ? chineseKeywordsHint
                : findPrefixCompact(paragraphs, Math.max(0, chineseTitle + 1), tocTitle < 0 ? size : tocTitle, "关键词");
        int englishAbstract = englishAbstractHint >= 0 ? englishAbstractHint
                : findPrefixCompact(paragraphs, Math.max(0, chineseKeywords + 1), tocTitle < 0 ? size : tocTitle, "Abstract");
        int englishTitle = englishAbstract < 0 ? -1
                : previousNonEmpty(paragraphs, englishAbstract - 1, Math.max(chineseKeywords + 1, 0));
        int englishKeywords = findPrefixCompact(paragraphs, Math.max(0, englishAbstract + 1), tocTitle < 0 ? size : tocTitle, "Keywords", "Keywords：", "Keywords:", "Key words", "Key words：", "Key words:");

        // 目录后可能再次出现论文题名，它是正文之前的独立标题，不属于目录条目。
        int afterToc = tocTitle < 0 ? Math.max(0, englishKeywords + 1) : tocTitle + 1;
        int repeatedTitle = titleText.isBlank() ? -1 : findSameText(paragraphs, afterToc, size, titleText);
        int mainStart = findNumberedHeading(paragraphs, repeatedTitle >= 0 ? repeatedTitle + 1 : afterToc, size, false);
        if (mainStart < 0) {
            mainStart = findNumberedHeading(paragraphs, afterToc, size, true);
        }
        if (mainStart < 0) {
            mainStart = size;
        }
        int tocEnd = tocTitle < 0 ? afterToc
                : (repeatedTitle >= 0 && repeatedTitle < mainStart ? repeatedTitle : mainStart);

        int referencesStart = findFirstCompact(paragraphs, mainStart, size, "参考文献");
        if (referencesStart < 0) referencesStart = size;
        int thanksStart = findFirstCompact(paragraphs, referencesStart + 1, size, "致谢");
        if (thanksStart < 0) thanksStart = size;
        int appendixStart = findFirstCompact(paragraphs, thanksStart + 1, size, "附录");
        if (appendixStart < 0) appendixStart = size;

        return new DocumentStructure(
                coverMarker, coverTitle, chineseTitle, chineseKeywords, englishTitle, englishAbstract,
                englishKeywords, tocTitle, tocEnd, repeatedTitle, mainStart,
                referencesStart, thanksStart, appendixStart, size
        );
    }

    /**
     * 将一个段落归类为封面题名、摘要、目录、正文标题等业务角色。
     */
    private static String classify(String text, int index, DocumentStructure structure, String beforeStyleName) {
        String compact = text.replace(" ", "");
        if (index == structure.coverTitle()) return "cover-title";
        if (index == structure.chineseTitle() || index == structure.englishTitle() || index == structure.repeatedTitle()) return "thesis-title";
        if (structure.chineseTitle() >= 0 && structure.englishTitle() >= 0
                && index > structure.chineseTitle() && index < structure.englishTitle()) {
            return index == structure.chineseKeywords() ? "keywords-zh" : "abstract-zh";
        }
        if (structure.englishAbstract() >= 0 && index >= structure.englishAbstract()
                && (structure.tocTitle() < 0 || index < structure.tocTitle())) {
            return index == structure.englishKeywords() ? "keywords-en" : "abstract-en";
        }
        if (index == structure.tocTitle()) return "toc-title";
        if (structure.tocTitle() >= 0 && index > structure.tocTitle() && index < structure.tocEnd()) {
            return isTocLevel2(text) ? "toc-level-2" : "toc-level-1";
        }
        if (compact.equals("参考文献") || compact.equals("致谢") || compact.equals("附录")) return "section-title";
        if (index > structure.referencesStart() && index < structure.thanksStart()) return "reference-item";
        if (index > structure.thanksStart() && index < structure.appendixStart()) return "thanks-body";
        if (index > structure.appendixStart() && index < structure.documentEnd()) return "appendix-body";
        if (index < structure.mainStart() || index >= structure.referencesStart()) return "front-or-back-matter";
        if (CONTINUED_TABLE.matcher(text).matches()) return "continued-table-caption";
        if (FIGURE_CAPTION.matcher(text).matches()) return "figure-caption";
        if (TABLE_CAPTION.matcher(text).matches()) return "table-caption";
        if (LEVEL_4.matcher(text).matches()) return "heading-4";
        if (LEVEL_3.matcher(text).matches()) return "heading-3";
        if (LEVEL_2.matcher(text).matches()) return "heading-2";
        if (LEVEL_1.matcher(text).matches()) return "heading-1";
        if ("论文一级标题".equals(beforeStyleName)) return "heading-1";
        if ("论文二级标题".equals(beforeStyleName)) return "heading-2";
        if ("论文三级标题".equals(beforeStyleName)) return "heading-3";
        if ("论文四级标题".equals(beforeStyleName)) return "heading-4";
        if ("论文表格后面段落正文".equals(beforeStyleName)) return "table-following-body";
        return "body";
    }

    /**
     * 把业务角色映射到模板中的段落样式名称。
     */
    private static String styleForRole(String role) {
        return switch (role) {
            case "cover-title" -> "论文封面课题名称";
            case "thesis-title" -> "论文摘要中课题名称";
            case "abstract-zh", "abstract-en" -> "论文摘要正文";
            case "keywords-zh", "keywords-en" -> "论文正文";
            case "toc-title" -> "参考文献及致谢";
            case "toc-level-1" -> "toc 1";
            case "toc-level-2" -> "toc 2";
            case "heading-1" -> "论文一级标题";
            case "heading-2" -> "论文二级标题";
            case "heading-3" -> "论文三级标题";
            case "heading-4" -> "论文四级标题";
            case "table-caption" -> "论文表格序号及题目";
            case "figure-caption" -> "论文图号图名";
            case "continued-table-caption" -> "论文续表";
            case "table-following-body" -> "论文表格后面段落正文";
            case "reference-item" -> "论文参考文献";
            case "section-title" -> "参考文献及致谢";
            case "body", "thanks-body", "appendix-body" -> "论文正文";
            default -> null;
        };
    }

    private static boolean shouldApplyStyle(String role, String beforeStyleName) {
        return styleForRole(role) != null;
    }

    private static boolean shouldNormalizeDirectOverrides(String role) {
        return switch (role) {
            case "cover-title", "thesis-title", "abstract-zh", "keywords-zh", "abstract-en", "keywords-en",
                    "toc-title", "toc-level-1", "toc-level-2", "body",
                    "heading-1", "heading-2", "heading-3", "heading-4",
                    "section-title", "reference-item", "thanks-body", "appendix-body",
                    "figure-caption", "table-caption", "continued-table-caption" -> true;
            default -> false;
        };
    }

    /**
     * Mixed label/body paragraphs cannot be represented by a single Word paragraph style.
     * Apply the paragraph rule first, then explicit character rules for each role.
     */
    private static void applyRoleFormatting(P paragraph, String role) {
        switch (role) {
            case "abstract-zh" -> {
                boolean hasLabel = compact(TextUtils.getText(paragraph)).startsWith("摘要");
                setFirstLineCharacters(paragraph, hasLabel ? 0 : 200);
                formatLeadingLabel(paragraph, "摘要", "黑体", "Times New Roman", 28);
            }
            case "keywords-zh" -> {
                setFirstLineCharacters(paragraph, 0);
                formatLeadingLabel(paragraph, "关键词", "黑体", "Times New Roman", 28);
            }
            case "abstract-en" -> {
                boolean hasLabel = compact(TextUtils.getText(paragraph)).toLowerCase(Locale.ROOT).startsWith("abstract");
                setFirstLineCharacters(paragraph, hasLabel ? 0 : 200);
                formatLeadingLabel(paragraph, "Abstract", null, "Times New Roman", 28);
            }
            case "keywords-en" -> {
                setFirstLineCharacters(paragraph, 0);
                String text = normalizeText(TextUtils.getText(paragraph)).toLowerCase(Locale.ROOT);
                formatLeadingLabel(paragraph, text.startsWith("key words") ? "Key words" : "Keywords",
                        null, "Times New Roman", 28);
            }
            case "heading-1" -> setOutlineLevel(paragraph, 0);
            case "heading-2" -> setOutlineLevel(paragraph, 1);
            case "heading-3" -> setOutlineLevel(paragraph, 2);
            case "heading-4" -> setOutlineLevel(paragraph, 3);
            default -> {
                // Named template style is sufficient for this role.
            }
        }
    }

    private static void setFirstLineCharacters(P paragraph, int characters) {
        PPr pPr = paragraph.getPPr();
        if (pPr == null) {
            pPr = new PPr();
            paragraph.setPPr(pPr);
        }
        PPrBase.Ind ind = pPr.getInd() == null ? new PPrBase.Ind() : pPr.getInd();
        ind.setHanging(null);
        ind.setHangingChars(null);
        if (characters == 0) {
            ind.setFirstLine(BigInteger.ZERO);
            ind.setFirstLineChars(BigInteger.ZERO);
        } else {
            ind.setFirstLine(null);
            ind.setFirstLineChars(BigInteger.valueOf(characters));
        }
        pPr.setInd(ind);
    }

    private static void setOutlineLevel(P paragraph, int level) {
        PPr pPr = paragraph.getPPr();
        if (pPr == null) {
            pPr = new PPr();
            paragraph.setPPr(pPr);
        }
        PPrBase.OutlineLvl outline = new PPrBase.OutlineLvl();
        outline.setVal(BigInteger.valueOf(level));
        pPr.setOutlineLvl(outline);
    }

    private static void setParagraphFormat(P paragraph, JcEnumeration alignment, int exactLineTwips, int firstLineChars) {
        PPr pPr = paragraph.getPPr();
        if (pPr == null) {
            pPr = new PPr();
            paragraph.setPPr(pPr);
        }
        Jc jc = new Jc();
        jc.setVal(alignment);
        pPr.setJc(jc);
        PPrBase.Spacing spacing = new PPrBase.Spacing();
        spacing.setBefore(BigInteger.ZERO);
        spacing.setAfter(BigInteger.ZERO);
        spacing.setLine(BigInteger.valueOf(exactLineTwips));
        spacing.setLineRule(org.docx4j.wml.STLineSpacingRule.EXACT);
        pPr.setSpacing(spacing);
        PPrBase.Ind ind = new PPrBase.Ind();
        if (firstLineChars > 0) ind.setFirstLineChars(BigInteger.valueOf(firstLineChars));
        pPr.setInd(ind);
    }

    private static void setAllRunFormat(P paragraph, String eastAsia, String latin, int halfPoints, boolean bold) {
        for (Object item : paragraph.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (value instanceof R run) setRunFormat(run, eastAsia, latin, halfPoints, bold);
        }
    }

    private static void setRunFormat(R run, String eastAsia, String latin, int halfPoints, boolean bold) {
        RPr rPr = run.getRPr();
        if (rPr == null) {
            rPr = new RPr();
            run.setRPr(rPr);
        }
        RFonts fonts = new RFonts();
        if (eastAsia != null) fonts.setEastAsia(eastAsia);
        if (latin != null) {
            fonts.setAscii(latin);
            fonts.setHAnsi(latin);
            fonts.setCs(latin);
        }
        rPr.setRFonts(fonts);
        HpsMeasure size = new HpsMeasure();
        size.setVal(BigInteger.valueOf(halfPoints));
        rPr.setSz(size);
        rPr.setSzCs(copy(size));
        BooleanDefaultTrue value = new BooleanDefaultTrue();
        value.setVal(bold);
        rPr.setB(value);
        rPr.setBCs(copy(value));
    }

    private static void formatLeadingLabel(
            P paragraph,
            String label,
            String eastAsia,
            String latin,
            int halfPoints
    ) {
        String compactLabel = compact(label).toLowerCase(Locale.ROOT);
        boolean matched = false;
        for (int itemIndex = 0; itemIndex < paragraph.getContent().size(); itemIndex++) {
            Object item = paragraph.getContent().get(itemIndex);
            Object value = XmlUtils.unwrap(item);
            if (!(value instanceof R run)) continue;
            String valueText = runText(run);
            if (valueText.isBlank()) continue;
            String normalized = compact(valueText).toLowerCase(Locale.ROOT);
            if (!matched && (normalized.startsWith(compactLabel)
                    || compactLabel.startsWith(normalized.replace(":", "").replace("：", "")))) {
                int separator = Math.max(valueText.indexOf('：'), valueText.indexOf(':'));
                if (separator >= 0 && separator + 1 < valueText.length()
                        && !valueText.substring(separator + 1).isBlank()) {
                    R labelRun = copy(run);
                    R bodyRun = copy(run);
                    replaceRunText(labelRun, valueText.substring(0, separator + 1));
                    replaceRunText(bodyRun, valueText.substring(separator + 1));
                    setRunFormat(labelRun, eastAsia, latin, halfPoints, true);
                    paragraph.getContent().set(itemIndex, labelRun);
                    paragraph.getContent().add(itemIndex + 1, bodyRun);
                    return;
                }
                setRunFormat(run, eastAsia, latin, halfPoints, true);
                matched = normalized.contains(":") || normalized.contains("：") || normalized.length() >= compactLabel.length();
            } else if (matched) {
                String punctuation = valueText.trim();
                if (punctuation.matches("^[：:]+$")) {
                    setRunFormat(run, eastAsia, latin, halfPoints, true);
                    continue;
                }
                break;
            }
        }
    }

    private static void replaceRunText(R run, String replacement) {
        boolean written = false;
        for (Object item : run.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (value instanceof Text text) {
                text.setValue(written ? "" : replacement);
                text.setSpace("preserve");
                written = true;
            }
        }
    }

    private static void applyParagraphStyle(P paragraph, String styleId) {
        PPr pPr = paragraph.getPPr();
        if (pPr == null) {
            pPr = new PPr();
            paragraph.setPPr(pPr);
        }
        PPrBase.PStyle pStyle = new PPrBase.PStyle();
        pStyle.setVal(styleId);
        pPr.setPStyle(pStyle);

        pPr.setJc(null);
        pPr.setSpacing(null);
        pPr.setInd(null);
        pPr.setTextAlignment(null);
        pPr.setTabs(null);

        // 清掉段落内局部字体/字号等直接格式，让新段落样式能够真正生效。
        for (Object content : paragraph.getContent()) {
            Object value = XmlUtils.unwrap(content);
            if (value instanceof R run) {
                clearConflictingRunFormatting(run);
            }
        }
    }

    /**
     * Word 自动目录会读取标题大纲级别。把目录标题设为 9 级，相当于排除在目录之外。
     */
    private static void excludeTocTitleFromGeneratedContents(P paragraph) {
        PPr pPr = paragraph.getPPr();
        if (pPr == null) {
            pPr = new PPr();
            paragraph.setPPr(pPr);
        }
        PPrBase.OutlineLvl outline = new PPrBase.OutlineLvl();
        outline.setVal(BigInteger.valueOf(9));
        pPr.setOutlineLvl(outline);
    }

    private static void requestWordFieldRefresh(WordprocessingMLPackage document) throws Exception {
        var settingsPart = document.getMainDocumentPart().getDocumentSettingsPart();
        if (settingsPart == null || settingsPart.getJaxbElement() == null) return;
        BooleanDefaultTrue updateFields = new BooleanDefaultTrue();
        updateFields.setVal(true);
        settingsPart.getJaxbElement().setUpdateFields(updateFields);
    }

    /**
     * On the local Windows edition, refresh cached TOC/page-number fields through Word.
     * Failure is non-fatal: updateFields remains set and Word can refresh on open.
     */
    private static boolean refreshFieldsWithMicrosoftWord(Path document) {
        if (!System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("win")) return false;
        String escaped = document.toAbsolutePath().toString().replace("'", "''");
        String script = "$ErrorActionPreference='Stop';"
                + "$w=New-Object -ComObject Word.Application;"
                + "$w.Visible=$false;$w.DisplayAlerts=0;$w.Options.Pagination=$true;"
                + "try{$d=$w.Documents.Open('" + escaped + "',$false,$false);"
                + "$d.Repaginate();"
                + "foreach($t in $d.TablesOfContents){$t.Update()};"
                + "foreach($s in $d.Sections){foreach($h in $s.Headers){$h.Range.Fields.Update()|Out-Null};"
                + "foreach($f in $s.Footers){$f.Range.Fields.Update()|Out-Null}};"
                + "$d.Fields.Update()|Out-Null;$d.Repaginate();"
                + "foreach($s in $d.Sections){foreach($f in $s.Footers){$f.Range.Fields.Update()|Out-Null}};"
                + "$d.Save();$d.Close()}"
                + "finally{$w.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($w)|Out-Null}";
        Process process = null;
        try {
            process = new ProcessBuilder("powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script)
                    .redirectErrorStream(true)
                    .start();
            boolean finished = process.waitFor(90, TimeUnit.SECONDS);
            if (!finished) {
                process.destroyForcibly();
                return false;
            }
            return process.exitValue() == 0;
        } catch (Exception ignored) {
            if (process != null) process.destroyForcibly();
            return false;
        }
    }

    private static void clearConflictingRunFormatting(R run) {
        RPr properties = run.getRPr();
        if (properties == null) {
            return;
        }
        properties.setRFonts(null);
        properties.setSz(null);
        properties.setSzCs(null);
        properties.setB(null);
        properties.setBCs(null);
        properties.setI(null);
        properties.setICs(null);
        properties.setColor(null);
        properties.setSpacing(null);
        properties.setPosition(null);
    }

    /**
     * 根据样式显示名称查找 styleId。
     *
     * <p>Word 内部应用样式用的是 styleId，但模板维护人员更容易理解中文样式名。</p>
     */
    private static Map<String, String> styleIdsByName(WordprocessingMLPackage document) throws Docx4JException {
        StyleDefinitionsPart part = document.getMainDocumentPart().getStyleDefinitionsPart();
        if (part == null || part.getJaxbElement() == null) {
            throw new IllegalStateException("Document has no style definitions part.");
        }
        Map<String, String> result = new LinkedHashMap<>();
        for (Style style : part.getJaxbElement().getStyle()) {
            if (style.getName() != null && style.getName().getVal() != null) {
                String name = style.getName().getVal();
                result.put(name, style.getStyleId());
                result.putIfAbsent(name.toLowerCase(Locale.ROOT), style.getStyleId());
            }
        }
        return result;
    }

    private static String styleNameById(WordprocessingMLPackage document, String styleId) throws Docx4JException {
        if (styleId == null) return "Normal";
        StyleDefinitionsPart part = document.getMainDocumentPart().getStyleDefinitionsPart();
        for (Style style : part.getJaxbElement().getStyle()) {
            if (styleId.equals(style.getStyleId())) {
                return style.getName() == null ? styleId : style.getName().getVal();
            }
        }
        return styleId;
    }

    private static String getStyleId(P paragraph) {
        return paragraph.getPPr() == null || paragraph.getPPr().getPStyle() == null
                ? null
                : paragraph.getPPr().getPStyle().getVal();
    }

    private static void ensureRequiredStyles(Map<String, String> styleIds) {
        for (String name : TEMPLATE_STYLE_NAMES) {
            if (!styleIds.containsKey(name)) {
                throw new IllegalStateException("Template is missing required style: " + name);
            }
        }
    }

    /**
     * 查找去空格后完全等于指定文本的段落，适合处理“目录”“参考文献”等独立标题。
     */
    private static int findFirstCompact(List<P> paragraphs, int start, int end, String expected) {
        String target = compact(expected);
        for (int i = Math.max(0, start); i < Math.min(end, paragraphs.size()); i++) {
            if (compact(TextUtils.getText(paragraphs.get(i))).equals(target)) {
                return i;
            }
        }
        return -1;
    }

    /**
     * 查找去空格后以前缀开头的段落，适合匹配“关键词：”“Keywords:”等多种写法。
     */
    private static int findPrefixCompact(List<P> paragraphs, int start, int end, String... prefixes) {
        List<String> targets = new ArrayList<>();
        for (String prefix : prefixes) {
            targets.add(compact(prefix).toLowerCase(Locale.ROOT));
        }
        for (int i = Math.max(0, start); i < Math.min(end, paragraphs.size()); i++) {
            String value = compact(TextUtils.getText(paragraphs.get(i))).toLowerCase(Locale.ROOT);
            for (String target : targets) {
                if (value.startsWith(target)) {
                    return i;
                }
            }
        }
        return -1;
    }

    private static int findSameText(List<P> paragraphs, int start, int end, String expected) {
        if (expected == null || expected.isBlank()) return -1;
        for (int i = Math.max(0, start); i < Math.min(end, paragraphs.size()); i++) {
            if (normalizeText(TextUtils.getText(paragraphs.get(i))).equals(expected)) {
                return i;
            }
        }
        return -1;
    }

    private static int findStyle(
            List<P> paragraphs,
            WordprocessingMLPackage document,
            int start,
            int end,
            String styleName
    ) throws Exception {
        for (int i = Math.max(0, start); i < Math.min(end, paragraphs.size()); i++) {
            if (styleName.equals(styleNameById(document, getStyleId(paragraphs.get(i))))) {
                return i;
            }
        }
        return -1;
    }

    private static int nextNonEmpty(List<P> paragraphs, int start, int end) {
        for (int i = Math.max(0, start); i < Math.min(end, paragraphs.size()); i++) {
            if (!normalizeText(TextUtils.getText(paragraphs.get(i))).isBlank()) {
                return i;
            }
        }
        return -1;
    }

    private static int previousNonEmpty(List<P> paragraphs, int start, int lowerBound) {
        for (int i = Math.min(start, paragraphs.size() - 1); i >= Math.max(0, lowerBound); i--) {
            if (!normalizeText(TextUtils.getText(paragraphs.get(i))).isBlank()) {
                return i;
            }
        }
        return -1;
    }

    /**
     * 查找正文第一个一级标题。目录条目也像“1 标题 3”，所以默认会排除疑似目录项。
     */
    private static int findNumberedHeading(List<P> paragraphs, int start, int end, boolean allowTocLikeText) {
        for (int i = Math.max(0, start); i < Math.min(end, paragraphs.size()); i++) {
            String text = normalizeText(TextUtils.getText(paragraphs.get(i)));
            if (!LEVEL_1.matcher(text).matches()) continue;
            if (!allowTocLikeText && looksLikeTocEntry(text)) continue;
            return i;
        }
        return -1;
    }

    private static boolean looksLikeTocEntry(String text) {
        return text.matches(".*\\s+(?:\\d+|[IVXLCDM]+)$");
    }

    private static boolean isTocLevel2(String text) {
        return LEVEL_2.matcher(text).matches() || LEVEL_3.matcher(text).matches() || LEVEL_4.matcher(text).matches();
    }

    private static String compact(String text) {
        return normalizeText(text).replace(" ", "");
    }

    private static String normalizeText(String text) {
        if (text == null) return "";
        return text.replace('\u00A0', ' ').replaceAll("[\\r\\n\\t]+", " ").replaceAll("\\s+", " ").trim();
    }

    private static String displayText(String text) {
        if (text == null) return "";
        return text.replace('\u00A0', ' ').replaceAll("[\\r\\n\\t]+", " ").trim();
    }

    private static String preview(String text) {
        return text.length() <= 90 ? text : text.substring(0, 90) + "...";
    }

    private static void requireDocx(Path path, String role) throws IOException {
        if (!Files.isRegularFile(path)) {
            throw new IllegalArgumentException(role + " file not found: " + path);
        }
        if (!path.getFileName().toString().toLowerCase(Locale.ROOT).endsWith(".docx")) {
            throw new IllegalArgumentException(role + " must be a .docx file: " + path);
        }
        try (InputStream input = Files.newInputStream(path)) {
            byte[] magic = input.readNBytes(4);
            if (magic.length != 4 || magic[0] != 'P' || magic[1] != 'K') {
                throw new IllegalArgumentException(role + " is not an OOXML ZIP package: " + path);
            }
        }
    }

    private static boolean isReadableZip(Path path) {
        try (ZipFile ignored = new ZipFile(path.toFile())) {
            return true;
        } catch (IOException error) {
            return false;
        }
    }

    private static String sha256(Path path) throws Exception {
        MessageDigest digest = MessageDigest.getInstance("SHA-256");
        try (InputStream input = Files.newInputStream(path)) {
            byte[] buffer = new byte[64 * 1024];
            int read;
            while ((read = input.read(buffer)) >= 0) {
                if (read > 0) digest.update(buffer, 0, read);
            }
        }
        StringBuilder result = new StringBuilder();
        for (byte value : digest.digest()) {
            result.append(String.format("%02X", value));
        }
        return result.toString();
    }

    private record SectionResult(int inspected, int adjusted) {
    }

    private record RunSpan(R run, int start, int end) {
    }

    /**
     * 文档各区域的边界索引。
     *
     * <p>使用 record 是因为这里只需要一组不可变数据，不需要额外行为。</p>
     */
    private record DocumentStructure(
            int coverMarker,
            int coverTitle,
            int chineseTitle,
            int chineseKeywords,
            int englishTitle,
            int englishAbstract,
            int englishKeywords,
            int tocTitle,
            int tocEnd,
            int repeatedTitle,
            int mainStart,
            int referencesStart,
            int thanksStart,
            int appendixStart,
            int documentEnd
    ) {
    }
}
