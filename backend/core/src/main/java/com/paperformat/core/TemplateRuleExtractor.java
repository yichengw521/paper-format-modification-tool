package com.paperformat.core;

import org.docx4j.TextUtils;
import org.docx4j.XmlUtils;
import org.docx4j.model.PropertyResolver;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tr;
import org.docx4j.wml.Tc;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Reads the template itself instead of returning a hard-coded description.
 * Named styles are authoritative. Human-readable instructions are retained as
 * evidence and become the fallback when a template only describes a rule in text.
 */
final class TemplateRuleExtractor {
    private static final Map<String, RoleDefinition> ROLES = roleDefinitions();
    private static final Map<String, String> CANONICAL_STYLES = canonicalStyles();

    private TemplateRuleExtractor() {
    }

    static void ensureOperationalStyles(WordprocessingMLPackage template) throws Exception {
        ensureOperationalStyles(template, null);
    }

    /**
     * Builds the canonical styles used by the formatter.  When a template contains only prose
     * requirements, the source document supplies the neutral base style and the prose requirement
     * is applied as an override.  This prevents a text-only requirements document from silently
     * imposing its own Normal style on the user's thesis.
     */
    static void ensureOperationalStyles(
            WordprocessingMLPackage template,
            WordprocessingMLPackage source
    ) throws Exception {
        StyleDefinitionsPart part = template.getMainDocumentPart().getStyleDefinitionsPart();
        Map<String, Style> styles = stylesByName(template);
        List<P> paragraphs = directParagraphs(template);
        List<String> instructions = extractInstructions(template);
        Map<String, String> instructionByRole = mapInstructionsToRoles(instructions);
        boolean textOnly = source != null
                && "TEXT_INSTRUCTIONS_ONLY".equals(detectMode(template, instructions, styles));
        PropertyResolver resolver = new PropertyResolver(template);
        Map<String, Style> sourceStyles = source == null ? Map.of() : stylesByName(source);
        List<P> sourceParagraphs = source == null ? List.of() : directParagraphs(source);

        for (Map.Entry<String, String> entry : CANONICAL_STYLES.entrySet()) {
            String role = entry.getKey();
            String canonicalName = entry.getValue();
            Style existing = styles.get(canonicalName.toLowerCase(Locale.ROOT));
            if (existing != null) {
                // 同名样式是最高优先级证据。文字说明只用于解释或补齐缺失样式，
                // 不能反过来覆盖模板样式中的缩进、段距和制表位。
                if (role.equals("toc1") || role.equals("toc2")) {
                    P sample = findParagraphUsingStyle(paragraphs, existing.getStyleId());
                    mergeDirectParagraphPropertiesIntoStyle(existing, sample);
                }
                if (role.equals("abstract")) forceFirstLine(existing, 0);
                continue;
            }

            RoleDefinition definition = ROLES.get(role);
            // A prose-only requirements document may happen to contain a localized built-in
            // style named "正文".  That is document chrome, not an authoritative format sample.
            // In this mode the source document must supply the base style/paragraph structure,
            // and the prose requirements must be applied as explicit overrides.
            Style alias = textOnly || definition == null
                    ? null : findStyle(styles, definition.styleCandidates());
            boolean sourceDerived = false;
            if (alias == null && definition != null && source != null) {
                alias = findStyle(sourceStyles, definition.styleCandidates());
                sourceDerived = alias != null;
            }
            Style created;
            if (alias != null) {
                created = cloneAs(alias, canonicalName, role);
            } else {
                P sample = textOnly ? null
                        : role.equals("tableCell") ? findTableCellSample(template)
                        : role.equals("figureCaption") ? findFigureCaptionSample(template)
                        : findSampleParagraph(paragraphs, role);
                if ((sample == null || textOnly) && source != null) {
                    sample = findSourceSampleParagraph(sourceParagraphs, role);
                    sourceDerived = sample != null;
                }
                if (sample != null) {
                    PropertyResolver sampleResolver = sourceDerived ? new PropertyResolver(source) : resolver;
                    created = styleFromSample(sample, canonicalName, role, sampleResolver);
                } else {
                    String instruction = instructionByRole.get(role);
                    if (instruction == null && role.equals("thanksBody")) instruction = instructionByRole.get("body");
                    if (instruction == null && role.equals("heading4")) instruction = instructionByRole.get("body");
                    if (instruction != null) {
                        created = styleFromInstruction(canonicalName, role, executionInstruction(role, instruction));
                    } else {
                        Style normal = findStyle(sourceStyles, List.of("normal", "正文"));
                        created = normal == null
                                ? baseStyle(canonicalName, role)
                                : cloneAs(normal, canonicalName, role);
                        sourceDerived = normal != null;
                    }
                }
            }
            String instruction = instructionByRole.get(role);
            if (instruction == null && role.equals("thanksBody")) instruction = instructionByRole.get("body");
            if (instruction == null && role.equals("heading4")) instruction = instructionByRole.get("body");
            if (instruction != null && (textOnly || sourceDerived || role.equals("tableCell"))) {
                applyInstructionOverrides(created, executionInstruction(role, instruction));
            }
            part.getJaxbElement().getStyle().add(created);
            styles.put(canonicalName.toLowerCase(Locale.ROOT), created);
        }
    }

    static ProcessingReport.TemplateRules extract(WordprocessingMLPackage template) throws Exception {
        return extract(template, detectMode(template));
    }

    static ProcessingReport.TemplateRules extract(
            WordprocessingMLPackage template,
            String detectedMode
    ) throws Exception {
        List<String> instructions = extractInstructions(template);
        Map<String, String> instructionByRole = mapInstructionsToRoles(instructions);
        Map<String, Style> stylesByName = stylesByName(template);
        PropertyResolver resolver = new PropertyResolver(template);

        Map<String, ProcessingReport.StyleRule> rules = new LinkedHashMap<>();
        for (Map.Entry<String, RoleDefinition> entry : ROLES.entrySet()) {
            String key = entry.getKey();
            RoleDefinition definition = entry.getValue();
            Style style = findStyle(stylesByName, definition.styleCandidates());
            String evidence = instructionByRole.get(key);
            if (style != null) {
                RPr rPr = resolver.getEffectiveRPr(style.getStyleId());
                PPr pPr = resolver.getEffectivePPr(style.getStyleId());
                RPr directRPr = style.getRPr();
                PPr directPPr = style.getPPr();
                RPr fontRPr = hasExplicitFont(directRPr) ? directRPr : rPr;
                RPr sizeRPr = directRPr != null && directRPr.getSz() != null ? directRPr : rPr;
                PPr paragraphPPr = hasExplicitParagraphFormatting(directPPr) ? directPPr : pPr;
                rules.put(key, new ProcessingReport.StyleRule(
                        definition.label(),
                        styleName(style),
                        describeFonts(fontRPr),
                        describeSize(sizeRPr),
                        describeParagraph(paragraphPPr),
                        evidence == null ? "NAMED_STYLE" : "NAMED_STYLE+TEXT_INSTRUCTION",
                        evidence == null ? "模板样式定义：" + styleName(style) : evidence
                ));
            } else if (evidence != null) {
                rules.put(key, new ProcessingReport.StyleRule(
                        definition.label(),
                        "（模板未定义命名样式）",
                        fontsFromText(evidence),
                        sizeFromText(evidence),
                        paragraphFromText(evidence),
                        "TEXT_INSTRUCTION",
                        evidence
                ));
            }
        }

        SectPr section = firstSection(template);
        String mode = detectedMode == null ? detectMode(template, instructions, stylesByName) : detectedMode;
        return new ProcessingReport.TemplateRules(
                mode,
                modeDescription(mode),
                describePageSize(section),
                describeMargins(section),
                describeHeaderFooterDistances(section),
                rules,
                instructions
        );
    }

    private static Map<String, RoleDefinition> roleDefinitions() {
        Map<String, RoleDefinition> result = new LinkedHashMap<>();
        result.put("coverTitle", new RoleDefinition("封面课题名称", List.of("论文封面课题名称", "封面课题名称", "封面标题")));
        result.put("title", new RoleDefinition("中英文论文题目", List.of("论文摘要中课题名称", "论文题目", "摘要中课题名称")));
        result.put("abstract", new RoleDefinition("摘要正文", List.of("论文摘要正文", "摘要正文")));
        result.put("figureCaption", new RoleDefinition("图号及图名", List.of("论文图号图名", "图号图名", "图题")));
        result.put("tableCell", new RoleDefinition("表格正文", List.of("论文表格正文", "表格正文", "表内文字")));
        result.put("tocTitle", new RoleDefinition("目录标题", List.of("参考文献及致谢", "目录标题")));
        result.put("toc1", new RoleDefinition("一级目录条目", List.of("toc 1", "目录 1", "目录一级")));
        result.put("toc2", new RoleDefinition("二级目录条目", List.of("toc 2", "目录 2", "目录二级")));
        result.put("heading1", new RoleDefinition("一级标题", List.of("论文一级标题", "一级标题")));
        result.put("heading2", new RoleDefinition("二级标题", List.of("论文二级标题", "二级标题")));
        result.put("heading3", new RoleDefinition("三级标题", List.of("论文三级标题", "三级标题")));
        result.put("heading4", new RoleDefinition("四级标题", List.of("论文四级标题", "四级标题")));
        result.put("body", new RoleDefinition("正文", List.of("论文正文", "正文")));
        result.put("tableCaption", new RoleDefinition("表格序号及题目", List.of("论文表格序号及题目", "表格序号及题目", "表题")));
        result.put("continuedTableCaption", new RoleDefinition("续表题注", List.of("论文续表", "续表")));
        result.put("tableFollowingBody", new RoleDefinition("表格后第一段", List.of("论文表格后面段落正文", "表格后面段落正文")));
        result.put("references", new RoleDefinition("参考文献条目", List.of("论文参考文献", "参考文献")));
        result.put("sectionTitle", new RoleDefinition("参考文献致谢附录标题", List.of("参考文献及致谢", "章节题名")));
        result.put("thanksBody", new RoleDefinition("致谢正文", List.of("论文正文", "正文")));
        return result;
    }

    private static Map<String, String> canonicalStyles() {
        Map<String, String> result = new LinkedHashMap<>();
        result.put("coverTitle", "论文封面课题名称");
        result.put("title", "论文摘要中课题名称");
        result.put("abstract", "论文摘要正文");
        result.put("figureCaption", "论文图号图名");
        result.put("tableCell", "论文表格正文");
        result.put("tocTitle", "参考文献及致谢");
        result.put("toc1", "toc 1");
        result.put("toc2", "toc 2");
        result.put("heading1", "论文一级标题");
        result.put("heading2", "论文二级标题");
        result.put("heading3", "论文三级标题");
        result.put("heading4", "论文四级标题");
        result.put("body", "论文正文");
        result.put("tableCaption", "论文表格序号及题目");
        result.put("continuedTableCaption", "论文续表");
        result.put("tableFollowingBody", "论文表格后面段落正文");
        result.put("references", "论文参考文献");
        return result;
    }

    private static Style cloneAs(Style source, String name, String role) {
        Style result = (Style) XmlUtils.deepCopy(source);
        result.setStyleId(styleId(role));
        Style.Name styleName = new Style.Name();
        styleName.setVal(name);
        result.setName(styleName);
        result.setDefault(false);
        result.setCustomStyle(true);
        return result;
    }

    private static Style styleFromSample(P sample, String name, String role, PropertyResolver resolver) {
        Style result = baseStyle(name, role);
        String sampleStyleId = sample.getPPr() == null || sample.getPPr().getPStyle() == null
                ? null : sample.getPPr().getPStyle().getVal();
        PPr effectivePPr = sampleStyleId == null ? null : resolver.getEffectivePPr(sampleStyleId);
        PPr pPr = effectivePPr == null ? new PPr() : (PPr) XmlUtils.deepCopy(effectivePPr);
        result.setPPr(pPr);
        mergeDirectParagraphPropertiesIntoStyle(result, sample);
        pPr = result.getPPr();
        pPr.setPStyle(null);
        pPr.setSectPr(null);
        result.setPPr(pPr);

        RPr rPr = null;
        for (Object item : sample.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (value instanceof R run && run.getRPr() != null) {
                rPr = (RPr) XmlUtils.deepCopy(run.getRPr());
                break;
            }
        }
        if (rPr == null) rPr = sampleStyleId == null
                ? resolver.getEffectiveRPr(null, sample.getPPr())
                : resolver.getEffectiveRPr(sampleStyleId);
        result.setRPr(rPr == null ? new RPr() : (RPr) XmlUtils.deepCopy(rPr));
        return result;
    }

    private static Style styleFromInstruction(String name, String role, String instruction) {
        Style result = baseStyle(name, role);
        result.setRPr(new RPr());
        result.setPPr(new PPr());
        applyInstructionOverrides(result, instruction);
        return result;
    }

    private static void applyInstructionOverrides(Style style, String instruction) {
        RPr rPr = style.getRPr();
        if (rPr == null) {
            rPr = new RPr();
            style.setRPr(rPr);
        }
        RFonts fonts = new RFonts();
        boolean hasFont = false;
        if (instruction.contains("黑体")) { fonts.setEastAsia("黑体"); hasFont = true; }
        else if (instruction.contains("宋体")) { fonts.setEastAsia("宋体"); hasFont = true; }
        if (instruction.toLowerCase(Locale.ROOT).contains("times new roman")) {
            fonts.setAscii("Times New Roman");
            fonts.setHAnsi("Times New Roman");
            hasFont = true;
        }
        if (hasFont) {
            RFonts current = rPr.getRFonts();
            if (current == null) current = new RFonts();
            if (fonts.getEastAsia() != null) current.setEastAsia(fonts.getEastAsia());
            if (fonts.getAscii() != null) current.setAscii(fonts.getAscii());
            if (fonts.getHAnsi() != null) current.setHAnsi(fonts.getHAnsi());
            rPr.setRFonts(current);
        }
        BigInteger halfPoints = halfPointsFromText(instruction);
        if (halfPoints != null) {
            org.docx4j.wml.HpsMeasure size = new org.docx4j.wml.HpsMeasure();
            size.setVal(halfPoints);
            rPr.setSz(size);
            rPr.setSzCs((org.docx4j.wml.HpsMeasure) XmlUtils.deepCopy(size));
        }
        if (instruction.contains("加粗")) {
            org.docx4j.wml.BooleanDefaultTrue bold = new org.docx4j.wml.BooleanDefaultTrue();
            bold.setVal(true);
            rPr.setB(bold);
            rPr.setBCs((org.docx4j.wml.BooleanDefaultTrue) XmlUtils.deepCopy(bold));
        }

        PPr pPr = style.getPPr();
        if (pPr == null) {
            pPr = new PPr();
            style.setPPr(pPr);
        }
        if (instruction.contains("居中")) {
            org.docx4j.wml.Jc jc = new org.docx4j.wml.Jc();
            jc.setVal(org.docx4j.wml.JcEnumeration.CENTER);
            pPr.setJc(jc);
        } else if (instruction.contains("两端对齐")) {
            org.docx4j.wml.Jc jc = new org.docx4j.wml.Jc();
            jc.setVal(org.docx4j.wml.JcEnumeration.BOTH);
            pPr.setJc(jc);
        } else if (instruction.contains("左对齐") || instruction.contains("左顶格")) {
            org.docx4j.wml.Jc jc = new org.docx4j.wml.Jc();
            jc.setVal(org.docx4j.wml.JcEnumeration.LEFT);
            pPr.setJc(jc);
        } else if (instruction.contains("右对齐")) {
            org.docx4j.wml.Jc jc = new org.docx4j.wml.Jc();
            jc.setVal(org.docx4j.wml.JcEnumeration.RIGHT);
            pPr.setJc(jc);
        }
        PPrBase.Spacing parsedSpacing = spacingFromText(instruction);
        if (parsedSpacing != null) {
            PPrBase.Spacing spacing = pPr.getSpacing();
            if (spacing == null) spacing = new PPrBase.Spacing();
            if (parsedSpacing.getLine() != null) spacing.setLine(parsedSpacing.getLine());
            if (parsedSpacing.getLineRule() != null) spacing.setLineRule(parsedSpacing.getLineRule());
            if (parsedSpacing.getBefore() != null) {
                spacing.setBefore(parsedSpacing.getBefore());
                spacing.setBeforeLines(null);
            }
            if (parsedSpacing.getBeforeLines() != null) {
                spacing.setBeforeLines(parsedSpacing.getBeforeLines());
                spacing.setBefore(null);
            }
            if (parsedSpacing.getAfter() != null) {
                spacing.setAfter(parsedSpacing.getAfter());
                spacing.setAfterLines(null);
            }
            if (parsedSpacing.getAfterLines() != null) {
                spacing.setAfterLines(parsedSpacing.getAfterLines());
                spacing.setAfter(null);
            }
            pPr.setSpacing(spacing);
        }
        if (instruction.contains("首行") && instruction.contains("2个字符")) {
            PPrBase.Ind ind = pPr.getInd() == null ? new PPrBase.Ind() : pPr.getInd();
            ind.setFirstLineChars(BigInteger.valueOf(200));
            pPr.setInd(ind);
        }
        if (instruction.contains("左缩进为0") || instruction.contains("左缩进0") || instruction.contains("左顶格")) {
            PPrBase.Ind ind = pPr.getInd() == null ? new PPrBase.Ind() : pPr.getInd();
            ind.setLeft(BigInteger.ZERO);
            ind.setLeftChars(BigInteger.ZERO);
            pPr.setInd(ind);
        }
        Matcher hanging = Pattern.compile("悬挂缩进\\s*(\\d+(?:\\.\\d+)?)\\s*(?:个)?字符").matcher(instruction.replace(" ", ""));
        if (hanging.find()) {
            PPrBase.Ind ind = pPr.getInd() == null ? new PPrBase.Ind() : pPr.getInd();
            int chars = (int) Math.round(Double.parseDouble(hanging.group(1)) * 100);
            ind.setHangingChars(BigInteger.valueOf(chars));
            ind.setLeftChars(BigInteger.valueOf(chars));
            pPr.setInd(ind);
        }
    }

    private static Style baseStyle(String name, String role) {
        Style result = new Style();
        result.setType("paragraph");
        result.setStyleId(styleId(role));
        result.setCustomStyle(true);
        Style.Name styleName = new Style.Name();
        styleName.setVal(name);
        result.setName(styleName);
        Style.BasedOn basedOn = new Style.BasedOn();
        basedOn.setVal("Normal");
        result.setBasedOn(basedOn);
        return result;
    }

    private static String styleId(String role) {
        return "PaperFormat" + role.substring(0, 1).toUpperCase(Locale.ROOT) + role.substring(1);
    }

    private static BigInteger halfPointsFromText(String text) {
        if (text.contains("小初")) return BigInteger.valueOf(72);
        if (text.contains("初号")) return BigInteger.valueOf(84);
        if (text.contains("小一")) return BigInteger.valueOf(48);
        if (text.contains("一号")) return BigInteger.valueOf(52);
        if (text.contains("小二")) return BigInteger.valueOf(36);
        if (text.contains("二号")) return BigInteger.valueOf(44);
        if (text.contains("小三")) return BigInteger.valueOf(30);
        if (text.contains("三号")) return BigInteger.valueOf(32);
        if (text.contains("小四")) return BigInteger.valueOf(24);
        if (text.contains("四号")) return BigInteger.valueOf(28);
        if (text.contains("小五")) return BigInteger.valueOf(18);
        if (text.contains("五号")) return BigInteger.valueOf(21);
        return null;
    }

    private static PPrBase.Spacing spacingFromText(String text) {
        String compact = text.replace(" ", "");
        PPrBase.Spacing spacing = new PPrBase.Spacing();
        boolean present = false;
        if (compact.contains("行距18磅") || compact.contains("固定值18磅")) {
            spacing.setLine(BigInteger.valueOf(360));
            spacing.setLineRule(org.docx4j.wml.STLineSpacingRule.EXACT);
            present = true;
        } else if (compact.contains("行距24磅")) {
            spacing.setLine(BigInteger.valueOf(480));
            spacing.setLineRule(org.docx4j.wml.STLineSpacingRule.EXACT);
            present = true;
        } else if (compact.contains("单倍行距")) {
            spacing.setLine(BigInteger.valueOf(240));
            spacing.setLineRule(org.docx4j.wml.STLineSpacingRule.AUTO);
            present = true;
        } else if (compact.contains("1.5倍行距") || compact.contains("1.5倍")) {
            spacing.setLine(BigInteger.valueOf(360));
            spacing.setLineRule(org.docx4j.wml.STLineSpacingRule.AUTO);
            present = true;
        } else if (compact.contains("2倍行距") || compact.contains("双倍行距")) {
            spacing.setLine(BigInteger.valueOf(480));
            spacing.setLineRule(org.docx4j.wml.STLineSpacingRule.AUTO);
            present = true;
        } else {
            Matcher exact = Pattern.compile("(?:固定值|行距)(\\d+(?:\\.\\d+)?)磅").matcher(compact);
            if (exact.find()) {
                int twips = (int) Math.round(Double.parseDouble(exact.group(1)) * 20);
                spacing.setLine(BigInteger.valueOf(twips));
                spacing.setLineRule(org.docx4j.wml.STLineSpacingRule.EXACT);
                present = true;
            }
        }
        if (compact.contains("段前、段后各1行")) {
            spacing.setBeforeLines(BigInteger.valueOf(100));
            spacing.setAfterLines(BigInteger.valueOf(100));
            present = true;
        }
        if (compact.contains("段前、段后各0.5行")) {
            spacing.setBeforeLines(BigInteger.valueOf(50));
            spacing.setAfterLines(BigInteger.valueOf(50));
            present = true;
        }
        if (compact.contains("段前1行")) { spacing.setBeforeLines(BigInteger.valueOf(100)); present = true; }
        if (compact.contains("段后1行")) { spacing.setAfterLines(BigInteger.valueOf(100)); present = true; }
        if (compact.contains("段前0.5行")) { spacing.setBeforeLines(BigInteger.valueOf(50)); present = true; }
        if (compact.contains("段后0.5行")) { spacing.setAfterLines(BigInteger.valueOf(50)); present = true; }
        if (compact.contains("段前0行")) { spacing.setBefore(BigInteger.ZERO); present = true; }
        if (compact.contains("段后0行")) { spacing.setAfter(BigInteger.ZERO); present = true; }
        return present ? spacing : null;
    }

    private static List<P> directParagraphs(WordprocessingMLPackage template) {
        List<P> result = new ArrayList<>();
        for (Object item : template.getMainDocumentPart().getJaxbElement().getBody().getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (value instanceof P paragraph) result.add(paragraph);
        }
        return result;
    }

    private static P findSampleParagraph(List<P> paragraphs, String role) {
        if (role.equals("coverTitle")) {
            int marker = findContaining(paragraphs, "毕业设计说明书", 0);
            return nextNonEmpty(paragraphs, marker + 1);
        }
        Map<String, String> annotations = Map.of(
                "title", "中文题目",
                "heading1", "一级标题：",
                "heading2", "二级标题：",
                "heading3", "三级标题：",
                "body", "正文格式："
        );
        String annotation = annotations.get(role);
        if (annotation != null) {
            int index = findContaining(paragraphs, annotation, 0);
            return previousNonEmpty(paragraphs, index - 1);
        }
        if (role.equals("tocTitle")) {
            for (P paragraph : paragraphs) if (normalize(TextUtils.getText(paragraph)).replace(" ", "").equals("目录")) return paragraph;
        }
        if (role.equals("abstract")) {
            for (P paragraph : paragraphs) if (normalize(TextUtils.getText(paragraph)).replace(" ", "").startsWith("摘要：")) return paragraph;
        }
        return null;
    }

    private static int findContaining(List<P> paragraphs, String token, int start) {
        for (int i = Math.max(0, start); i < paragraphs.size(); i++) {
            if (normalize(TextUtils.getText(paragraphs.get(i))).contains(token)) return i;
        }
        return -1;
    }

    private static P nextNonEmpty(List<P> paragraphs, int start) {
        for (int i = Math.max(0, start); i < paragraphs.size(); i++) {
            if (!normalize(TextUtils.getText(paragraphs.get(i))).isBlank()) return paragraphs.get(i);
        }
        return null;
    }

    private static P previousNonEmpty(List<P> paragraphs, int start) {
        for (int i = Math.min(start, paragraphs.size() - 1); i >= 0; i--) {
            if (!normalize(TextUtils.getText(paragraphs.get(i))).isBlank()) return paragraphs.get(i);
        }
        return null;
    }

    private static P findParagraphUsingStyle(List<P> paragraphs, String styleId) {
        if (styleId == null) return null;
        for (P paragraph : paragraphs) {
            if (paragraph.getPPr() != null && paragraph.getPPr().getPStyle() != null
                    && styleId.equals(paragraph.getPPr().getPStyle().getVal())) {
                return paragraph;
            }
        }
        return null;
    }

    private static void mergeDirectParagraphPropertiesIntoStyle(Style style, P sample) {
        if (sample == null || sample.getPPr() == null) return;
        PPr direct = sample.getPPr();
        PPr merged = style.getPPr() == null ? new PPr() : (PPr) XmlUtils.deepCopy(style.getPPr());
        if (direct.getJc() != null) merged.setJc((org.docx4j.wml.Jc) XmlUtils.deepCopy(direct.getJc()));
        if (direct.getSpacing() != null) merged.setSpacing((PPrBase.Spacing) XmlUtils.deepCopy(direct.getSpacing()));
        if (direct.getInd() != null) merged.setInd((PPrBase.Ind) XmlUtils.deepCopy(direct.getInd()));
        if (direct.getTabs() != null) merged.setTabs((org.docx4j.wml.Tabs) XmlUtils.deepCopy(direct.getTabs()));
        style.setPPr(merged);
    }

    private static void forceFirstLine(Style style, int characters) {
        PPr pPr = style.getPPr() == null ? new PPr() : style.getPPr();
        PPrBase.Ind ind = pPr.getInd() == null ? new PPrBase.Ind() : pPr.getInd();
        if (characters == 0) {
            ind.setFirstLine(BigInteger.ZERO);
            ind.setFirstLineChars(BigInteger.ZERO);
        } else {
            ind.setFirstLine(null);
            ind.setFirstLineChars(BigInteger.valueOf(characters));
        }
        pPr.setInd(ind);
        style.setPPr(pPr);
    }

    private static P findSourceSampleParagraph(List<P> paragraphs, String role) {
        for (P paragraph : paragraphs) {
            String text = normalize(TextUtils.getText(paragraph));
            if (text.isBlank()) continue;
            String compact = text.replace(" ", "");
            boolean match = switch (role) {
                case "coverTitle" -> false;
                case "title" -> compact.length() >= 4 && compact.length() <= 60
                        && !compact.matches("^\\d+(?:\\.\\d+){0,3}.*$");
                case "abstract" -> compact.startsWith("摘要") || compact.toLowerCase(Locale.ROOT).startsWith("abstract");
                case "tocTitle" -> compact.equals("目录");
                case "toc1" -> compact.matches("^\\d+[^.].*\\d+$");
                case "toc2" -> compact.matches("^\\d+\\.\\d+.*\\d+$");
                case "heading1" -> compact.matches("^\\d+[^\\d.].*$");
                case "heading2" -> compact.matches("^\\d+\\.\\d+[^.].*$");
                case "heading3" -> compact.matches("^\\d+\\.\\d+\\.\\d+[^.].*$");
                case "heading4" -> compact.matches("^\\d+\\.\\d+\\.\\d+\\.\\d+.*$");
                case "figureCaption" -> compact.matches("^图\\d+[-－]\\d+.*$");
                case "tableCaption", "continuedTableCaption" -> compact.matches("^(续)?表\\d+[-－]\\d+.*$");
                case "references" -> compact.matches("^\\[?\\d+].*$");
                case "sectionTitle" -> compact.equals("参考文献") || compact.equals("致谢") || compact.equals("附录");
                case "body", "thanksBody", "tableFollowingBody", "tableCell" -> text.length() >= 20
                        && !looksLikeInstruction(text)
                        && !compact.matches("^\\d+(?:\\.\\d+){0,3}.*$");
                default -> false;
            };
            if (match) return paragraph;
        }
        if (role.equals("coverTitle")) return nextNonEmpty(paragraphs, 0);
        return null;
    }

    private static P findTableCellSample(WordprocessingMLPackage document) {
        Body body = document.getMainDocumentPart().getJaxbElement().getBody();
        for (Object item : body.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (!(value instanceof Tbl table)) continue;
            String text = normalize(TextUtils.getText(table)).replace(" ", "");
            String xml = XmlUtils.marshaltoString(table, true, true).toLowerCase(Locale.ROOT);
            if (text.contains("专业") && (text.contains("学生姓名") || text.contains("指导教师"))) continue;
            if (xml.contains(":drawing") || xml.contains(":pict") || xml.contains(":imagedata")) continue;
            int rows = 0;
            int maxCells = 0;
            for (Object rowItem : table.getContent()) {
                Object rowValue = XmlUtils.unwrap(rowItem);
                if (!(rowValue instanceof Tr row)) continue;
                rows++;
                int cells = 0;
                for (Object cellItem : row.getContent()) {
                    if (XmlUtils.unwrap(cellItem) instanceof Tc) cells++;
                }
                maxCells = Math.max(maxCells, cells);
            }
            if (rows < 2 || maxCells < 2) continue;
            List<P> paragraphs = new ArrayList<>();
            collectParagraphs(table, paragraphs);
            for (P paragraph : paragraphs) {
                if (!normalize(TextUtils.getText(paragraph)).isBlank()) return paragraph;
            }
        }
        return null;
    }

    private static P findFigureCaptionSample(WordprocessingMLPackage document) {
        List<P> paragraphs = new ArrayList<>();
        collectParagraphs(document.getMainDocumentPart().getJaxbElement().getBody(), paragraphs);
        for (P paragraph : paragraphs) {
            if (normalize(TextUtils.getText(paragraph)).replace(" ", "").matches("^图\\d+[-－]\\d+.*$")) {
                return paragraph;
            }
        }
        return null;
    }

    private static List<String> extractInstructions(WordprocessingMLPackage template) {
        Set<String> result = new LinkedHashSet<>();
        Body body = template.getMainDocumentPart().getJaxbElement().getBody();
        List<P> paragraphs = new ArrayList<>();
        collectParagraphs(body, paragraphs);
        for (P paragraph : paragraphs) {
            String text = normalize(TextUtils.getText(paragraph));
            if (text.length() < 5 || text.length() > 500) continue;
            if (looksLikeInstruction(text)) result.add(text);
        }
        return new ArrayList<>(result);
    }

    private static void collectParagraphs(Object node, List<P> result) {
        Object value = XmlUtils.unwrap(node);
        if (value instanceof P paragraph) {
            result.add(paragraph);
            return;
        }
        if (value instanceof Body body) {
            body.getContent().forEach(item -> collectParagraphs(item, result));
            return;
        }
        if (value instanceof org.docx4j.wml.Tbl table) {
            table.getContent().forEach(item -> collectParagraphs(item, result));
            return;
        }
        if (value instanceof org.docx4j.wml.Tr row) {
            row.getContent().forEach(item -> collectParagraphs(item, result));
            return;
        }
        if (value instanceof org.docx4j.wml.Tc cell) {
            cell.getContent().forEach(item -> collectParagraphs(item, result));
        }
    }

    private static boolean looksLikeInstruction(String text) {
        String value = text.replace(" ", "");
        return value.contains("规定用")
                || value.contains("正文格式")
                || value.contains("标题：")
                || value.contains("标题:")
                || value.contains("目录只显示")
                || value.contains("目录是生成的")
                || value.contains("页边距")
                || value.contains("页眉")
                || value.contains("页脚")
                || value.contains("行距")
                || value.contains("首行")
                || value.contains("左缩进")
                || value.contains("字体")
                || value.contains("字号")
                || value.contains("黑体")
                || value.contains("宋体")
                || value.toLowerCase(Locale.ROOT).contains("timesnewroman")
                || value.contains("居中") && (value.contains("字体") || value.contains("字号"));
    }

    private static Map<String, String> mapInstructionsToRoles(List<String> instructions) {
        Map<String, String> result = new LinkedHashMap<>();
        for (String text : instructions) {
            String compact = text.replace(" ", "");
            if (compact.contains("中文题目") || compact.contains("英文题目")) putEvidence(result, "title", text);
            if (compact.contains("摘要") && (compact.contains("行距") || compact.contains("首行"))) putEvidence(result, "abstract", text);
            if (compact.contains("目录只显示") || compact.contains("目录是生成的")) {
                putEvidence(result, "toc1", text);
                putEvidence(result, "toc2", text);
            }
            if (compact.contains("文本组成部分的题目") || compact.contains("参考文献") && compact.contains("居中")) {
                putEvidence(result, "tocTitle", text);
                putEvidence(result, "sectionTitle", text);
            }
            if (compact.contains("第一级标题") || compact.contains("一级标题：")) putEvidence(result, "heading1", text);
            if (compact.contains("第二级标题") || compact.contains("二级标题：")) putEvidence(result, "heading2", text);
            if (compact.contains("第三级标题") || compact.contains("三级标题：")) putEvidence(result, "heading3", text);
            if (compact.contains("第四") && compact.contains("标题")) putEvidence(result, "heading4", text);
            if (compact.contains("正文格式") || compact.contains("正文和第四") || compact.contains("正文使用")) {
                putEvidence(result, "body", text);
                putEvidence(result, "thanksBody", text);
            }
            if (compact.contains("表格序号") || compact.contains("表格题目")) putEvidence(result, "tableCaption", text);
            if (compact.contains("图号") || compact.contains("图题") || compact.contains("插图") && compact.contains("居中")) {
                putEvidence(result, "figureCaption", text);
            }
            if (compact.contains("表格内") || compact.contains("表内") || compact.contains("表格正文")
                    || compact.contains("表中文字") || compact.contains("表格内容")) {
                putEvidence(result, "tableCell", text);
            }
            if (compact.contains("表格后面") || compact.contains("表格后的第一行")) putEvidence(result, "tableFollowingBody", text);
            if (compact.contains("参考文献列表") || compact.contains("参考文献")
                    && (compact.contains("悬挂") || compact.contains("标点") || compact.contains("编号"))) {
                putEvidence(result, "references", text);
            }
        }
        return result;
    }

    static String detectMode(WordprocessingMLPackage template) throws Exception {
        List<String> instructions = extractInstructions(template);
        return detectMode(template, instructions, stylesByName(template));
    }

    private static String detectMode(
            WordprocessingMLPackage template,
            List<String> instructions,
            Map<String, Style> styles
    ) {
        int canonical = 0;
        for (String name : CANONICAL_STYLES.values()) {
            if (styles.containsKey(name.toLowerCase(Locale.ROOT))) canonical++;
        }
        if (!instructions.isEmpty() && canonical < 3) return "TEXT_INSTRUCTIONS_ONLY";
        if (!instructions.isEmpty()) return "SAMPLE_AND_INSTRUCTIONS";
        return "SAMPLE_OR_STYLE";
    }

    private static String modeDescription(String mode) {
        return switch (mode) {
            case "TEXT_INSTRUCTIONS_ONLY" -> "纯文字要求模式：先识别待修改文档结构，再把文字要求转换为结构化格式规则。";
            case "SAMPLE_AND_INSTRUCTIONS" -> "样例与文字要求混合模式：样例格式优先，文字要求用于补充缺失规则。";
            default -> "样例或样式模式：从模板的命名样式和示例组件提取格式。";
        };
    }

    private static void putEvidence(Map<String, String> target, String role, String text) {
        target.merge(role, text, (left, right) -> left.contains(right) ? left : left + " | " + right);
    }

    private static String executionInstruction(String role, String evidence) {
        List<String> markers;
        if (role.equals("abstract")) {
            markers = List.of("摘要内容", "摘要正文", "正文内容");
        } else if (role.equals("tableCell")) {
            markers = List.of("表格内容", "表内文字", "表格内文字", "表中文字");
        } else {
            return evidence;
        }
        int offset = -1;
        for (String marker : markers) {
            int index = evidence.lastIndexOf(marker);
            if (index >= 0) offset = Math.max(offset, index);
        }
        return offset < 0 ? evidence : evidence.substring(offset);
    }

    private static Map<String, Style> stylesByName(WordprocessingMLPackage template) throws Exception {
        StyleDefinitionsPart part = template.getMainDocumentPart().getStyleDefinitionsPart();
        Map<String, Style> result = new LinkedHashMap<>();
        for (Style style : part.getJaxbElement().getStyle()) {
            String name = styleName(style);
            if (name != null) result.put(name.toLowerCase(Locale.ROOT), style);
        }
        return result;
    }

    private static Style findStyle(Map<String, Style> styles, List<String> candidates) {
        for (String candidate : candidates) {
            Style exact = styles.get(candidate.toLowerCase(Locale.ROOT));
            if (exact != null) return exact;
        }
        for (String candidate : candidates) {
            String needle = candidate.toLowerCase(Locale.ROOT);
            for (Map.Entry<String, Style> entry : styles.entrySet()) {
                if (entry.getKey().contains(needle)) return entry.getValue();
            }
        }
        return null;
    }

    private static String styleName(Style style) {
        if (style == null) return null;
        return style.getName() == null ? style.getStyleId() : style.getName().getVal();
    }

    private static String describeFonts(RPr rPr) {
        if (rPr == null || rPr.getRFonts() == null) return "模板继承字体";
        Set<String> names = new LinkedHashSet<>();
        add(names, rPr.getRFonts().getEastAsia());
        add(names, rPr.getRFonts().getAscii());
        add(names, rPr.getRFonts().getHAnsi());
        return names.isEmpty() ? "模板主题字体" : String.join(" / ", names);
    }

    private static boolean hasExplicitFont(RPr rPr) {
        if (rPr == null || rPr.getRFonts() == null) return false;
        RFonts fonts = rPr.getRFonts();
        return fonts.getEastAsia() != null || fonts.getAscii() != null
                || fonts.getHAnsi() != null || fonts.getCs() != null;
    }

    private static boolean hasExplicitParagraphFormatting(PPr pPr) {
        return pPr != null && (pPr.getSpacing() != null || pPr.getInd() != null || pPr.getJc() != null);
    }

    private static String describeSize(RPr rPr) {
        if (rPr == null || rPr.getSz() == null || rPr.getSz().getVal() == null) return "模板继承字号";
        BigDecimal points = new BigDecimal(rPr.getSz().getVal()).divide(BigDecimal.valueOf(2), 2, RoundingMode.HALF_UP).stripTrailingZeros();
        boolean bold = rPr.getB() != null && rPr.getB().isVal();
        return points.toPlainString() + " pt" + (bold ? "，加粗" : "");
    }

    private static String describeParagraph(PPr pPr) {
        if (pPr == null) return "模板继承段落格式";
        List<String> values = new ArrayList<>();
        if (pPr.getJc() != null) values.add("对齐=" + pPr.getJc().getVal());
        PPrBase.Spacing spacing = pPr.getSpacing();
        if (spacing != null) {
            if (spacing.getBefore() != null) values.add("段前=" + twipsToPoints(spacing.getBefore()) + " pt");
            if (spacing.getBeforeLines() != null) values.add("段前=" + lineHundredths(spacing.getBeforeLines()) + " 行");
            if (spacing.getAfter() != null) values.add("段后=" + twipsToPoints(spacing.getAfter()) + " pt");
            if (spacing.getAfterLines() != null) values.add("段后=" + lineHundredths(spacing.getAfterLines()) + " 行");
            if (spacing.getLine() != null) {
                String unit = spacing.getLineRule() == null || spacing.getLineRule().toString().equals("AUTO")
                        ? lineMultiple(spacing.getLine()) + " 倍"
                        : twipsToPoints(spacing.getLine()) + " pt";
                values.add("行距=" + unit + (spacing.getLineRule() == null ? "" : " (" + spacing.getLineRule() + ")"));
            }
        }
        PPrBase.Ind ind = pPr.getInd();
        if (ind != null) {
            if (ind.getFirstLineChars() != null) values.add("首行缩进=" + lineHundredths(ind.getFirstLineChars()) + " 字符");
            else if (ind.getFirstLine() != null) values.add("首行缩进=" + twipsToPoints(ind.getFirstLine()) + " pt");
            if (ind.getLeftChars() != null) values.add("左缩进=" + lineHundredths(ind.getLeftChars()) + " 字符");
            else if (ind.getLeft() != null) values.add("左缩进=" + twipsToPoints(ind.getLeft()) + " pt");
        }
        if (pPr.getKeepNext() != null && pPr.getKeepNext().isVal()) values.add("与下段同页");
        if (pPr.getPageBreakBefore() != null && pPr.getPageBreakBefore().isVal()) values.add("段前分页");
        return values.isEmpty() ? "模板继承段落格式" : String.join("；", values);
    }

    private static String fontsFromText(String text) {
        List<String> result = new ArrayList<>();
        if (text.contains("黑体")) result.add("黑体");
        if (text.contains("宋体")) result.add("宋体");
        if (text.toLowerCase(Locale.ROOT).contains("times new roman")) result.add("Times New Roman");
        return result.isEmpty() ? "文字说明未明确字体" : String.join(" / ", result);
    }

    private static String sizeFromText(String text) {
        String[][] sizes = {
                {"小初", "36 pt"}, {"初号", "42 pt"}, {"小一", "24 pt"}, {"一号", "26 pt"},
                {"小二", "18 pt"}, {"二号", "22 pt"}, {"小三", "15 pt"}, {"三号", "16 pt"},
                {"小四", "12 pt"}, {"四号", "14 pt"}, {"小五", "9 pt"}, {"五号", "10.5 pt"}
        };
        for (String[] entry : sizes) {
            if (text.contains(entry[0])) return entry[0] + "（" + entry[1] + "）" + (text.contains("加粗") ? "，加粗" : "");
        }
        return "文字说明未明确字号";
    }

    private static String paragraphFromText(String text) {
        List<String> values = new ArrayList<>();
        if (text.contains("居中")) values.add("居中");
        if (text.contains("两端对齐")) values.add("两端对齐");
        if (text.contains("左顶格") || text.contains("左缩进为0")) values.add("左缩进0");
        if (text.contains("首行") && text.contains("2个字符")) values.add("首行缩进2字符");
        if (text.contains("18磅")) values.add("行距18磅");
        if (text.contains("24磅")) values.add("行距24磅");
        if (text.contains("单倍行距")) values.add("单倍行距");
        if (text.contains("段前1行")) values.add("段前1行");
        if (text.contains("段后1行")) values.add("段后1行");
        if (text.contains("段前0.5行")) values.add("段前0.5行");
        if (text.contains("段后0.5行")) values.add("段后0.5行");
        return values.isEmpty() ? text : String.join("；", values);
    }

    private static SectPr firstSection(WordprocessingMLPackage document) {
        Body body = document.getMainDocumentPart().getJaxbElement().getBody();
        for (Object item : body.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (value instanceof P paragraph && paragraph.getPPr() != null && paragraph.getPPr().getSectPr() != null) {
                return paragraph.getPPr().getSectPr();
            }
        }
        return body.getSectPr();
    }

    private static String describePageSize(SectPr section) {
        if (section == null || section.getPgSz() == null) return "模板未声明页面尺寸";
        double width = twipsToCm(section.getPgSz().getW());
        double height = twipsToCm(section.getPgSz().getH());
        String orientation = section.getPgSz().getOrient() == null ? (width <= height ? "纵向" : "横向") : section.getPgSz().getOrient().toString();
        return String.format(Locale.ROOT, "%.2f cm × %.2f cm，%s", width, height, orientation);
    }

    private static Map<String, Double> describeMargins(SectPr section) {
        Map<String, Double> result = new LinkedHashMap<>();
        if (section == null || section.getPgMar() == null) return result;
        result.put("top", twipsToCm(section.getPgMar().getTop()));
        result.put("bottom", twipsToCm(section.getPgMar().getBottom()));
        result.put("left", twipsToCm(section.getPgMar().getLeft()));
        result.put("right", twipsToCm(section.getPgMar().getRight()));
        return result;
    }

    private static Map<String, Double> describeHeaderFooterDistances(SectPr section) {
        Map<String, Double> result = new LinkedHashMap<>();
        if (section == null || section.getPgMar() == null) return result;
        result.put("header", twipsToCm(section.getPgMar().getHeader()));
        result.put("footer", twipsToCm(section.getPgMar().getFooter()));
        return result;
    }

    private static double twipsToCm(BigInteger value) {
        if (value == null) return 0.0;
        return BigDecimal.valueOf(value.doubleValue() * 2.54 / 1440.0).setScale(3, RoundingMode.HALF_UP).doubleValue();
    }

    private static String twipsToPoints(BigInteger value) {
        return new BigDecimal(value).divide(BigDecimal.valueOf(20), 2, RoundingMode.HALF_UP).stripTrailingZeros().toPlainString();
    }

    private static String lineHundredths(BigInteger value) {
        return new BigDecimal(value).divide(BigDecimal.valueOf(100), 2, RoundingMode.HALF_UP).stripTrailingZeros().toPlainString();
    }

    private static String lineMultiple(BigInteger value) {
        return new BigDecimal(value).divide(BigDecimal.valueOf(240), 2, RoundingMode.HALF_UP).stripTrailingZeros().toPlainString();
    }

    private static void add(Set<String> target, String value) {
        if (value != null && !value.isBlank()) target.add(value);
    }

    private static String normalize(String text) {
        if (text == null) return "";
        return text.replace('\u00A0', ' ').replaceAll("[\\r\\n\\t]+", " ").replaceAll("\\s+", " ").trim();
    }

    private record RoleDefinition(String label, List<String> styleCandidates) {
    }
}
