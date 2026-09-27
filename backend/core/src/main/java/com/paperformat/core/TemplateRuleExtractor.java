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
        StyleDefinitionsPart part = template.getMainDocumentPart().getStyleDefinitionsPart();
        Map<String, Style> styles = stylesByName(template);
        List<P> paragraphs = directParagraphs(template);
        Map<String, String> instructionByRole = mapInstructionsToRoles(extractInstructions(template));
        PropertyResolver resolver = new PropertyResolver(template);

        for (Map.Entry<String, String> entry : CANONICAL_STYLES.entrySet()) {
            String role = entry.getKey();
            String canonicalName = entry.getValue();
            Style existing = styles.get(canonicalName.toLowerCase(Locale.ROOT));
            if (existing != null) {
                // 同名样式是最高优先级证据。文字说明只用于解释或补齐缺失样式，
                // 不能反过来覆盖模板样式中的缩进、段距和制表位。
                continue;
            }

            RoleDefinition definition = ROLES.get(role);
            Style alias = definition == null ? null : findStyle(styles, definition.styleCandidates());
            Style created;
            if (alias != null) {
                created = cloneAs(alias, canonicalName, role);
            } else {
                P sample = findSampleParagraph(paragraphs, role);
                if (sample != null) {
                    created = styleFromSample(sample, canonicalName, role, resolver);
                } else {
                    String instruction = instructionByRole.get(role);
                    if (instruction == null && role.equals("thanksBody")) instruction = instructionByRole.get("body");
                    if (instruction == null && role.equals("heading4")) instruction = instructionByRole.get("body");
                    if (instruction == null) continue;
                    created = styleFromInstruction(canonicalName, role, instruction);
                }
            }
            part.getJaxbElement().getStyle().add(created);
            styles.put(canonicalName.toLowerCase(Locale.ROOT), created);
        }
    }

    static ProcessingReport.TemplateRules extract(WordprocessingMLPackage template) throws Exception {
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
                rules.put(key, new ProcessingReport.StyleRule(
                        definition.label(),
                        styleName(style),
                        describeFonts(rPr),
                        describeSize(rPr),
                        describeParagraph(pPr),
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
        return new ProcessingReport.TemplateRules(
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
        PPr pPr = sample.getPPr() == null ? new PPr() : (PPr) XmlUtils.deepCopy(sample.getPPr());
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
        if (rPr == null) {
            rPr = resolver.getEffectiveRPr(null, sample.getPPr());
        }
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

    private static List<String> extractInstructions(WordprocessingMLPackage template) {
        Set<String> result = new LinkedHashSet<>();
        Body body = template.getMainDocumentPart().getJaxbElement().getBody();
        for (Object item : body.getContent()) {
            Object value = XmlUtils.unwrap(item);
            if (!(value instanceof P paragraph)) continue;
            String text = normalize(TextUtils.getText(paragraph));
            if (text.length() < 8 || text.length() > 260) continue;
            if (looksLikeInstruction(text)) result.add(text);
        }
        return new ArrayList<>(result);
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
            if (compact.contains("表格后面") || compact.contains("表格后的第一行")) putEvidence(result, "tableFollowingBody", text);
            if (compact.contains("参考文献列表")) putEvidence(result, "references", text);
        }
        return result;
    }

    private static void putEvidence(Map<String, String> target, String role, String text) {
        target.merge(role, text, (left, right) -> left.contains(right) ? left : left + " | " + right);
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
        Map<String, String> sizes = Map.of(
                "小初", "36 pt", "三号", "16 pt", "四号", "14 pt",
                "小四", "12 pt", "五号", "10.5 pt", "小五", "9 pt"
        );
        for (Map.Entry<String, String> entry : sizes.entrySet()) {
            if (text.contains(entry.getKey())) return entry.getKey() + "（" + entry.getValue() + "）" + (text.contains("加粗") ? "，加粗" : "");
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
