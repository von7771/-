package com.papercheck.service;

import com.papercheck.model.CheckResult;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

/**
 * 论文格式检查服务
 */
public class PaperFormatChecker {
    private static final Logger logger = LoggerFactory.getLogger(PaperFormatChecker.class);
    
    // 定义论文格式规范
    private static final int TITLE_FONT_SIZE = 16; // 三号字约16pt
    private static final String TITLE_FONT_FAMILY = "黑体";
    private static final int BODY_FONT_SIZE = 12; // 小四号字约12pt
    private static final String BODY_FONT_FAMILY = "宋体";
    private static final double LINE_SPACING = 1.5;
    private static final int FIRST_LINE_INDENT = 2; // 首行缩进2字符
    private static final double MARGIN_TOP = 2.5; // 上边距2.5厘米
    private static final double MARGIN_BOTTOM = 2.5; // 下边距2.5厘米
    private static final double MARGIN_LEFT = 3.0; // 左边距3.0厘米
    private static final double MARGIN_RIGHT = 3.0; // 右边距3.0厘米

    /**
     * 检查Word文档格式
     *
     * @param file Word文档文件
     * @return 检查结果列表
     * @throws IOException 如果文件读取失败
     */
    public List<CheckResult> checkDocument(File file) throws IOException {
        logger.info("开始检查文档: {}", file.getName());
        List<CheckResult> results = new ArrayList<>();
        
        try (FileInputStream fis = new FileInputStream(file)) {
            XWPFDocument document = new XWPFDocument(fis);
            
            // 检查文档属性
            checkDocumentProperties(document, results);
            
            // 检查段落格式
            checkParagraphs(document, results);
            
            // 检查表格格式
            checkTables(document, results);
            
            // 检查图片格式
            checkPictures(document, results);
            
            logger.info("文档检查完成，发现 {} 个问题", results.size());
            return results;
        } catch (Exception e) {
            logger.error("检查文档时发生错误", e);
            throw new IOException("检查文档时发生错误: " + e.getMessage(), e);
        }
    }
    
    /**
     * 创建修复后的文档预览
     * 
     * @param originalFile 原始文档
     * @param previewFile 预览文档保存路径
     * @return 修复的问题列表
     * @throws IOException 如果文件操作失败
     */
    public List<CheckResult> createFixedDocumentPreview(File originalFile, File previewFile) throws IOException {
        logger.info("创建修复后的文档预览: {}", originalFile.getName());
        List<CheckResult> fixedResults = new ArrayList<>();
        
        try (FileInputStream fis = new FileInputStream(originalFile)) {
            XWPFDocument document = new XWPFDocument(fis);
            
            // 修复文档属性
            fixDocumentProperties(document, fixedResults);
            
            // 修复段落格式
            fixParagraphs(document, fixedResults);
            
            // 修复表格格式
            fixTables(document, fixedResults);
            
            // 修复图片格式
            fixPictures(document, fixedResults);
            
            // 保存预览文档
            try (FileOutputStream fos = new FileOutputStream(previewFile)) {
                document.write(fos);
            }
            
            logger.info("文档修复预览完成，修复了 {} 个问题", fixedResults.size());
            return fixedResults;
        } catch (Exception e) {
            logger.error("创建修复预览时发生错误", e);
            throw new IOException("创建修复预览时发生错误: " + e.getMessage(), e);
        }
    }
    
    /**
     * 应用修复到原始文档
     * 
     * @param originalFile 原始文档
     * @param fixedFile 修复后的文档保存路径
     * @return 修复的问题列表
     * @throws IOException 如果文件操作失败
     */
    public List<CheckResult> fixDocument(File originalFile, File fixedFile) throws IOException {
        logger.info("修复文档: {}", originalFile.getName());
        List<CheckResult> fixedResults = new ArrayList<>();
        
        try (FileInputStream fis = new FileInputStream(originalFile)) {
            XWPFDocument document = new XWPFDocument(fis);
            
            // 修复文档属性
            fixDocumentProperties(document, fixedResults);
            
            // 修复段落格式
            fixParagraphs(document, fixedResults);
            
            // 修复表格格式
            fixTables(document, fixedResults);
            
            // 修复图片格式
            fixPictures(document, fixedResults);
            
            // 保存修复后的文档
            try (FileOutputStream fos = new FileOutputStream(fixedFile)) {
                document.write(fos);
            }
            
            logger.info("文档修复完成，修复了 {} 个问题", fixedResults.size());
            return fixedResults;
        } catch (Exception e) {
            logger.error("修复文档时发生错误", e);
            throw new IOException("修复文档时发生错误: " + e.getMessage(), e);
        }
    }

    /**
     * 修复文档属性（页边距、页码等）
     */
    private void fixDocumentProperties(XWPFDocument document, List<CheckResult> fixedResults) {
        logger.debug("修复文档属性");
        
        // 修复页边距
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr == null) {
            sectPr = document.getDocument().getBody().addNewSectPr();
        }
        
        // 设置页边距（Word中的单位是dxa，1厘米约等于567 dxa）
        int dxaPerCm = 567;
        CTPageMar pgMar = sectPr.isSetPgMar() ? sectPr.getPgMar() : sectPr.addNewPgMar();
        
        // 设置上边距
        BigInteger topMargin = BigInteger.valueOf((int)(MARGIN_TOP * dxaPerCm));
        pgMar.setTop(topMargin);
        fixedResults.add(new CheckResult(
            "页边距",
            "文档属性",
            "已修复上边距",
            "已设置为" + MARGIN_TOP + "厘米"
        ));
        
        // 设置下边距
        BigInteger bottomMargin = BigInteger.valueOf((int)(MARGIN_BOTTOM * dxaPerCm));
        pgMar.setBottom(bottomMargin);
        fixedResults.add(new CheckResult(
            "页边距",
            "文档属性",
            "已修复下边距",
            "已设置为" + MARGIN_BOTTOM + "厘米"
        ));
        
        // 设置左边距
        BigInteger leftMargin = BigInteger.valueOf((int)(MARGIN_LEFT * dxaPerCm));
        pgMar.setLeft(leftMargin);
        fixedResults.add(new CheckResult(
            "页边距",
            "文档属性",
            "已修复左边距",
            "已设置为" + MARGIN_LEFT + "厘米"
        ));
        
        // 设置右边距
        BigInteger rightMargin = BigInteger.valueOf((int)(MARGIN_RIGHT * dxaPerCm));
        pgMar.setRight(rightMargin);
        fixedResults.add(new CheckResult(
            "页边距",
            "文档属性",
            "已修复右边距",
            "已设置为" + MARGIN_RIGHT + "厘米"
        ));
        
        // 设置页码
        if (!sectPr.isSetPgNumType()) {
            CTPageNumber pgNum = sectPr.addNewPgNumType();
            pgNum.setStart(BigInteger.valueOf(1));
            fixedResults.add(new CheckResult(
                "页码",
                "文档属性",
                "已添加页码",
                "已设置页码从1开始"
            ));
        }
    }

    /**
     * 修复段落格式（字体、行距、缩进等）
     */
    private void fixParagraphs(XWPFDocument document, List<CheckResult> fixedResults) {
        logger.debug("修复段落格式");
        
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        boolean foundTitle = false;
        int paragraphIndex = 0;
        
        for (XWPFParagraph paragraph : paragraphs) {
            paragraphIndex++;
            String paragraphText = paragraph.getText().trim();
            
            // 跳过空段落
            if (paragraphText.isEmpty()) {
                continue;
            }
            
            // 假设第一个非空段落是标题
            if (!foundTitle) {
                foundTitle = true;
                fixTitleFormat(paragraph, paragraphIndex, fixedResults);
            } else {
                // 修复正文段落格式
                fixBodyParagraphFormat(paragraph, paragraphIndex, fixedResults);
            }
        }
    }

    /**
     * 修复标题格式
     */
    private void fixTitleFormat(XWPFParagraph paragraph, int paragraphIndex, List<CheckResult> fixedResults) {
        logger.debug("修复标题格式: 第{}段落", paragraphIndex);
        
        // 修复标题对齐方式
        if (paragraph.getAlignment() != ParagraphAlignment.CENTER) {
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            fixedResults.add(new CheckResult(
                "标题格式",
                "第" + paragraphIndex + "段落",
                "已修复标题对齐方式",
                "已设置为居中对齐"
            ));
        }
        
        // 修复标题字体
        for (XWPFRun run : paragraph.getRuns()) {
            // 修复字体大小
            if (run.getFontSize() != -1 && run.getFontSize() < TITLE_FONT_SIZE) {
                run.setFontSize(TITLE_FONT_SIZE);
                fixedResults.add(new CheckResult(
                    "标题格式",
                    "第" + paragraphIndex + "段落",
                    "已修复标题字号",
                    "已设置为三号字（" + TITLE_FONT_SIZE + "磅）"
                ));
            }
            
            // 修复字体类型
            String fontFamily = run.getFontFamily();
            if (fontFamily == null || !fontFamily.contains(TITLE_FONT_FAMILY)) {
                run.setFontFamily(TITLE_FONT_FAMILY);
                fixedResults.add(new CheckResult(
                    "标题格式",
                    "第" + paragraphIndex + "段落",
                    "已修复标题字体",
                    "已设置为" + TITLE_FONT_FAMILY
                ));
            }
            
            // 修复是否加粗
            if (!run.isBold()) {
                run.setBold(true);
                fixedResults.add(new CheckResult(
                    "标题格式",
                    "第" + paragraphIndex + "段落",
                    "已修复标题加粗",
                    "已设置为加粗"
                ));
            }
        }
    }

    /**
     * 修复正文段落格式
     */
    private void fixBodyParagraphFormat(XWPFParagraph paragraph, int paragraphIndex, List<CheckResult> fixedResults) {
        logger.debug("修复正文格式: 第{}段落", paragraphIndex);
        
        // 修复段落缩进
        CTP ctp = paragraph.getCTP();
        CTPPr pPr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
        CTInd ind = pPr.isSetInd() ? pPr.getInd() : pPr.addNewInd();
        
        // 设置首行缩进为2字符
        int requiredIndent = 420; // 约等于2个字符的缩进
        ind.setFirstLine(BigInteger.valueOf(requiredIndent));
        fixedResults.add(new CheckResult(
            "段落格式",
            "第" + paragraphIndex + "段落",
            "已修复段落首行缩进",
            "已设置为2字符缩进"
        ));
        
        // 修复行距
        CTSpacing spacing = pPr.isSetSpacing() ? pPr.getSpacing() : pPr.addNewSpacing();
        // 设置1.5倍行距
        spacing.setLine(BigInteger.valueOf(360));
        spacing.setLineRule(STLineSpacingRule.AUTO);
        fixedResults.add(new CheckResult(
            "段落格式",
            "第" + paragraphIndex + "段落",
            "已修复段落行距",
            "已设置为1.5倍行距"
        ));
        
        // 修复字体
        for (XWPFRun run : paragraph.getRuns()) {
            // 修复字体大小
            if (run.getFontSize() != -1 && run.getFontSize() > BODY_FONT_SIZE + 2) {
                run.setFontSize(BODY_FONT_SIZE);
                fixedResults.add(new CheckResult(
                    "正文格式",
                    "第" + paragraphIndex + "段落",
                    "已修复正文字号",
                    "已设置为小四号字（" + BODY_FONT_SIZE + "磅）"
                ));
            }
            
            // 修复字体类型
            String fontFamily = run.getFontFamily();
            if (fontFamily == null || !fontFamily.contains(BODY_FONT_FAMILY)) {
                run.setFontFamily(BODY_FONT_FAMILY);
                fixedResults.add(new CheckResult(
                    "正文格式",
                    "第" + paragraphIndex + "段落",
                    "已修复正文字体",
                    "已设置为" + BODY_FONT_FAMILY
                ));
            }
        }
    }

    /**
     * 修复表格格式
     */
    private void fixTables(XWPFDocument document, List<CheckResult> fixedResults) {
        logger.debug("修复表格格式");
        
        List<XWPFTable> tables = document.getTables();
        int tableIndex = 0;
        
        for (XWPFTable table : tables) {
            tableIndex++;
            
            // 修复表格标题
            XWPFTableRow firstRow = table.getRow(0);
            if (firstRow != null) {
                for (XWPFTableCell cell : firstRow.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        if (paragraph.getAlignment() != ParagraphAlignment.CENTER) {
                            paragraph.setAlignment(ParagraphAlignment.CENTER);
                            fixedResults.add(new CheckResult(
                                "表格格式",
                                "表格" + tableIndex,
                                "已修复表格标题对齐方式",
                                "已设置为居中对齐"
                            ));
                        }
                    }
                }
            }
            
            // 修复表格内容字体
            for (int i = 0; i < table.getNumberOfRows(); i++) {
                XWPFTableRow row = table.getRow(i);
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        for (XWPFRun run : paragraph.getRuns()) {
                            // 修复表格内容字体大小
                            if (run.getFontSize() != -1 && run.getFontSize() > BODY_FONT_SIZE + 2) {
                                run.setFontSize(BODY_FONT_SIZE);
                                fixedResults.add(new CheckResult(
                                    "表格格式",
                                    "表格" + tableIndex + "，第" + (i + 1) + "行",
                                    "已修复表格内容字号",
                                    "已设置为小四号字（" + BODY_FONT_SIZE + "磅）"
                                ));
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * 修复图片格式
     */
    private void fixPictures(XWPFDocument document, List<CheckResult> fixedResults) {
        logger.debug("修复图片格式");
        
        // 修复图片标题段落
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        int paragraphIndex = 0;
        
        for (XWPFParagraph paragraph : paragraphs) {
            paragraphIndex++;
            String text = paragraph.getText().trim();
            
            // 识别图片标题（通常以"图"开头）
            if (text.startsWith("图") && text.contains("：")) {
                if (paragraph.getAlignment() != ParagraphAlignment.CENTER) {
                    paragraph.setAlignment(ParagraphAlignment.CENTER);
                    fixedResults.add(new CheckResult(
                        "图片格式",
                        "第" + paragraphIndex + "段落",
                        "已修复图片标题对齐方式",
                        "已设置为居中对齐"
                    ));
                }
            }
        }
    }

    /**
     * 检查文档属性（页边距、页码等）
     */
    private void checkDocumentProperties(XWPFDocument document, List<CheckResult> results) {
        logger.debug("检查文档属性");
        
        // 检查页边距
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr != null) {
            // 检查页边距（Word中的单位是dxa，1厘米约等于567 dxa）
            int dxaPerCm = 567;
            double marginTopRequired = MARGIN_TOP * dxaPerCm;
            double marginBottomRequired = MARGIN_BOTTOM * dxaPerCm;
            double marginLeftRequired = MARGIN_LEFT * dxaPerCm;
            double marginRightRequired = MARGIN_RIGHT * dxaPerCm;
            
            if (sectPr.getPgMar() != null) {
                CTPageMar pgMar = sectPr.getPgMar();
                
                try {
                    // 上边距 - 安全地获取数值
                    double topMargin = 0;
                    if (pgMar.getTop() != null) {
                        topMargin = Double.parseDouble(pgMar.getTop().toString());
                    }
                    
                    if (topMargin < marginTopRequired * 0.9) {
                        results.add(new CheckResult(
                            "页边距",
                            "文档属性",
                            "上边距不符合要求，当前值小于" + MARGIN_TOP + "厘米",
                            "将上边距设置为" + MARGIN_TOP + "厘米"
                        ));
                    }
                    
                    // 下边距
                    double bottomMargin = 0;
                    if (pgMar.getBottom() != null) {
                        bottomMargin = Double.parseDouble(pgMar.getBottom().toString());
                    }
                    
                    if (bottomMargin < marginBottomRequired * 0.9) {
                        results.add(new CheckResult(
                            "页边距",
                            "文档属性",
                            "下边距不符合要求，当前值小于" + MARGIN_BOTTOM + "厘米",
                            "将下边距设置为" + MARGIN_BOTTOM + "厘米"
                        ));
                    }
                    
                    // 左边距
                    double leftMargin = 0;
                    if (pgMar.getLeft() != null) {
                        leftMargin = Double.parseDouble(pgMar.getLeft().toString());
                    }
                    
                    if (leftMargin < marginLeftRequired * 0.9) {
                        results.add(new CheckResult(
                            "页边距",
                            "文档属性",
                            "左边距不符合要求，当前值小于" + MARGIN_LEFT + "厘米",
                            "将左边距设置为" + MARGIN_LEFT + "厘米"
                        ));
                    }
                    
                    // 右边距
                    double rightMargin = 0;
                    if (pgMar.getRight() != null) {
                        rightMargin = Double.parseDouble(pgMar.getRight().toString());
                    }
                    
                    if (rightMargin < marginRightRequired * 0.9) {
                        results.add(new CheckResult(
                            "页边距",
                            "文档属性",
                            "右边距不符合要求，当前值小于" + MARGIN_RIGHT + "厘米",
                            "将右边距设置为" + MARGIN_RIGHT + "厘米"
                        ));
                    }
                } catch (NumberFormatException e) {
                    logger.warn("解析页边距时出错", e);
                    results.add(new CheckResult(
                        "页边距",
                        "文档属性",
                        "无法解析页边距值",
                        "请手动检查页边距设置"
                    ));
                }
            } else {
                results.add(new CheckResult(
                    "页边距",
                    "文档属性",
                    "未设置页边距",
                    "设置页边距：上下" + MARGIN_TOP + "厘米，左右" + MARGIN_LEFT + "厘米"
                ));
            }
            
            // 检查页码位置
            if (sectPr.isSetPgNumType()) {
                // 页码检查逻辑
            } else {
                results.add(new CheckResult(
                    "页码",
                    "文档属性",
                    "未设置页码",
                    "在页面底部居中添加页码"
                ));
            }
        }
    }

    /**
     * 检查段落格式（字体、行距、缩进等）
     */
    private void checkParagraphs(XWPFDocument document, List<CheckResult> results) {
        logger.debug("检查段落格式");
        
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        boolean foundTitle = false;
        int paragraphIndex = 0;
        
        for (XWPFParagraph paragraph : paragraphs) {
            paragraphIndex++;
            String paragraphText = paragraph.getText().trim();
            
            // 跳过空段落
            if (paragraphText.isEmpty()) {
                continue;
            }
            
            // 假设第一个非空段落是标题
            if (!foundTitle) {
                foundTitle = true;
                checkTitleFormat(paragraph, paragraphIndex, results);
            } else {
                // 检查正文段落格式
                checkBodyParagraphFormat(paragraph, paragraphIndex, results);
            }
        }
    }

    /**
     * 检查标题格式
     */
    private void checkTitleFormat(XWPFParagraph paragraph, int paragraphIndex, List<CheckResult> results) {
        logger.debug("检查标题格式: 第{}段落", paragraphIndex);
        
        // 检查标题对齐方式
        if (paragraph.getAlignment() != ParagraphAlignment.CENTER) {
            results.add(new CheckResult(
                "标题格式",
                "第" + paragraphIndex + "段落",
                "标题未居中对齐",
                "将标题设置为居中对齐"
            ));
        }
        
        // 检查标题字体
        for (XWPFRun run : paragraph.getRuns()) {
            // 检查字体大小
            if (run.getFontSize() != -1 && run.getFontSize() < TITLE_FONT_SIZE) {
                results.add(new CheckResult(
                    "标题格式",
                    "第" + paragraphIndex + "段落",
                    "标题字号不符合要求，应为三号字（约" + TITLE_FONT_SIZE + "磅）",
                    "将标题字号设置为三号字"
                ));
            }
            
            // 检查字体类型
            String fontFamily = run.getFontFamily();
            if (fontFamily != null && !fontFamily.contains(TITLE_FONT_FAMILY)) {
                results.add(new CheckResult(
                    "标题格式",
                    "第" + paragraphIndex + "段落",
                    "标题字体不符合要求，应为" + TITLE_FONT_FAMILY,
                    "将标题字体设置为" + TITLE_FONT_FAMILY
                ));
            }
            
            // 检查是否加粗
            if (!run.isBold()) {
                results.add(new CheckResult(
                    "标题格式",
                    "第" + paragraphIndex + "段落",
                    "标题未加粗",
                    "将标题设置为加粗"
                ));
            }
        }
    }

    /**
     * 检查正文段落格式
     */
    private void checkBodyParagraphFormat(XWPFParagraph paragraph, int paragraphIndex, List<CheckResult> results) {
        logger.debug("检查正文格式: 第{}段落", paragraphIndex);
        
        // 检查段落缩进
        CTP ctp = paragraph.getCTP();
        CTPPr pPr = ctp.getPPr();
        if (pPr != null && pPr.isSetInd()) {
            CTInd ind = pPr.getInd();
            // 首行缩进，单位是字符数的二十分之一英寸
            try {
                if (ind.getFirstLine() != null) {
                    double firstLineIndent = Double.parseDouble(ind.getFirstLine().toString());
                    // 约等于2个中文字符的缩进
                    int requiredIndent = 420; // 约等于2个字符的缩进
                    if (firstLineIndent < requiredIndent * 0.8) {
                        results.add(new CheckResult(
                            "段落格式",
                            "第" + paragraphIndex + "段落",
                            "段落首行缩进不足2字符",
                            "设置段落首行缩进为2字符"
                        ));
                    }
                } else {
                    results.add(new CheckResult(
                        "段落格式",
                        "第" + paragraphIndex + "段落",
                        "段落未设置首行缩进",
                        "设置段落首行缩进为2字符"
                    ));
                }
            } catch (NumberFormatException e) {
                logger.warn("解析段落缩进时出错", e);
                results.add(new CheckResult(
                    "段落格式",
                    "第" + paragraphIndex + "段落",
                    "无法解析段落缩进值",
                    "请手动检查段落缩进设置"
                ));
            }
        } else {
            results.add(new CheckResult(
                "段落格式",
                "第" + paragraphIndex + "段落",
                "段落未设置首行缩进",
                "设置段落首行缩进为2字符"
            ));
        }
        
        // 检查行距
        if (pPr != null && pPr.isSetSpacing()) {
            CTSpacing spacing = pPr.getSpacing();
            try {
                if (spacing.isSetLine() && spacing.getLine() != null) {
                    // Word中行距单位是twip，1.5倍行距约为360
                    double lineSpacing = Double.parseDouble(spacing.getLine().toString());
                    if (lineSpacing < 360) {
                        results.add(new CheckResult(
                            "段落格式",
                            "第" + paragraphIndex + "段落",
                            "段落行距小于1.5倍",
                            "设置段落行距为1.5倍"
                        ));
                    }
                } else {
                    results.add(new CheckResult(
                        "段落格式",
                        "第" + paragraphIndex + "段落",
                        "段落未设置行距",
                        "设置段落行距为1.5倍"
                    ));
                }
            } catch (NumberFormatException e) {
                logger.warn("解析行距时出错", e);
                results.add(new CheckResult(
                    "段落格式",
                    "第" + paragraphIndex + "段落",
                    "无法解析行距值",
                    "请手动检查行距设置"
                ));
            }
        } else {
            results.add(new CheckResult(
                "段落格式",
                "第" + paragraphIndex + "段落",
                "段落未设置行距",
                "设置段落行距为1.5倍"
            ));
        }
        
        // 检查字体
        for (XWPFRun run : paragraph.getRuns()) {
            // 检查字体大小
            if (run.getFontSize() != -1 && run.getFontSize() > BODY_FONT_SIZE + 2) {
                results.add(new CheckResult(
                    "正文格式",
                    "第" + paragraphIndex + "段落",
                    "正文字号过大，应为小四号字（约" + BODY_FONT_SIZE + "磅）",
                    "将正文字号设置为小四号字"
                ));
            }
            
            // 检查字体类型
            String fontFamily = run.getFontFamily();
            if (fontFamily != null && !fontFamily.contains(BODY_FONT_FAMILY)) {
                results.add(new CheckResult(
                    "正文格式",
                    "第" + paragraphIndex + "段落",
                    "正文字体不符合要求，应为" + BODY_FONT_FAMILY,
                    "将正文字体设置为" + BODY_FONT_FAMILY
                ));
            }
        }
    }

    /**
     * 检查表格格式
     */
    private void checkTables(XWPFDocument document, List<CheckResult> results) {
        logger.debug("检查表格格式");
        
        List<XWPFTable> tables = document.getTables();
        int tableIndex = 0;
        
        for (XWPFTable table : tables) {
            tableIndex++;
            
            // 检查表格标题
            XWPFTableRow firstRow = table.getRow(0);
            if (firstRow != null) {
                for (XWPFTableCell cell : firstRow.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        if (paragraph.getAlignment() != ParagraphAlignment.CENTER) {
                            results.add(new CheckResult(
                                "表格格式",
                                "表格" + tableIndex,
                                "表格标题未居中对齐",
                                "将表格标题设置为居中对齐"
                            ));
                            break;
                        }
                    }
                }
            }
            
            // 检查表格内容字体
            for (int i = 0; i < table.getNumberOfRows(); i++) {
                XWPFTableRow row = table.getRow(i);
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        for (XWPFRun run : paragraph.getRuns()) {
                            // 检查表格内容字体大小
                            if (run.getFontSize() != -1 && run.getFontSize() > BODY_FONT_SIZE + 2) {
                                results.add(new CheckResult(
                                    "表格格式",
                                    "表格" + tableIndex + "，第" + (i + 1) + "行",
                                    "表格内容字号过大",
                                    "将表格内容字号设置为小四号字或更小"
                                ));
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * 检查图片格式
     */
    private void checkPictures(XWPFDocument document, List<CheckResult> results) {
        logger.debug("检查图片格式");
        
        // 检查图片标题段落
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        int paragraphIndex = 0;
        
        for (XWPFParagraph paragraph : paragraphs) {
            paragraphIndex++;
            String text = paragraph.getText().trim();
            
            // 识别图片标题（通常以"图"开头）
            if (text.startsWith("图") && text.contains("：")) {
                if (paragraph.getAlignment() != ParagraphAlignment.CENTER) {
                    results.add(new CheckResult(
                        "图片格式",
                        "第" + paragraphIndex + "段落",
                        "图片标题未居中对齐",
                        "将图片标题设置为居中对齐"
                    ));
                }
            }
        }
    }

    /**
     * 导出检查报告为HTML文件
     *
     * @param results 检查结果
     * @param file    输出文件
     * @throws IOException 如果写入文件失败
     */
    public void exportReportToHtml(List<CheckResult> results, File file) throws IOException {
        logger.info("导出检查报告到: {}", file.getPath());
        
        try (FileWriter writer = new FileWriter(file)) {
            // HTML头部
            writer.write("<!DOCTYPE html>\n");
            writer.write("<html lang=\"zh-CN\">\n");
            writer.write("<head>\n");
            writer.write("    <meta charset=\"UTF-8\">\n");
            writer.write("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n");
            writer.write("    <title>论文格式检查报告</title>\n");
            writer.write("    <style>\n");
            writer.write("        body { font-family: Arial, sans-serif; margin: 40px; }\n");
            writer.write("        h1 { color: #333; text-align: center; }\n");
            writer.write("        .summary { margin: 20px 0; padding: 10px; background-color: #f5f5f5; border-radius: 5px; }\n");
            writer.write("        table { width: 100%; border-collapse: collapse; margin-top: 20px; }\n");
            writer.write("        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }\n");
            writer.write("        th { background-color: #4CAF50; color: white; }\n");
            writer.write("        tr:hover { background-color: #f5f5f5; }\n");
            writer.write("        .issue-type { font-weight: bold; }\n");
            writer.write("        .footer { margin-top: 30px; text-align: center; color: #777; font-size: 14px; }\n");
            writer.write("    </style>\n");
            writer.write("</head>\n");
            writer.write("<body>\n");
            
            // 报告标题
            writer.write("    <h1>论文格式检查报告</h1>\n");
            
            // 摘要信息
            writer.write("    <div class=\"summary\">\n");
            writer.write("        <p>检查时间: " + java.time.LocalDateTime.now() + "</p>\n");
            writer.write("        <p>发现问题数量: " + results.size() + "</p>\n");
            writer.write("    </div>\n");
            
            // 结果表格
            writer.write("    <table>\n");
            writer.write("        <tr>\n");
            writer.write("            <th>问题类型</th>\n");
            writer.write("            <th>位置</th>\n");
            writer.write("            <th>问题描述</th>\n");
            writer.write("            <th>修改建议</th>\n");
            writer.write("        </tr>\n");
            
            // 输出每个问题
            for (CheckResult result : results) {
                writer.write("        <tr>\n");
                writer.write("            <td class=\"issue-type\">" + result.getIssueType() + "</td>\n");
                writer.write("            <td>" + result.getLocation() + "</td>\n");
                writer.write("            <td>" + result.getDescription() + "</td>\n");
                writer.write("            <td>" + result.getSuggestion() + "</td>\n");
                writer.write("        </tr>\n");
            }
            
            writer.write("    </table>\n");
            
            // 如果没有问题
            if (results.isEmpty()) {
                writer.write("    <p style=\"text-align: center; color: green; font-weight: bold; margin-top: 30px;\">");
                writer.write("恭喜！未发现格式问题。</p>\n");
            }
            
            // 页脚
            writer.write("    <div class=\"footer\">\n");
            writer.write("        <p>论文格式检查工具 - 自动生成报告</p>\n");
            writer.write("    </div>\n");
            
            // HTML尾部
            writer.write("</body>\n");
            writer.write("</html>");
        }
        
        logger.info("报告导出完成");
    }
} 