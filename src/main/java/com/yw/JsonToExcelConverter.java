package com.yw;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class JsonToExcelConverter {

    private static final ObjectMapper objectMapper = new ObjectMapper();
    private static final int MAX_CELL_LENGTH = 32700; // 留一些余量
    private static final int COMMENT_PREVIEW_LENGTH = 1000; // 注释中预览的字符数

    /**
     * 读取 JSON 文件并转换为 Excel（带注释版本）
     *
     * @param jsonFilePath  JSON 文件路径
     * @param excelFilePath 输出的 Excel 文件路径
     */
    public static void convertJsonToExcel(String jsonFilePath, String excelFilePath) {
        Workbook workbook = null;

        try {
            // 1. 读取 JSON 文件
            ApiResponse apiResponse = readJsonFile(jsonFilePath);

            // 2. 验证数据
            if (apiResponse == null || apiResponse.getData() == null ||
                    apiResponse.getData().getRecords() == null) {
                System.out.println("JSON 数据格式错误或数据为空");
                return;
            }

            List<Map<String, Object>> records = apiResponse.getData().getRecords();
            if (records.isEmpty()) {
                System.out.println("数据列表为空");
                return;
            }

            // 3. 创建 Excel 文件
            workbook = createExcelFileWithComments(records, excelFilePath);

            System.out.println("Excel 文件生成成功: " + excelFilePath);
            System.out.println("共处理 " + records.size() + " 条数据");

        } catch (Exception e) {
            System.err.println("转换过程中发生错误: " + e.getMessage());
            e.printStackTrace();
        } finally {
            // 确保工作簿被关闭
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    System.err.println("关闭工作簿时发生错误: " + e.getMessage());
                }
            }
        }
    }

    /**
     * 读取 JSON 文件
     */
    private static ApiResponse readJsonFile(String jsonFilePath) throws IOException {
        FileInputStream inputStream = new FileInputStream(jsonFilePath);
        return objectMapper.readValue(inputStream, ApiResponse.class);
    }

    /**
     * 创建带注释的 Excel 文件
     */
    private static Workbook createExcelFileWithComments(List<Map<String, Object>> records, String excelFilePath)
            throws IOException {

        // 创建 Workbook
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("数据导出");

        // 创建样式
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);

        // 获取所有字段名（表头）
        Set<String> allFields = getAllFields(records);
        List<String> fieldList = new ArrayList<>(allFields);

        // 创建表头
        createHeaderRow(sheet, headerStyle, fieldList);

        // 填充数据（带注释）
        fillDataRowsWithComments(workbook, sheet, dataStyle, records, fieldList);

//        // 自动调整列宽
//        autoSizeColumns(sheet, fieldList.size());

        // 写入文件
        try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
            workbook.write(outputStream);
        }

        return workbook;
    }

    /**
     * 获取所有字段名
     */
    private static Set<String> getAllFields(List<Map<String, Object>> records) {
        Set<String> allFields = new HashSet<>();
        for (Map<String, Object> record : records) {
            allFields.addAll(record.keySet());
        }
        return allFields;
    }

    /**
     * 创建表头行
     */
    private static void createHeaderRow(Sheet sheet, CellStyle headerStyle, List<String> fields) {
        Row headerRow = sheet.createRow(0);
        int colIndex = 0;

        for (String field : fields) {
            Cell cell = headerRow.createCell(colIndex++);
            cell.setCellValue(field);
            cell.setCellStyle(headerStyle);
        }
    }

    /**
     * 填充数据行（带注释版本）
     */
    private static void fillDataRowsWithComments(Workbook workbook, Sheet sheet, CellStyle dataStyle,
                                                 List<Map<String, Object>> records, List<String> fields) {

        int rowIndex = 1;

        for (Map<String, Object> record : records) {
            Row row = sheet.createRow(rowIndex);

            for (int colIndex = 0; colIndex < fields.size(); colIndex++) {
                String field = fields.get(colIndex);
                Object value = record.get(field);
                Cell cell = row.createCell(colIndex);
                cell.setCellStyle(dataStyle);

                // 使用带注释的单元格值设置方法
                setCellValueWithComment(workbook, cell, value, field);
            }
            rowIndex++;
        }
    }

    /**
     * 设置单元格值并添加注释（核心方法）
     */
    private static void setCellValueWithComment(Workbook workbook, Cell cell, Object value, String fieldName) {
        if (value == null) {
            cell.setCellValue("");
            return;
        }

        String stringValue;
        if (value instanceof String) {
            stringValue = (String) value;
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
            return;
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
            return;
        } else {
            stringValue = value.toString();
        }

        // 检查文本长度
        if (stringValue.length() > MAX_CELL_LENGTH) {
            // 处理超长文本：单元格显示摘要，完整内容存入注释
            handleLongTextWithComment(workbook, cell, stringValue, fieldName);
        } else {
            // 正常文本直接显示
            cell.setCellValue(stringValue);
        }
    }

    /**
     * 处理超长文本并添加注释
     */
    private static void handleLongTextWithComment(Workbook workbook, Cell cell, String fullText, String fieldName) {
        // 1. 在单元格中显示摘要信息
        String displayText = createDisplayText(fullText, fieldName);
        cell.setCellValue(displayText);

        // 2. 添加注释显示完整内容（前一部分）
        addCommentToCell(workbook, cell, fullText, fieldName);

        // 3. 可选：添加单元格样式提示
        addVisualIndicator(cell);
    }

    /**
     * 创建单元格显示文本
     */
    private static String createDisplayText(String fullText, String fieldName) {
        int totalLength = fullText.length();

        // 根据内容类型创建不同的显示文本
        if (isJsonLike(fullText)) {
            String preview = fullText.substring(0, Math.min(200, fullText.length()));
            return "📊 [JSON数据: " + totalLength + "字符] " + preview + "...";
        } else if (isXmlLike(fullText)) {
            String preview = fullText.substring(0, Math.min(200, fullText.length()));
            return "📋 [XML数据: " + totalLength + "字符] " + preview + "...";
        } else if (isBase64Like(fullText)) {
            return "🔒 [Base64数据: " + totalLength + "字符]";
        } else {
            // 普通文本，显示开头部分
            String preview = fullText.substring(0, Math.min(100, fullText.length()));
            return "📝 [" + fieldName + ": " + totalLength + "字符] " + preview + "...";
        }
    }

    /**
     * 为单元格添加注释
     */
    private static void addCommentToCell(Workbook workbook, Cell cell, String fullText, String fieldName) {
        try {
            // 获取或创建绘图 patriarch
            Drawing<?> drawing = cell.getSheet().createDrawingPatriarch();
            if (drawing == null) {
                drawing = cell.getSheet().createDrawingPatriarch();
            }

            // 创建注释锚点
            ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
            anchor.setCol1(cell.getColumnIndex());
            anchor.setCol2(cell.getColumnIndex() + 3);
            anchor.setRow1(cell.getRowIndex());
            anchor.setRow2(cell.getRowIndex() + 5);

            // 创建注释
            Comment comment = drawing.createCellComment(anchor);

            // 设置注释作者
            comment.setAuthor("数据导出系统");

            // 创建注释内容
            String commentContent = createCommentContent(fullText, fieldName);
            RichTextString commentString = workbook.getCreationHelper().createRichTextString(commentContent);

            // 设置注释样式（如果支持）
            try {
                // 尝试设置注释字体（可能在某些版本中不支持）
                Font commentFont = workbook.createFont();
                commentFont.setFontName("宋体");
                commentFont.setFontHeightInPoints((short) 9);
                commentString.applyFont(commentFont);
            } catch (Exception e) {
                // 忽略字体设置错误
                System.out.println("注释字体设置失败，使用默认字体");
            }

            comment.setString(commentString);
            cell.setCellComment(comment);

        } catch (Exception e) {
            System.err.println("添加注释失败: " + e.getMessage());
            // 即使注释失败，也要确保单元格有值
            cell.setCellValue("[内容过长: " + fullText.length() + "字符]");
        }
    }

    /**
     * 创建注释内容
     */
    private static String createCommentContent(String fullText, String fieldName) {
        StringBuilder comment = new StringBuilder();
        comment.append("字段: ").append(fieldName).append("\n");
        comment.append("总长度: ").append(fullText.length()).append(" 字符\n");
        comment.append("预览内容:\n");
        comment.append("----------------------------------------\n");

        // 添加预览内容
        String preview = fullText.substring(0, Math.min(COMMENT_PREVIEW_LENGTH, fullText.length()));
        comment.append(preview);

        if (fullText.length() > COMMENT_PREVIEW_LENGTH) {
            comment.append("\n----------------------------------------\n");
            comment.append("... [剩余 ").append(fullText.length() - COMMENT_PREVIEW_LENGTH).append(" 字符未显示]");
        }

        // 添加内容类型提示
        if (isJsonLike(fullText)) {
            comment.append("\n\n📌 内容类型: JSON 数据");
        } else if (isXmlLike(fullText)) {
            comment.append("\n\n📌 内容类型: XML 数据");
        } else if (isBase64Like(fullText)) {
            comment.append("\n\n📌 内容类型: Base64 编码数据");
        } else {
            comment.append("\n\n📌 内容类型: 文本数据");
        }

        return comment.toString();
    }

    /**
     * 添加视觉指示器
     */
    private static void addVisualIndicator(Cell cell) {
        // 可以设置单元格背景色来提示有注释
        CellStyle style = cell.getCellStyle();
        CellStyle newStyle = cell.getSheet().getWorkbook().createCellStyle();
        newStyle.cloneStyleFrom(style);

        // 设置浅黄色背景提示
        newStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cell.setCellStyle(newStyle);
    }

    /**
     * 内容类型判断方法
     */
    private static boolean isJsonLike(String text) {
        if (text == null || text.trim().isEmpty()) return false;
        String trimmed = text.trim();
        return (trimmed.startsWith("{") && trimmed.endsWith("}")) ||
                (trimmed.startsWith("[") && trimmed.endsWith("]"));
    }

    private static boolean isXmlLike(String text) {
        if (text == null || text.trim().isEmpty()) return false;
        String trimmed = text.trim();
        return trimmed.startsWith("<?xml") ||
                (trimmed.startsWith("<") && trimmed.endsWith(">"));
    }

    private static boolean isBase64Like(String text) {
        if (text == null || text.length() < 20) return false;
        // 简单的Base64特征检查
        return text.matches("^[A-Za-z0-9+/]*={0,2}$") && text.length() % 4 == 0;
    }

    /**
     * 自动调整列宽
     */
    private static void autoSizeColumns(Sheet sheet, int columnCount) {
        for (int i = 0; i < columnCount; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    /**
     * 创建表头样式
     */
    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();

        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        font.setColor(IndexedColors.WHITE.getIndex());

        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    /**
     * 创建数据样式
     */
    private static CellStyle createDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setWrapText(true); // 允许文本换行

        return style;
    }
}

