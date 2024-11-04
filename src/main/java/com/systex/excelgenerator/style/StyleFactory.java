//package com.systex.excelgenerator.style;
//
//import org.apache.poi.ss.usermodel.CellStyle;
//import org.apache.poi.ss.usermodel.FillPatternType;
//import org.apache.poi.ss.usermodel.Font;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import java.util.Map;
//
//public class StyleFactory {
//
//    private Workbook workbook;
//
//    public StyleFactory(Workbook workbook) {
//        this.workbook = workbook;
//    }
//
//    // 生成一組模板
//    public CellStyle createFromTemplate(StyleTemplate template, Map<String, Object> overrides) {
//        CellStyle style = workbook.createCellStyle();
//        Font font = workbook.createFont();
//
//        // 使用模板的預設值
//        font.setFontHeightInPoints(template.getFontSize());
//        font.setBold(template.isBold());
//        font.setColor(template.getFontColor());
//        style.setFillForegroundColor(template.getBackgroundColor());
//        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//
//        // 覆蓋模板一部分+轉型別
//        if (overrides != null) {
//            // 覆蓋字體大小、轉型別
//            if (overrides.containsKey("fontSize")) {
//                Object value = overrides.get("fontSize");
//                if (value instanceof Number) {
//                    font.setFontHeightInPoints(((Number) value).shortValue());
//                } else {
//                    System.out.println("請輸入數字喔^_^");
//                }
//            }
//
//            // 覆蓋粗體
//            if (overrides.containsKey("isBold")) {
//                Object value = overrides.get("isBold");
//                if (value instanceof Boolean) {
//                    font.setBold((Boolean) value);
//                } else {
//                    System.out.println("請輸入boolean值喔^_^");
//                }
//            }
//
//            // 覆蓋字的顏色
//            //IndexedColors.RED.getIndex()
//            if (overrides.containsKey("fontColor")) {
//                Object value = overrides.get("fontColor");
//                if (value instanceof Short) {
//                    font.setColor((Short) value);
//                } else {
//                    System.out.println("請輸入short值喔^_^");
//                }
//            }
//
//            // 覆蓋背景顏色
//            if (overrides.containsKey("backgroundColor")) {
//                Object value = overrides.get("backgroundColor");
//                if (value instanceof Short) {
//                    style.setFillForegroundColor((Short) value);
//                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//                } else {
//                    System.out.println("請輸入short值喔^_^");
//                }
//            }
//        }
//        return style;
//    }
//}
