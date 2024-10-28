//package com.systex.excelgenerator.component;
//
//import com.systex.excelgenerator.model.Candidate;
//import com.systex.excelgenerator.component.Section;
//import org.apache.poi.ss.usermodel.CellStyle;
//import org.apache.poi.ss.usermodel.IndexedColors;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.ss.util.CellRangeAddress;
//import com.systex.excelgenerator.style.StyleEnums;
//import com.systex.excelgenerator.style.StyleBuilder;
//
//import java.text.SimpleDateFormat;
//
//
//public class PersonalInfoSection extends Section {
//
//    private Candidate candidate;
//    private Workbook workbook;
//
//    private CellStyle headerStyle;
//    private CellStyle labelStyle;
//    private CellStyle infoStyle;
//
//    public PersonalInfoSection(Candidate candidate, Workbook workbook) {
//        super("Personal Information");
//        this.candidate = candidate;
//        this.workbook = workbook;
//        StyleBuilder styleBuilder = new StyleBuilder(workbook);
//
//        this.headerStyle = styleBuilder.setFontStyle(StyleEnums.FontStyle.BOLD)
//                .setFontSize((short) 14)
//                .setTextAlign(StyleEnums.TextAlign.CENTER)
//                .build();
//
//        this.labelStyle = styleBuilder.setFontStyle(StyleEnums.FontStyle.BOLD)
//                .setFontSize((short) 12)
//                .setTextAlign(StyleEnums.TextAlign.CENTER)
//                .setBackgroundColor(IndexedColors.LIGHT_YELLOW.getIndex())
//                .build();
//
//        this.infoStyle = styleBuilder.setFontStyle(StyleEnums.FontStyle.NORMAL)
//                .setFontSize((short) 12)
//                .setTextAlign(StyleEnums.TextAlign.CENTER)
//                .build();
//    }
//
//    @Override
//    public int populate(XSSFSheet sheet, int rowNum) {
//        // 合併 "Personal Information" 標題跨越兩列
//        Row headerRow = sheet.createRow(rowNum++);
//        createStyledCell(headerRow, 0, "Personal Information", headerStyle);
//
//        Row row = sheet.createRow(rowNum++);
//        createStyledCell(row, 0, "Name", labelStyle);
//        createStyledCell(row, 1, candidate.getName(), infoStyle);
//
//        row = sheet.createRow(rowNum++);
//        createStyledCell(row, 0, "Gender", labelStyle);
//        createStyledCell(row, 1, candidate.getGender(), infoStyle);
//
//        row = sheet.createRow(rowNum++);
//        createStyledCell(row, 0, "Birthday", labelStyle);
//        createStyledCell(row, 1, SimpleDateFormat.getDateInstance().format(candidate.getBirthday()), infoStyle);
//
//        row = sheet.createRow(rowNum++);
//        createStyledCell(row, 0, "Phone", labelStyle);
//        createStyledCell(row, 1, candidate.getPhone(), infoStyle);
//
//        row = sheet.createRow(rowNum++);
//        createStyledCell(row, 0, "Email", labelStyle);
//        createStyledCell(row, 1, candidate.getEmail(), infoStyle);
//
//        row = sheet.createRow(rowNum++);
//        createStyledCell(row, 0, "Address", labelStyle);
//        createStyledCell(row, 1, candidate.getAddress().toString(), infoStyle);
//
//        return rowNum;
//    }
//
//    // 幫助方法：創建應用樣式的單元格
//    private void createStyledCell(Row row, int column, String value, CellStyle style) {
//        row.createCell(column).setCellValue(value);
//        row.getCell(column).setCellStyle(style);
//    }
//
//    // 大標題樣式 (14號字，粗體，置中)
//    private CellStyle createHeaderStyle(Workbook workbook) {
//        StyleBuilder styleBuilder = new StyleBuilder(workbook);
//        return styleBuilder.setFontStyle(StyleEnums.FontStyle.BOLD)
//                .setFontSize((short) 14)
//                .setTextAlign(StyleEnums.TextAlign.CENTER)
//                .build();
//    }
//
//    // 小標題樣式 (12號字，粗體，背景黃色，置中)
//    private CellStyle createLabelStyle(Workbook workbook) {
//        StyleBuilder styleBuilder = new StyleBuilder(workbook);
//        return styleBuilder.setFontStyle(StyleEnums.FontStyle.BOLD)
//                .setFontSize((short) 12)
//                .setTextAlign(StyleEnums.TextAlign.CENTER)
//                .setBackgroundColor(IndexedColors.LIGHT_YELLOW.getIndex())
//                .build();
//    }
//
//    // 資訊樣式 (12號字，正常字體，置中)
//    private CellStyle createInfoStyle(Workbook workbook) {
//        StyleBuilder styleBuilder = new StyleBuilder(workbook);
//        return styleBuilder.setFontStyle(StyleEnums.FontStyle.NORMAL)
//                .setFontSize((short) 12)
//                .setTextAlign(StyleEnums.TextAlign.CENTER)
//                .build();
//    }
//}