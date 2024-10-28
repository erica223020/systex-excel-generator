package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Candidate;
import com.systex.excelgenerator.component.Section;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.util.CellRangeAddress;
import com.systex.excelgenerator.style.StyleEnums;
import com.systex.excelgenerator.style.StyleBuilder;

import java.text.SimpleDateFormat;

public class PersonalInfoSection extends Section {

    private Candidate candidate;

    public PersonalInfoSection(Candidate candidate) {
        super("Personal Information");
        this.candidate = candidate;
    }

    @Override
    public int populate(XSSFSheet sheet, int rowNum) {
        // 創建 StyleBuilder
        StyleBuilder styleBuilder = new StyleBuilder(sheet.getWorkbook());

        CellStyle headerStyle = createHeaderStyle(new StyleBuilder(sheet.getWorkbook()));
        CellStyle dataStyle = createDataStyle(new StyleBuilder(sheet.getWorkbook()));
        CellStyle emailStyle = createEmailStyle(new StyleBuilder(sheet.getWorkbook()));
        CellStyle leftColumnStyle = createLeftColumnStyle(new StyleBuilder(sheet.getWorkbook()));
        // 設置標題行樣式
//        CellStyle headerStyle = createHeaderStyle(styleBuilder);
        // 設置數據行樣式
//        CellStyle dataStyle = createDataStyle(styleBuilder);
//        // 設置 Email 的特殊樣式
//        CellStyle emailStyle = createEmailStyle(styleBuilder);
//        // 設置左邊的標題欄的樣式，使用黃色背景
//        CellStyle leftColumnStyle = createLeftColumnStyle(styleBuilder);

        // 合併 "Personal Information" 標題跨越兩列
        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 0, 1));
        Row headerRow = sheet.createRow(rowNum++);
        createStyledCell(headerRow, 0, "Personal Information", headerStyle);

        // 填充個人信息數據
        Row row = sheet.createRow(rowNum++);

        createStyledCell(row, 0, "Name", leftColumnStyle);
        createStyledCell(row, 1, candidate.getName(), dataStyle);

        row = sheet.createRow(rowNum++);
        createStyledCell(row, 0, "Gender", leftColumnStyle);
        createStyledCell(row, 1, candidate.getGender(), dataStyle);

        row = sheet.createRow(rowNum++);
        createStyledCell(row, 0, "Birthday", leftColumnStyle);
        createStyledCell(row, 1, SimpleDateFormat.getDateInstance().format(candidate.getBirthday()), dataStyle);

        row = sheet.createRow(rowNum++);
        createStyledCell(row, 0, "Phone", leftColumnStyle);
        createStyledCell(row, 1, candidate.getPhone(), dataStyle);

        row = sheet.createRow(rowNum++);
        createStyledCell(row, 0, "Email", leftColumnStyle);
        createStyledCell(row, 1, candidate.getEmail(), emailStyle); // 使用特殊樣式

        row = sheet.createRow(rowNum++);
        createStyledCell(row, 0, "Address", leftColumnStyle);
        createStyledCell(row, 1, candidate.getAddress().toString(), dataStyle);

        return rowNum;
    }

    // 幫助方法：創建應用樣式的單元格
    private void createStyledCell(Row row, int column, String value, CellStyle style) {
        row.createCell(column).setCellValue(value);
        row.getCell(column).setCellStyle(style);
    }

    // 提取樣式創建部分
    private CellStyle createHeaderStyle(StyleBuilder styleBuilder) {
        return styleBuilder.setFontStyle(StyleEnums.FontStyle.BOLD)
                .setTextAlign(StyleEnums.TextAlign.CENTER) // 水平居中
                .setFontSize((short) 16)
                .setBorderStyle(StyleEnums.CustomBorderStyle.SOLID)
                .build();
    }

    private CellStyle createDataStyle(StyleBuilder styleBuilder) {
        return styleBuilder.setFontStyle(StyleEnums.FontStyle.NORMAL)
                .setTextAlign(StyleEnums.TextAlign.CENTER) // 水平居中
                .setFontSize((short) 10)
                .setBorderStyle(StyleEnums.CustomBorderStyle.SOLID)
                .build();
    }

    private CellStyle createEmailStyle(StyleBuilder styleBuilder) {
        return styleBuilder.setFontStyle(StyleEnums.FontStyle.ITALIC)
                .setTextAlign(StyleEnums.TextAlign.CENTER) // 水平居中
                .setFontSize((short) 20)
                .setFontColor(IndexedColors.BLUE.getIndex())
                .build();
    }

    private CellStyle createLeftColumnStyle(StyleBuilder styleBuilder) {
        return styleBuilder.setFontStyle(StyleEnums.FontStyle.BOLD)
                .setTextAlign(StyleEnums.TextAlign.CENTER) // 水平居中
                .setFontSize((short) 12)
                .setBackgroundColor(IndexedColors.LIGHT_YELLOW.getIndex())
                .setBorderStyle(StyleEnums.CustomBorderStyle.SOLID)
                .build();
    }
}
