package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Education;
import com.systex.excelgenerator.model.Experience;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

public abstract class AbstractSection<T> implements Section<T> {
    protected String title;
    protected XSSFWorkbook workbook;
    protected CellStyle headerStyle;
    protected CellStyle dataStyle;

    public AbstractSection(String title, XSSFWorkbook workbook) {
        this.title = title;
        this.workbook = workbook;
        initializeStyles();
    }

    private void initializeStyles() {
        // 設定標題的樣式
        headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 18);
        headerStyle.setFont(headerFont);

        // 設定內文的樣式
        dataStyle = workbook.createCellStyle();
        Font dataFont = workbook.createFont();
        dataFont.setFontHeightInPoints((short) 12);
        dataStyle.setFont(dataFont);
    }

    protected CellStyle getHeaderStyle() {
        return headerStyle;
    }

    protected CellStyle getDataStyle() {
        return dataStyle;
    }

    // wrong naming
    protected abstract int generateHeader(XSSFSheet sheet, int rowNum);
    protected abstract int generateData(XSSFSheet sheet, int rowNum);
    protected abstract int generateFooter(XSSFSheet sheet, int rowNum);


    // problem with pass by value, should we use a rowNum or primitive type to determine in this way
    // probably using pass by reference could be better
    public int populate(XSSFSheet sheet, int rowNum) {
        addSectionTitle(sheet, rowNum);
        rowNum++; // put this in the addSectionTitle
        rowNum = generateHeader(sheet, rowNum);
        rowNum = generateData(sheet, rowNum);
        rowNum = generateFooter(sheet, rowNum);
        return rowNum;
    }


    public void addSectionTitle(XSSFSheet sheet, int rowNum) {
        Row headerRow = sheet.createRow(rowNum);
        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue(this.title);
        headerRow.getCell(0).setCellStyle(getHeaderStyle());

        headerCell.setCellStyle(getHeaderStyle());
//        // Apply style if needed (e.g., bold, font size)
//        CellStyle style = sheet.getWorkbook().createCellStyle();
//        Font font = sheet.getWorkbook().createFont();
//        font.setBold(true);
//        font.setFontHeightInPoints((short) 14);
//        style.setFont(font);
//        headerCell.setCellStyle(style);
    }

}
