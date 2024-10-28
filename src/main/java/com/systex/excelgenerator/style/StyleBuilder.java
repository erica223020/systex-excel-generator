package com.systex.excelgenerator.style;

import org.apache.poi.ss.usermodel.*;
import static com.systex.excelgenerator.style.StyleEnums.FontStyle;
import static com.systex.excelgenerator.style.StyleEnums.TextAlign;
import static com.systex.excelgenerator.style.StyleEnums.CustomBorderStyle;
import org.apache.poi.ss.util.CellRangeAddress;

public class StyleBuilder {

    private Workbook workbook;
    private CellStyle cellStyle;
    private Font font;

    public StyleBuilder(Workbook workbook) {
        this.workbook = workbook;
        this.cellStyle = workbook.createCellStyle();
        this.font = workbook.createFont();
    }

    // setFontSize
    public StyleBuilder setFontSize(short size) {
        font.setFontHeightInPoints(size);
        return this;
    }

    // setFontStyle
    public StyleBuilder setFontStyle(FontStyle... styles) {
        for(FontStyle style :styles) {
            style.applyFontStyle(font);
        }
        return this;
    }

    // setColor
    public StyleBuilder setFontColor(short color) {
        font.setColor(color);
        return this;
    }

    // setBG
    public StyleBuilder setBackgroundColor(short color) {
        cellStyle.setFillForegroundColor(color);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return this;
    }

    // setTextAlign
    public StyleBuilder setTextAlign(TextAlign... aligns) {
        for(TextAlign align :aligns) {
            align.applyTextAlign(cellStyle);
        }
        return this;
    }

    // setBorderStyle
    public StyleBuilder setBorderStyle(CustomBorderStyle... styles) {
        for(CustomBorderStyle style :styles) {
            style.applyBorderStyle(cellStyle);
        }
        return this;
    }

    // mergeCell
    public StyleBuilder mergeCells(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        return this;
    }

    public CellStyle build() {
        cellStyle.setFont(font);
        return cellStyle;
    }
}
