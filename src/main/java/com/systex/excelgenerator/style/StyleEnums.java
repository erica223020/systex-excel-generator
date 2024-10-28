package com.systex.excelgenerator.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.BorderStyle;

public class StyleEnums {

    // FontStyle
    public enum FontStyle {
        BOLD {
            @Override
            public void applyFontStyle(Font font) {
                font.setBold(true);
            }
        },
        ITALIC {
            @Override
            public void applyFontStyle(Font font) {
                font.setItalic(true);
            }
        },
        UNDERLINE {
            @Override
            public void applyFontStyle(Font font) {
                font.setUnderline(Font.U_SINGLE);
            }
        },
        NORMAL {
            @Override
            public void applyFontStyle(Font font) {
                // 對 NORMAL 不做任何特殊設置，保持默認樣式
            }
        };

        public abstract void applyFontStyle(Font font);
    }

    // TextAlign
    public enum TextAlign {
        LEFT {
            @Override
            public void applyTextAlign(CellStyle style) {
                style.setAlignment(HorizontalAlignment.LEFT);
            }
        },
        CENTER {
            @Override
            public void applyTextAlign(CellStyle style) {
                style.setAlignment(HorizontalAlignment.CENTER);
            }
        },
        RIGHT {
            @Override
            public void applyTextAlign(CellStyle style) {
                style.setAlignment(HorizontalAlignment.RIGHT);
            }
        };

        public abstract void applyTextAlign(CellStyle style);
    }

    // BorderStyle
    public enum CustomBorderStyle {
        SOLID {
            @Override
            public void applyBorderStyle(CellStyle style) {
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
            }
        },
        DASHED {
            @Override
            public void applyBorderStyle(CellStyle style) {
                style.setBorderTop(BorderStyle.DASHED);
                style.setBorderBottom(BorderStyle.DASHED);
                style.setBorderLeft(BorderStyle.DASHED);
                style.setBorderRight(BorderStyle.DASHED);
            }
        },
        DOTTED {
            @Override
            public void applyBorderStyle(CellStyle style) {
                style.setBorderTop(BorderStyle.DOTTED);
                style.setBorderBottom(BorderStyle.DOTTED);
                style.setBorderLeft(BorderStyle.DOTTED);
                style.setBorderRight(BorderStyle.DOTTED);
            }
        };

        public abstract void applyBorderStyle(CellStyle style);
    }
}
