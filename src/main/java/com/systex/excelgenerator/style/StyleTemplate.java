//package com.systex.excelgenerator.style;
//
//import org.apache.poi.ss.usermodel.IndexedColors;
//
//public class StyleTemplate {
//    private short fontSize;
//    private boolean isBold;
//    private short fontColor;
//    private short backgroundColor;
//
//    // 預設模板
//    public StyleTemplate() {
//        this.fontSize = 12;
//        this.isBold = false;
//        this.fontColor = IndexedColors.BLACK.getIndex();
//        this.backgroundColor = IndexedColors.WHITE.getIndex();
//    }
//    // 複製CellStyle 複製樣式(深拷貝，不可互相影響)。
//    // 設定好後，可以重複使用到別的Cell
//
//    // 寫一個空的建構子
//
//
//
//    // 自訂模板
//    //DTO
//    public StyleTemplate(short fontSize, boolean isBold, short fontColor, short backgroundColor) {
//        this.fontSize = fontSize;
//        this.isBold = isBold;
//        this.fontColor = fontColor;
//        this.backgroundColor = backgroundColor;
//    }
//
//    public short getFontSize() {
//        return fontSize;
//    }
//
//    public void setFontSize(short fontSize) {
//        this.fontSize = fontSize;
//    }
//
//    public boolean isBold() {
//        return isBold;
//    }
//
//    public void setBold(boolean bold) {
//        isBold = bold;
//    }
//
//    public short getFontColor() {
//        return fontColor;
//    }
//
//    public void setFontColor(short fontColor) {
//        this.fontColor = fontColor;
//    }
//
//    public short getBackgroundColor() {
//        return backgroundColor;
//    }
//
//    public void setBackgroundColor(short backgroundColor) {
//        this.backgroundColor = backgroundColor;
//    }
//}
