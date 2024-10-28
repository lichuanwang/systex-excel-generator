package com.systex.excelgenerator.style;

import org.apache.poi.ss.usermodel.*;

public class StyleBuilder {

    private CellStyle cellStyle;
    private Font font;

    public StyleBuilder(Workbook workbook) {
        this.cellStyle = workbook.createCellStyle();
        this.font = workbook.createFont();
    }

    public StyleBuilder setBold(boolean bold) {
        font.setBold(bold);
        return this;
    }

    public StyleBuilder setFontSize(short size) {
        font.setFontHeightInPoints(size);
        return this;
    }

    public StyleBuilder setBorder(BorderStyle borderStyle) {
        cellStyle.setBorderBottom(borderStyle);
        cellStyle.setBorderTop(borderStyle);
        cellStyle.setBorderLeft(borderStyle);
        cellStyle.setBorderRight(borderStyle);
        return this;
    }

    public StyleBuilder setAlignment(HorizontalAlignment alignment) {
        cellStyle.setAlignment(alignment);
        return this;
    }

    public CellStyle build() {
        cellStyle.setFont(font);
        return cellStyle;
    }
}
