package com.systex.excelgenerator.component;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public abstract class Section {
    protected String title;

    public Section(String title) {
        this.title = title;
    }

    public abstract int populate(XSSFSheet sheet, int rowNum);

    public void addHeader(XSSFSheet sheet, int rowNum) {
        Row headerRow = sheet.createRow(rowNum);
        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue(this.title);

        // Apply style if needed (e.g., bold, font size)
        CellStyle style = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        style.setFont(font);
        headerCell.setCellStyle(style);
    }
}
