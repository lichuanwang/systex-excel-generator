package com.systex.excelgenerator.component;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public abstract class Section {
    protected String title;
    static int relativeRow = 0;
    static int relativeColumn = 0;
    static int maxCol = 12;
    static int nextRelativeRow = 0;
    public Section(String title) {
        this.title = title;
    }

    public abstract int populate(XSSFSheet sheet);

    public void addHeader(XSSFSheet sheet) {
        Row headerRow = (sheet.getRow(relativeRow) == null)?
                sheet.createRow(relativeRow):sheet.getRow(relativeRow);
        Cell headerCell = headerRow.createCell(relativeColumn);
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
