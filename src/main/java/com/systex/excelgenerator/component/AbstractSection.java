package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public abstract class AbstractSection<T> implements Section<T> {
    protected String title;

    public AbstractSection(String title) {
        this.title = title;
    }

    // wrong naming
    protected abstract void populateHeader(ExcelSheet sheet, int startRow, int startCol);
    protected abstract void populateBody(ExcelSheet sheet, int startRow, int startCol);
    protected abstract void populateFooter(ExcelSheet sheet, int startRow, int startCol);


    // problem with pass by value, should we use a rowNum or primitive type to determine in this way
    // probably using pass by reference could be better
    public void render(ExcelSheet sheet, int startRow, int startCol) {
        addSectionTitle(sheet, startRow, startCol);
        populateHeader(sheet, startRow + 1, startCol);
        populateBody(sheet, startRow + 2, startCol);
        populateFooter(sheet, startRow + 3, startCol);
//        XSSFSheet sheet = excelSheet.getUnderlyingSheet();
//        if (sheet.exceedMaxColPerRow()) {
//            sheet.jumpToNextAvailableRow();
//        }
//        addSectionTitle(sheet);
//        sheet.setStartingCol(sheet.getStartingCol() + 2);
    }


//    public void addSectionTitle(ExcelSheet sheet) {
//        int rowNum = sheet.getStartingRow();
//        Row headerRow = sheet.createOrGetRow(rowNum);
//        Cell headerCell = headerRow.createCell(sheet.getStartingCol());
//        headerCell.setCellValue(this.title);
//
//        // Apply style if needed (e.g., bold, font size)
//        CellStyle style = sheet.getUnderlyingSheet().getWorkbook().createCellStyle();
//        Font font = sheet.getUnderlyingSheet().getWorkbook().createFont();
//        font.setBold(true);
//        font.setFontHeightInPoints((short) 14);
//        style.setFont(font);
//        headerCell.setCellStyle(style);
//    }

    public void addSectionTitle(ExcelSheet excelSheet, int startRow, int startCol) {
        Row headerRow = excelSheet.createOrGetRow(startRow);
        Cell headerCell = headerRow.createCell(startCol);
        headerCell.setCellValue(this.title);

        // Apply style if needed (e.g., bold, font size)
        XSSFSheet sheet = excelSheet.getUnderlyingSheet();
        CellStyle style = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        style.setFont(font);
        headerCell.setCellStyle(style);
    }



}
