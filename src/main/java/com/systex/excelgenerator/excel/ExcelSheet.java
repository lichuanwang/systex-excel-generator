package com.systex.excelgenerator.excel;

import com.systex.excelgenerator.component.Section;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class ExcelSheet {
    private final XSSFSheet xssfSheet;
    private int startingRow = 0;
    private int startingCol = 0;
    private int maxColPerRow;
    private int deepestRowOnCurrentLevel = 0;

    public ExcelSheet(Workbook workbook, String sheetName, int maxColPerRow) {
        this.xssfSheet = (XSSFSheet) workbook.createSheet(sheetName); // directly new a sheet
        this.maxColPerRow = maxColPerRow;
    }
    public Workbook getWorkbook() {
        return xssfSheet.getWorkbook();
    }

    public <T> void addSection(Section<T> section) {
        // Check if adding the section would exceed maxColPerRow
        if (startingCol + section.getWidth() > maxColPerRow) {
            // Move to next row if max columns exceeded
            startingRow = deepestRowOnCurrentLevel + 2;
            startingCol = 0;
        }

        // Set section start point and render the section
        section.render(this, startingRow, startingCol);

        // Update layout positions
        startingCol += section.getWidth();
        deepestRowOnCurrentLevel = Math.max(deepestRowOnCurrentLevel, startingRow + section.getHeight() + 1);
    }

    // Method to create or get a row
    public Row createOrGetRow(int rowNum) {
        Row row = xssfSheet.getRow(rowNum);
        if (row == null) {
            row = xssfSheet.createRow(rowNum);
        }
        return row;
    }

    // Getter for the underlying XSSFSheet, if needed
    public XSSFSheet getXssfSheet() {
        return xssfSheet;
    }


    public int getStartingRow() {
        return startingRow;
    }

    public void setStartingRow(int startingRow) {
        this.startingRow = startingRow;
    }

    public int getStartingCol() {
        return startingCol;
    }

    public void setStartingCol(int startingCol) {
        this.startingCol = startingCol;
    }

    public int getMaxColPerRow() {
        return maxColPerRow;
    }

    public int getDeepestRowOnCurrentLevel() {
        return deepestRowOnCurrentLevel;
    }

    public void setDeepestRowOnCurrentLevel(int deepestRowOnCurrentLevel) {
        this.deepestRowOnCurrentLevel = deepestRowOnCurrentLevel;
    }

    public void setMaxColPerRow(int maxColPerRow) {
        this.maxColPerRow = maxColPerRow;
    }

    public int getMaxColPerRow(int maxColPerRow) {
        return maxColPerRow;
    }
}