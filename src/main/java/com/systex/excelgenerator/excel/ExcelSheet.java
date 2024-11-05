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

    public void addSection(Section section) {
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




//package com.systex.excelgenerator.excel;
//
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//public class ExcelSheet {
//    private final XSSFSheet sheet;
////    private int currentTitleRow = 0;
//    private int startingRow = 0;
//    private int startingCol = 0;
//    private int maxColPerRow = 15;
//    private int deepestRowOnCurrentLevel;
//
//    // Constructor to initialize the sheet
//    public ExcelSheet(XSSFSheet sheet) {
//        this.sheet = sheet;
//    }
//
//    // Getters and setters for custom properties
//    public int getStartingRow() {
//        return startingRow;
//    }
//
//    public void setStartingRow(int startingRow) {
//        this.startingRow = startingRow;
//    }
//
//    public int getStartingCol() {
//        return startingCol;
//    }
//
//    public void setStartingCol(int startingCol) {
//        this.startingCol = startingCol;
//    }
//
//    public int getMaxColPerRow() {
//        return maxColPerRow;
//    }
//
//    public void setMaxColPerRow(int maxColPerRow) {
//        this.maxColPerRow = maxColPerRow;
//    }
//
//    public int getDeepestRowOnCurrentLevel() {
//        return deepestRowOnCurrentLevel;
//    }
//
//    public void setDeepestRowOnCurrentLevel(int deepestRowOnCurrentLevel) {
//        this.deepestRowOnCurrentLevel = deepestRowOnCurrentLevel;
//    }
//
//    // Method to check if the column limit is exceeded
//    public boolean exceedMaxColPerRow() {
//        return startingCol >= maxColPerRow;
//    }
//
//    // Method to jump to the next available row
//    public void jumpToNextAvailableRow() {
//        this.startingRow = startingRow + deepestRowOnCurrentLevel + 2;
//        this.deepestRowOnCurrentLevel = 0;
//        this.startingCol = 0;
//    }
//
//    // Method to create or get a row
//    public Row createOrGetRow(int rowNum) {
//        Row row = sheet.getRow(rowNum);
//        if (row == null) {
//            row = sheet.createRow(rowNum);
//        }
//        return row;
//    }
//
//    // Getter for the underlying XSSFSheet, if needed
//    public XSSFSheet getUnderlyingSheet() {
//        return sheet;
//    }
//}
