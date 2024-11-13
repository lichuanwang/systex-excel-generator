package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public abstract class AbstractDataSection<T> implements Section<T> {
    protected String title;

    public AbstractDataSection(String title) {
        this.title = title;
    }

    // 標題樣式另外寫
    protected CellStyle createTitleStyle(ExcelSheet excelSheet) {
        XSSFSheet sheet = excelSheet.getUnderlyingSheet();
        CellStyle style = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        style.setFont(font);
        return style;
    }

//    // 自訂 cloneFont(只有常見的)
//    protected Font cloneFont(XSSFWorkbook workbook, Font originalFont) {
//        Font newFont = workbook.createFont();
//        newFont.setBold(originalFont.getBold());
//        newFont.setFontHeightInPoints(originalFont.getFontHeightInPoints());
//        newFont.setFontName(originalFont.getFontName());
//        newFont.setColor(originalFont.getColor());
//        newFont.setUnderline(originalFont.getUnderline());
//        newFont.setItalic(originalFont.getItalic());
//        return newFont;
//    }
//
//    protected CellStyle cloneStyle(XSSFWorkbook workbook, CellStyle originalStyle) {
//        CellStyle newStyle = workbook.createCellStyle();
//        newStyle.cloneStyleFrom(originalStyle);
//
//        Font originalFont = workbook.getFontAt(originalStyle.getFontIndex());
//        Font clonedFont = cloneFont(workbook, originalFont);
//        newStyle.setFont(clonedFont);
//        return newStyle;
//    }
//
//    protected CellStyle createSpecialStyle(XSSFWorkbook workbook) {
//        CellStyle specialStyle = workbook.createCellStyle();
//        Font font = workbook.createFont();
//        font.setBold(true);
//        specialStyle.setFont(font);
//
//        specialStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
//        specialStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//
//        return specialStyle;
//    }

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

        CellStyle style = createTitleStyle(excelSheet);
        headerCell.setCellStyle(style);
        // Apply style if needed (e.g., bold, font size)
//        XSSFSheet sheet = excelSheet.getUnderlyingSheet();
//        CellStyle style = sheet.getWorkbook().createCellStyle();
//        Font font = sheet.getWorkbook().createFont();
//        font.setBold(true);
//        font.setFontHeightInPoints((short) 14);
//        style.setFont(font);
//        headerCell.setCellStyle(style);
    }



}
