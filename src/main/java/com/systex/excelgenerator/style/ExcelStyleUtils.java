package com.systex.excelgenerator.style;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelStyleUtils {

    // 自訂 cloneFont(只有常見的)
    public static Font cloneFont(XSSFWorkbook workbook, Font originalFont) {
        Font newFont = workbook.createFont();
        newFont.setBold(originalFont.getBold());
        newFont.setFontHeightInPoints(originalFont.getFontHeightInPoints());
        newFont.setFontName(originalFont.getFontName());
        newFont.setColor(originalFont.getColor());
        newFont.setUnderline(originalFont.getUnderline());
        newFont.setItalic(originalFont.getItalic());
        return newFont;
    }

    public static CellStyle cloneStyle(XSSFWorkbook workbook, CellStyle originalStyle) {
        CellStyle newStyle = workbook.createCellStyle();
        newStyle.cloneStyleFrom(originalStyle);

        Font originalFont = workbook.getFontAt(originalStyle.getFontIndex());
        Font clonedFont = cloneFont(workbook, originalFont);
        newStyle.setFont(clonedFont);
        return newStyle;
    }

    public static CellStyle createSpecialStyle(XSSFWorkbook workbook) {
        CellStyle specialStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        specialStyle.setFont(font);

        specialStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        specialStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        return specialStyle;
    }
}