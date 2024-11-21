package com.systex.excelgenerator.style;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class StyleTemplate {

    public static CellStyle createCommonStyle(XSSFWorkbook workbook) {
        CellStyle specialStyle = workbook.createCellStyle();

        Font font = workbook.createFont();
        font.setBold(true);
        specialStyle.setFont(font);

        specialStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        specialStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        return specialStyle;
    }

    public static CellStyle createCommonTitleStyle(XSSFWorkbook workbook) {
        CellStyle titleStyle = workbook.createCellStyle();

        Font font1 = workbook.createFont();
        font1.setBold(true);
        titleStyle.setFont(font1);

        font1.setFontHeightInPoints((short)14);

        return titleStyle;
    }
}
