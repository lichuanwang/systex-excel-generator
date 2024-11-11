package com.systex.excelgenerator.style;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CustomStyle {

    /**
     * 創建帶黃色填充的自訂樣式
     */
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
