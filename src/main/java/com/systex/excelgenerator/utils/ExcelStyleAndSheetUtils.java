package com.systex.excelgenerator.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class ExcelStyleAndSheetUtils {
    // 日期格式
    public static CellStyle dateFormatting(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy/mm/dd"));
        return cellStyle;
    }

    // 文字格式設定 (適用於電話號碼等需要純文字格式的欄位)
    public static CellStyle textFormatting(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("@"));
        return cellStyle;
    }

    // 隱藏列 (可以多選)
    public static void hideColumns(Sheet sheet, boolean isRange, int... columnIndices) {
        if (isRange) {
            int startColumn = columnIndices[0];
            int endColumn = columnIndices[1];
            for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++) {
                sheet.setColumnHidden(columnIndex, true);
            }
        } else {
            for (int columnIndex : columnIndices) {
                sheet.setColumnHidden(columnIndex, true);
            }
        }
    }

    // 指定保護Sheet
    public static void protectSheet(XSSFSheet sheet, String password) {
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet cannot be null.");
        }
        sheet.protectSheet(password);
    }
}


