package com.systex.excelgenerator.utils;

import org.apache.poi.ss.usermodel.*;

public class ExcelStyleUtils {
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

    // 指定保護儲存格範圍
        public static void protectSheet(Sheet sheet, String password) {
            sheet.protectSheet(password);
        }
}

