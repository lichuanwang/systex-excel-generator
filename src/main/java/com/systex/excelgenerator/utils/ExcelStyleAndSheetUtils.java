package com.systex.excelgenerator.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 *  工具類，包含格式化日期與文字、隱藏列與保護工作表的功能。
 */

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

    /**
     * 隱藏指定一到多個列
     *
     * @param sheet         指定的工作表，不能為空
     * @param columnIndices 要隱藏的列(從0開始)
     * @throws IllegalArgumentException 如果 sheet 為 null 或空、列為空
     */
    // 隱藏列 (一到多)
    public static void hideColumns(Sheet sheet, int... columnIndices) {
        if (sheet == null) {
            throw new IllegalArgumentException("sheet不可以為null");
        }
        if (columnIndices == null || columnIndices.length == 0) {
            throw new IllegalArgumentException("列不可為null或空");
        }
        for (int columnIndex : columnIndices) {
            sheet.setColumnHidden(columnIndex, true);
        }
    }

    /**
     * 隱藏指定列的範圍
     *
     * @param sheet       指定的工作表，不可為 null 或空
     * @param startColumn 要隱藏的起始列 (包含本身，從0開始)
     * @param endColumn   要隱藏的結束列 (包含本身，從0開始)
     * @throws IllegalArgumentException 如果 sheet 為 null 或空、範圍無效
     */

    // 隱藏列 (範圍)
    public static void hideColumnRange(Sheet sheet, int startColumn, int endColumn) {
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet不可以為null");
        }
        if (startColumn < 0 || endColumn < startColumn) {
            throw new IllegalArgumentException(startColumn + "到" + endColumn + "是無效的範圍");
        }
        for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++) {
            sheet.setColumnHidden(columnIndex, true);
        }
    }

    /**
     * 設定保護工作表與解鎖密碼
     *
     * @param sheet    要保護的工作表，不可為 null 或空
     * @param password 解鎖工作表的密碼，不能為 null 或空
     * @throws SheetProtectionException 如果 sheet 或密碼為 null 或空值
     */

    // 指定保護Sheet
    public static void protectSheet(XSSFSheet sheet, String password) {
        if (sheet == null) {
            throw new SheetProtectionException("Sheet不可以為空");
        }
        if (password == null || password.isEmpty()) {
            throw new SheetProtectionException("密碼不可以為null或空");
        }
        sheet.protectSheet(password);
    }

    /**
     * 自訂異常狀況，用於鎖定工作表相關的異常狀況
     */
    public static class SheetProtectionException extends RuntimeException {
        /**
         * 自行指定錯誤訊息。
         *
         * @param message 錯誤訊息
         */
        public SheetProtectionException(String message) {
            super(message);
        }
        /**
         * 自行指定錯誤訊息和導致例外的原因。
         *
         * @param message 錯誤訊息
         * @param cause   導致此例外的原因
         */
        public SheetProtectionException(String message, Throwable cause) {
            super(message, cause);
        }
    }
}


