package com.systex.excelgenerator.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelFormat {
    // 日期格式
    public static CellStyle DateFormatting(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy/mm/dd"));
        return cellStyle;
    }

    // 文字格式設定 (適用於電話號碼等需要純文字格式的欄位)
    public static CellStyle TextFormatting(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("@"));
        return cellStyle;
    }
}
