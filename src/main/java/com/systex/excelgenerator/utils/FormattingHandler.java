package com.systex.excelgenerator.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.time.LocalDate;

public class FormattingHandler {

    // 格式設定
    // date format
    public CellStyle DateFormatting(LocalDate date, Workbook workbook) {
        // 共用workbook , cellStyle
        CellStyle cellstyle = workbook.createCellStyle();

        // 日期格式設定:yyyy/mm/dd
        //CreationHelper createHelper = workbook.getCreationHelper();
        cellstyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy/mm/dd"));

        return cellstyle;
    }

    // 文字格式設定(ex: phone number)
    public CellStyle TextFormatting(Workbook workbook){
        // 共用workbook , sheet , cellStyle ...blablabla
        CellStyle cellstyle = workbook.createCellStyle();

        // "@" : 文字格式
        cellstyle.setDataFormat(workbook.createDataFormat().getFormat("@"));

        return cellstyle;
    }

    // 條件式格式
    public void ConditionalFormatting(Sheet sheet , String conditionalvalue ,
                                      int startRow , int endRow , int col){
        // set conditional rule
        // if skill level > 2 -> fill cell background color
        SheetConditionalFormatting sheetcf = sheet.getSheetConditionalFormatting();
        ConditionalFormattingRule rule = sheetcf.createConditionalFormattingRule(ComparisonOperator.GT , conditionalvalue);

        // 填充顏色example(也可以改變文字顏色)
        PatternFormatting fill = rule.createPatternFormatting();
        fill.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
        // FontFormatting fontFormat = rule.createFontFormatting();
        // fontFormat.setFontColorIndex(IndexedColors.RED.getIndex());

        // 設定條件式cell範圍
        String startCellRef = new CellReference(startRow, col).formatAsString();
        String endCellRef = new CellReference(endRow, col).formatAsString();
        String range = startCellRef + ":" + endCellRef;

        CellRangeAddress[] regions = { CellRangeAddress.valueOf(range) };
        sheetcf.addConditionalFormatting(regions, rule);
    }

    // 凍結儲存格
    public void freezeCell(Sheet sheet , int first, int last){
        // 凍結儲存格的某一列到某一列
        sheet.createFreezePane(first, last);
    }
}
