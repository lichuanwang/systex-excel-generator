package com.systex.excelgenerator.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import java.time.LocalDate;

public class FormattingAndFilter {

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
    public void freezeCell(Sheet sheet , int startCol , int startRow){
        // 凍結儲存格的某一列到某一列或是某一行到某一行
        sheet.createFreezePane(startCol, startRow);
    }

    // 篩選器
    public void CellFilter (Sheet sheet , int startRow ,  int endRow , int startCol , int endCol){
        sheet.setAutoFilter(new CellRangeAddress(startRow, endRow, startCol, endCol));
    }
}
