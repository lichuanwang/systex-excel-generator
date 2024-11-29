package com.systex.excelgenerator.utils;

import com.systex.excelgenerator.excel.ExcelSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

public class FormattingAndFilterUtil {

    // 條件式格式
    public static void applyConditionalFormatting(ExcelSheet sheet , int startRow , int endRow , int startCol , int endCol
                                            , String conditionalvalue ){
        // set conditional rule
        // if skill level > 2 -> fill cell background color
        SheetConditionalFormatting sheetcf = sheet.getXssfSheet().getSheetConditionalFormatting();
        ConditionalFormattingRule rule = sheetcf.createConditionalFormattingRule(ComparisonOperator.GT , conditionalvalue);

        // 填充顏色example(也可以改變文字顏色)
        PatternFormatting fill = rule.createPatternFormatting();
        fill.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
        // FontFormatting fontFormat = rule.createFontFormatting();
        // fontFormat.setFontColorIndex(IndexedColors.RED.getIndex());

        CellRangeAddress[] regions = { new CellRangeAddress(startRow, endRow, startCol, endCol)};
        sheetcf.addConditionalFormatting(regions, rule);
    }

    // 凍結儲存格
    public static void freezeCell(ExcelSheet sheet , int startCol , int startRow){
        // 凍結儲存格的某一列到某一列或是某一行到某一行
        sheet.getXssfSheet().createFreezePane(startCol, startRow);
    }

    // 篩選器
    public static void applyCellFilter(ExcelSheet sheet , int startRow , int endRow , int startCol , int endCol){
        sheet.getXssfSheet().setAutoFilter(new CellRangeAddress(startRow, endRow, startCol, endCol));
    }
}
