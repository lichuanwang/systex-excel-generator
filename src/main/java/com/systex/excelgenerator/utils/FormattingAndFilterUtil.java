package com.systex.excelgenerator.utils;

import com.systex.excelgenerator.excel.ExcelSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * 工具類，用於設定條件式格式、凍結儲存格與篩選功能。
 */
public class FormattingAndFilterUtil {

    // 條件式格式
    /**
     * 套用條件式格式。
     * <p>
     * 根據條件限制，對儲存格做樣式改變。
     * </p>
     *
     * @param sheet            應用條件格式的工作表
     * @param startRow         條件格式應用的開始行數
     * @param endRow           條件格式應用的結束行數
     * @param startCol         條件格式應用的開始列數
     * @param endCol           條件格是應用的結束列數
     * @param conditionalvalue 條件格式的比較值
     * @throws IllegalArgumentException 如果參數無效或範圍錯誤
     */
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
    /**
     * 凍結儲存格。
     * <p>
     * 凍結儲存格的行或列，將工作表中的某些行固定在可視範圍內。
     * </p>
     *
     * @param sheet    凍結行列的工作表
     * @param startCol 開始凍結的行
     * @param startRow 開始凍結的列
     * @throws IllegalArgumentException 如果sheet為null或索引無效
     */
    public static void freezeCell(ExcelSheet sheet , int startCol , int startRow){
        // 凍結儲存格的某一列到某一列或是某一行到某一行
        sheet.getXssfSheet().createFreezePane(startCol, startRow);
    }

    // 篩選器
    /**
     * 套用篩選器。
     * <p>
     * 用於在指定範圍內啟用篩選器功能，方便用戶進行數據過濾。
     * </p>
     *
     * @param sheet    設定篩選器的工作表
     * @param startRow 設定篩選器開始的列數
     * @param endRow   設定篩選器結束的列數
     * @param startCol 設定篩選器開始的行數
     * @param endCol   設定篩選器結束的行數
     * @throws IllegalArgumentException 如果sheet為null或範圍無效
     */
    public static void applyCellFilter(ExcelSheet sheet , int startRow , int endRow , int startCol , int endCol){
        sheet.getXssfSheet().setAutoFilter(new CellRangeAddress(startRow, endRow, startCol, endCol));
    }
}
