package com.systex.excelgenerator.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.time.LocalDate;

public class FormattingHandler {

    // 格式設定
    // date format
    public CellStyle DateFormatting(LocalDate date, Workbook workbook) {
        // 共用workbook , cellStyle
        CellStyle cellstyle = workbook.createCellStyle();

        // 日期格式設定:yyyy/mm/dd
        CreationHelper createHelper = workbook.getCreationHelper();
        cellstyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy/mm/dd"));
        return cellstyle;
    }

    // 文字格式設定(ex: phone number)
    public CellStyle TextFormatting(String text , Workbook workbook){
        // 共用workbook , sheet , cellStyle ...blablabla
        CellStyle cellstyle = workbook.createCellStyle();

        // "@" : 文字格式
        cellstyle.setDataFormat(workbook.createDataFormat().getFormat("@"));

        return cellstyle;
    }

    // 條件式格式
    public void ConditionalFormatting(){
        // input data : range , conditional value

        // 共用workbook , sheet
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();

        // set conditional rule
        // example : if value > 100 fill color
        SheetConditionalFormatting sheetcf = sheet.getSheetConditionalFormatting();
        ConditionalFormattingRule rule = sheetcf.createConditionalFormattingRule(ComparisonOperator.GT , "100");

        // 填充顏色example(也可以改變文字顏色)
        PatternFormatting fill = rule.createPatternFormatting();
        fill.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());

        // 設定條件式cell範圍
        CellRangeAddress[] regions = { CellRangeAddress.valueOf("A1:A10") };
        sheetcf.addConditionalFormatting(regions, rule);
    }

    // 合併儲存格
    public void mergeCell(int firstRow, int lastRow, int firstCol, int lastCol){
        // 共用workbook , sheet
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();

        // set merge cell region
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }

    // 凍結窗格
    public void freezeCell(int firstRow, int lastRow, int firstCol, int lastCol){
        // 共用workbook , sheet
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();

        // 凍結儲存格
        sheet.createFreezePane(firstRow, lastRow, firstCol, lastCol);
    }
}
