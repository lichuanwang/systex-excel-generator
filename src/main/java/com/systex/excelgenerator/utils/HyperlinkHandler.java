package com.systex.excelgenerator.utils;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HyperlinkHandler {

    private CreationHelper createHelper;

    // 設定文字外部超連結
    public void setHyperLink(String text , String link){
        // 要設定共用的workbook...blablabla
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);

        // 設定文字中的連結 , 導道設定的連結網頁
        cell.setCellValue(text);
        Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.URL);
        hyperlink.setAddress(link);
        applyLink(hyperlink);
    }

    // 設定文字內部超連結
    public void setInternalLink(String sheetname , String text){
        // 要設定共用的workbook...blablabla
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);

        // 設定文字中的連結 , 導到同個excel的不同sheet
        cell.setCellValue(text);
        Hyperlink internallink = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
        internallink.setAddress("'"+sheetname+"'!A1");
        applyLink(internallink);
    }

    public void applyLink(Hyperlink link){
        // 要設定共用的workbook...blablabla
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);

        cell.setHyperlink(link);
    }
}
