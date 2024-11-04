package com.systex.excelgenerator.utils;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;

public class HyperlinkHandler {

    private CreationHelper createHelper;

    // 設定文字外部超連結
    public void setHyperLink(String link , Cell cell , Workbook workbook) {
        createHelper = workbook.getCreationHelper();

        // 設定文字中的連結 , 導到設定的連結網頁
        Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.URL);
        hyperlink.setAddress(link);
        cell.setHyperlink(hyperlink);
    }

    // 設定文字內部超連結
    public void setInternalLink(String sheetname ,Cell cell , Workbook workbook) {
        createHelper = workbook.getCreationHelper();

        // 設定文字中的連結 , 導到同個excel的不同sheet
        Hyperlink internallink = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
        internallink.setAddress("'"+sheetname+"'!A1");
        cell.setHyperlink(internallink);
    }

    // 設定文字mail連結
    public void setEmailLink(){

    }
}
