package com.systex.excelgenerator.utils;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Workbook;

public class HyperLinkUtil {

    private HyperLinkUtil(){}

    // 設定文字外部超連結
    public static void setHyperLink(String link , Cell cell , Workbook workbook) {
        CreationHelper createHelper = workbook.getCreationHelper();

        // 設定文字中的連結 , 導到設定的連結網頁
        Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.URL);
        hyperlink.setAddress(link);

        cell.setHyperlink(hyperlink);
    }

    // 設定文字內部超連結
    public static void setInternalLink(String sheetname , Cell cell , Workbook workbook) {
        CreationHelper createHelper = workbook.getCreationHelper();

        // 設定文字中的連結 , 導到同個excel的不同sheet
        Hyperlink internallink = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
        internallink.setAddress("'"+sheetname+"'!A1");

        cell.setHyperlink(internallink);
    }

    // 設定文字mail連結
    public static void setEmailLink(String email , Cell cell , Workbook workbook){
        CreationHelper createHelper = workbook.getCreationHelper();

        // 設定email的連結 , 點擊可以打開撰寫郵件
        Hyperlink emaillink = createHelper.createHyperlink(HyperlinkType.EMAIL);
        emaillink.setAddress("mailto:" +email);

        cell.setHyperlink(emaillink);
    }
}
