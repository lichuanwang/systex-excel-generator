package com.systex.excelgenerator.utils;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 工具類，用於在Excel單元格中設置外部、內部和電子郵件超連結。
 */
public class HyperLinkUtil {

    /**
     * 私有構造方法，防止實例化工具類。
     */
    private HyperLinkUtil(){}

    // 設定文字外部超連結
    /**
     * 為儲存格設置外部超連結。
     * <p>
     * 將儲存格設定為超連結，點擊後可導向指定的外部連結。
     * </p>
     *
     * @param link     要設置的外部連結（URL）
     * @param cell     要設置超連結的儲存格
     * @param workbook 當前的workbook
     * @throws IllegalArgumentException 如果link, cell或workbook為null或link為空
     */
    public static void setHyperLink(String link , Cell cell , Workbook workbook) {
        CreationHelper createHelper = workbook.getCreationHelper();

        // 設定文字中的連結 , 導到設定的連結網頁
        Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.URL);
        hyperlink.setAddress(link);

        cell.setHyperlink(hyperlink);
    }

    // 設定文字內部超連結
    /**
     * 為儲存格設置內部超連結。
     * <p>
     * 將儲存格設定為超連結，點擊後可導向工作表中的指定內部內容。
     * </p>
     *
     * @param sheetname 要連接到的工作表的名稱
     * @param cell      要設置超連結的儲存格
     * @param workbook  當前的workbook
     * @throws IllegalArgumentException 如果sheetname, cell或workbook為null或link為空
     */
    public static void setInternalLink(String sheetname , Cell cell , Workbook workbook) {
        CreationHelper createHelper = workbook.getCreationHelper();

        // 設定文字中的連結 , 導到同個excel的不同sheet
        Hyperlink internallink = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
        internallink.setAddress("'"+sheetname+"'!A1");

        cell.setHyperlink(internallink);
    }

    // 設定文字mail連結
    /**
     * 為儲存格設置電子郵件超連結。
     * <p>
     * 將儲存格設定為郵件連結，點擊後可啟動預設郵件客戶端並撰寫至指定的電子郵件地址。
     * </p>
     *
     * @param email    目標電子郵件地址
     * @param cell     要設置超連結的儲存格
     * @param workbook 當前的workbook
     * @throws IllegalArgumentException 如果email, cell或workbook為null或email為空
     */
    public static void setEmailLink(String email , Cell cell , Workbook workbook){
        CreationHelper createHelper = workbook.getCreationHelper();

        // 設定email的連結 , 點擊可以打開撰寫郵件
        Hyperlink emaillink = createHelper.createHyperlink(HyperlinkType.EMAIL);
        emaillink.setAddress("mailto:" +email);

        cell.setHyperlink(emaillink);
    }
}