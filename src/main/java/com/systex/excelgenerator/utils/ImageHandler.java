package com.systex.excelgenerator.utils;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;

public class ImageHandler {
//
//    // 讀取圖片文件 (圖片路徑)
//    FileInputStream inputStream = new FileInputStream("example.jpg");
//    byte[] bytes = IOUtils.toByteArray(inputStream);  // 將圖片讀入 byte array
//    int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);  // 加入圖片，指定類型
//
//    // 釋放圖片文件的資源
//        inputStream.close();
//
//    // 創建繪圖對象
//    Drawing drawing = sheet.createDrawingPatriarch();
//
//    // 創建錨點，設定圖片的位置（起始列和行）
//    ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
//        anchor.setCol1(1);  // 開始的列
//        anchor.setRow1(2);  // 開始的行
//        anchor.setCol2(3);  // 結束的列 (可調整圖片的寬度)
//        anchor.setRow2(9);  // 結束的行 (可調整圖片的高度)
//
//    // 插入圖片
//        drawing.createPicture(anchor, pictureIdx);
}
