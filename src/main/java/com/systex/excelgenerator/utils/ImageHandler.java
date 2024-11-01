package com.systex.excelgenerator.utils;

import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;

public class ImageHandler {

    public void insertImage(Sheet sheet , int col , int row , String imagepath){

        // 讀取圖片
        try (FileInputStream inputStream = new FileInputStream(imagepath)) {
            byte[] bytes = IOUtils.toByteArray(inputStream);
            int pictureIndex = sheet.getWorkbook().addPicture(bytes, Workbook.PICTURE_TYPE_PNG);

            // 創建畫布
            Drawing<?> drawing = sheet.createDrawingPatriarch();

            // 設置圖片錨點位置，並將圖片設置為跟隨儲存格移動但不改變大小
            ClientAnchor anchor = sheet.getWorkbook().getCreationHelper().createClientAnchor();
            anchor.setCol1(col+1);
            anchor.setRow1(row+1);
            //anchor.setCol1(col+1);
            //anchor.setCol2(col+3); // 指定範圍的列
            //anchor.setRow1(row+1);
            //anchor.setRow2(row+3); // 指定範圍的行
            anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);

            // 插入圖片
            Picture picture = drawing.createPicture(anchor, pictureIndex);

            // 設定圖片大小 (以像素為單位)
            double widthRatio = 0.225;
            double heightRatio = 0.28;
            picture.resize(widthRatio, heightRatio);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
