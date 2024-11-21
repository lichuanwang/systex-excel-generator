package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;

import java.io.FileInputStream;
import java.io.IOException;

public class ImageDataSection extends AbstractDataSection<String> {

    private static final Logger log = LogManager.getLogger(PersonalInfoDataSection.class);
    private String imageType;

    public ImageDataSection() {
        super("Image");
    }

    public void setImageType(String imageType) {
        this.imageType = imageType;
    }

    @Override
    protected void renderHeader(ExcelSheet sheet, int startRow, int startCol) {

    }

    @Override
    protected void renderBody(ExcelSheet sheet, int startRow, int startCol) {

        for (String imagepath : content) {

            // 讀取圖片
            try (FileInputStream inputStream = new FileInputStream(imagepath)) {
                byte[] bytes = IOUtils.toByteArray(inputStream);
                int pictureIndex = sheet.getWorkbook().addPicture(bytes, converetImageType(imageType));

                XSSFClientAnchor anchor = sheet.getXssfSheet().getWorkbook().getCreationHelper().createClientAnchor();
                anchor.setCol1(startCol);
                anchor.setRow1(startRow);
                anchor.setCol2(startCol + 3);
                anchor.setRow2(startRow + 7);

                // Insert the image into the sheet
                XSSFDrawing drawing = sheet.getXssfSheet().createDrawingPatriarch();
                drawing.createPicture(anchor, pictureIndex);


                log.info("Image added");

            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    @Override
    protected void renderFooter(ExcelSheet sheet, int startRow, int startCol) {

    }

    protected int converetImageType(String imageType) {
        String imagetype = imageType.toLowerCase();
        switch (imagetype) {
            case "emf":
                return 2;
            case "wnf":
                return 3;
            case "pict":
                return 4;
            case "jpeg":
                return 5;
            case "png":
                return 6;
            case "dib":
                return 7;
            default:
                throw new IllegalArgumentException("沒有這種圖片類型: " + imageType);
        }
    }

    @Override
    public boolean isEmpty() {
        return false;
    }

    @Override
    public int getWidth() {
        return 0;
    }

    @Override
    public int getHeight() {
        return 0;
    }
}
