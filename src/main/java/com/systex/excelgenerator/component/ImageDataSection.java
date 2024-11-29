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

    private static final Logger log = LogManager.getLogger(ImageDataSection.class);
    private String imageType;

    public ImageDataSection() {
        super("Image");
    }

    public void setImageType(String imageType) {
        this.imageType = imageType;
    }

    @Override
    protected void renderHeader(ExcelSheet sheet, int startRow, int startCol) {
        // add header related content if necessary
    }

    @Override
    protected void renderBody(ExcelSheet sheet, int startRow, int startCol) {

        for (String imagePath : content) {

            // 讀取圖片
            try (FileInputStream inputStream = new FileInputStream(imagePath)) {
                byte[] bytes = IOUtils.toByteArray(inputStream);
                int pictureIndex = sheet.getWorkbook().addPicture(bytes, converetImageType(sheet , imageType));

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
        // add footer related content if necessary
    }

    protected int converetImageType(ExcelSheet sheet , String imageType) {
        switch (imageType.toLowerCase()) {
            case "emf":
                return sheet.getWorkbook().PICTURE_TYPE_EMF;
            case "wnf":
                return sheet.getWorkbook().PICTURE_TYPE_WMF;
            case "pict":
                return sheet.getWorkbook().PICTURE_TYPE_PICT;
            case "jpeg":
                return sheet.getWorkbook().PICTURE_TYPE_JPEG;
            case "png":
                return sheet.getWorkbook().PICTURE_TYPE_PNG;
            case "dib":
                return sheet.getWorkbook().PICTURE_TYPE_DIB;
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
        return 3;
    }

    @Override
    public int getHeight() {
        return 3;
    }
}
