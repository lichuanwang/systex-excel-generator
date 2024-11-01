package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Candidate;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;

import java.text.SimpleDateFormat;
import java.util.Collection;

public class PersonalInfoSection extends AbstractSection<Candidate> {

    private Candidate candidate;

    public PersonalInfoSection() {
        super("Personal Information");
    }

    public void setData(Candidate candidate) {
        if (candidate != null) {
            this.candidate = candidate;
        }
    }

    @Override
    public void setData(Collection<Candidate> dataCollection) {
        if (dataCollection != null && !dataCollection.isEmpty()) {
            this.candidate = dataCollection.iterator().next();
        }
    }

    @Override
    public boolean isEmpty() {
        return candidate == null;
    }

    @Override
    protected int generateHeader(XSSFSheet sheet, int rowNum) {
        int initialRowNum = rowNum;

        Row row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Name");
        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Gender");
        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Birthday");
        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Phone");
        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Email");
        row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue("Address");

        return initialRowNum;
    }

    @Override
    protected int generateData(XSSFSheet sheet, int rowNum) {
        //Add image in the third column of the "Personal Information" section
        try {
            addImageToSheet(sheet, rowNum, 5); // Position image in the third column of the current row
        } catch (IOException e) {
            e.printStackTrace();
        }
        Row row = sheet.getRow(rowNum++);
        row.createCell(1).setCellValue(candidate.getName());
        row = sheet.getRow(rowNum++);
        row.createCell(1).setCellValue(candidate.getGender());
        row = sheet.getRow(rowNum++);
        row.createCell(1).setCellValue(SimpleDateFormat.getDateInstance().format(candidate.getBirthday()));
        row = sheet.getRow(rowNum++);
        row.createCell(1).setCellValue(candidate.getPhone());
        row = sheet.getRow(rowNum++);
        row.createCell(1).setCellValue(candidate.getEmail());
        row = sheet.getRow(rowNum++);
        row.createCell(1).setCellValue(candidate.getAddress().toString());

        return rowNum;
    }

    @Override
    protected int generateFooter(XSSFSheet sheet, int rowNum) {
        return rowNum;
    }


    private void addImageToSheet(XSSFSheet sheet, int rowNum, int colNum) throws IOException {
        // Load the image file
        try (FileInputStream imageStream = new FileInputStream("profile.jpg")) {
            byte[] bytes = IOUtils.toByteArray(imageStream);
            int pictureIdx = sheet.getWorkbook().addPicture(bytes, XSSFWorkbook.PICTURE_TYPE_JPEG);

            // Create an anchor to position the image
            XSSFClientAnchor anchor = sheet.getWorkbook().getCreationHelper().createClientAnchor();
            anchor.setCol1(colNum);
            anchor.setRow1(rowNum);
            anchor.setCol2(colNum + 2);
            anchor.setRow2(rowNum + 6);

            anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);

            // Insert the image into the sheet
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            Picture picture = drawing.createPicture(anchor, pictureIdx);

            System.out.println("Image added");
        }
    }



}
