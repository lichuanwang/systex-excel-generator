package com.systex.excelgenerator.component;

import com.systex.excelgenerator.style.TemplateStyle;
import com.systex.excelgenerator.style.ExcelFormat;
import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Candidate;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.ss.usermodel.*;
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
    public int getWidth() {
        return 3;
    }

    @Override
    public int getHeight() {
        return 7; // Need to include the section title
    }

    @Override
    protected void populateHeader(ExcelSheet sheet, int startRow, int startCol) {

        String[] headers = {"Name", "Gender", "Birthday", "Phone", "Email", "Address"};

        for (String header : headers) {
            Row row = sheet.createOrGetRow(startRow++);
            Cell cell = row.createCell(startCol);
            cell.setCellValue(header);
        }
    }

    @Override
    protected void populateBody(ExcelSheet sheet, int startRow, int startCol) {

        try {
            addImageToSheet(sheet);
        } catch (IOException e) {
            e.printStackTrace();
        }

        XSSFWorkbook workbook = (XSSFWorkbook) sheet.getUnderlyingSheet().getWorkbook();
        CellStyle cloneStyle = TemplateStyle.createSpecialStyle(workbook);
        CellStyle phoneStyle = ExcelFormat.TextFormatting(workbook);

        // Fill in the data
        Row row = sheet.createOrGetRow(startRow++);
        row.createCell(startCol).setCellValue(candidate.getName());
        row = sheet.createOrGetRow(startRow++);
        row.createCell(startCol).setCellValue(candidate.getGender());
        row = sheet.createOrGetRow(startRow++);
        row.createCell(startCol).setCellValue(SimpleDateFormat.getDateInstance().format(candidate.getBirthday()));
        row = sheet.createOrGetRow(startRow++);
        Cell phoneCell = row.createCell(startCol);
        phoneCell.setCellValue(candidate.getPhone());
        phoneCell.setCellStyle(cloneStyle);
        phoneCell.setCellStyle(phoneStyle);

        row = sheet.createOrGetRow(startRow++);
        Cell emailCell = row.createCell(startCol);
        emailCell.setCellValue(candidate.getEmail());
        emailCell.setCellStyle(cloneStyle);
        row = sheet.createOrGetRow(startRow);
        row.createCell(startCol).setCellValue(candidate.getAddress().toString());
    }

    @Override
    protected void populateFooter(ExcelSheet sheet, int startRow, int startCol){}

    private void addImageToSheet(ExcelSheet sheet) throws IOException {
        // Load the image file
        try (FileInputStream imageStream = new FileInputStream("profile.jpg")) {
            byte[] bytes = IOUtils.toByteArray(imageStream);
            int pictureIdx = sheet.getUnderlyingSheet().getWorkbook().addPicture(bytes, XSSFWorkbook.PICTURE_TYPE_JPEG);

            // Create an anchor to position the image
            XSSFClientAnchor anchor = sheet.getUnderlyingSheet().getWorkbook().getCreationHelper().createClientAnchor();
            anchor.setCol1(sheet.getMaxColPerRow() + 1);
            anchor.setRow1(sheet.getStartingRow());
            anchor.setCol2(sheet.getMaxColPerRow() + 3);
            anchor.setRow2(sheet.getStartingRow() + 7);

            anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);

            // Insert the image into the sheet
            XSSFDrawing drawing = sheet.getUnderlyingSheet().createDrawingPatriarch();
            drawing.createPicture(anchor, pictureIdx);

            System.out.println("Image added");
        }
    }

    @Override
    public void render(ExcelSheet sheet, int startRow, int startCol) {
        addSectionTitle(sheet, startRow, startCol);
        populateHeader(sheet, startRow + 1, startCol);
        populateBody(sheet, startRow + 1, startCol + 1);
        populateFooter(sheet, startRow + 1, startCol + 2);
    }
}



//package com.systex.excelgenerator.component;
//
//import com.systex.excelgenerator.excel.ExcelSheet;
//import com.systex.excelgenerator.model.Candidate;
//import org.apache.commons.compress.utils.IOUtils;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.ClientAnchor;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.*;
//
//import java.io.FileInputStream;
//import java.io.IOException;
//
//import java.text.SimpleDateFormat;
//import java.util.Collection;
//
//public class PersonalInfoSection extends AbstractSection<Candidate> {
//
//    private Candidate candidate;
//
//    public PersonalInfoSection() {
//        super("Personal Information");
//    }
//
//    public void setData(Candidate candidate) {
//        if (candidate != null) {
//            this.candidate = candidate;
//        }
//    }
//
//    @Override
//    public void setData(Collection<Candidate> dataCollection) {
//        if (dataCollection != null && !dataCollection.isEmpty()) {
//            this.candidate = dataCollection.iterator().next();
//        }
//    }
//
//    @Override
//    public boolean isEmpty() {
//        return candidate == null;
//    }
//
//    @Override
//    protected void populateHeader(ExcelSheet sheet) {
//        int rowNum = sheet.getStartingRow() + 1;
//        int colNum = sheet.getStartingCol();
//
//        String[] headers = {"Name", "Gender", "Birthday", "Phone", "Email", "Address"};
//
//        for (String header : headers) {
//            Row row = sheet.createOrGetRow(rowNum++);
//            Cell cell = row.createCell(colNum);
//            cell.setCellValue(header);
//        }
//
//        sheet.setStartingCol(++colNum);
//
//        if (rowNum > sheet.getDeepestRowOnCurrentLevel()) {
//            sheet.setDeepestRowOnCurrentLevel(rowNum);
//        }
//    }
//
//    @Override
//    protected void populateData(ExcelSheet sheet) {
//        int rowNum = sheet.getStartingRow() + 1;
//        int colNum = sheet.getStartingCol();
//
//        try {
//            addImageToSheet(sheet);
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//
//        // Fill in the data
//        Row row = sheet.createOrGetRow(rowNum++);
//        row.createCell(colNum).setCellValue(candidate.getName());
//        row = sheet.createOrGetRow(rowNum++);
//        row.createCell(colNum).setCellValue(candidate.getGender());
//        row = sheet.createOrGetRow(rowNum++);
//        row.createCell(colNum).setCellValue(SimpleDateFormat.getDateInstance().format(candidate.getBirthday()));
//        row = sheet.createOrGetRow(rowNum++);
//        row.createCell(colNum).setCellValue(candidate.getPhone());
//        row = sheet.createOrGetRow(rowNum++);
//        row.createCell(colNum).setCellValue(candidate.getEmail());
//        row = sheet.createOrGetRow(rowNum++);
//        row.createCell(colNum).setCellValue(candidate.getAddress().toString());
//
//        // Set the starting column for the next section
//        sheet.setStartingCol(++colNum);
//
//        // Update the deepest row on current level
//        if (rowNum > sheet.getDeepestRowOnCurrentLevel()) {
//            sheet.setDeepestRowOnCurrentLevel(rowNum);
//        }
//    }
//
//    @Override
//    protected void populateFooter(ExcelSheet sheet) {
////        sheet.setStartingCol(sheet.getStartingCol() + 2);
//    }
//
//    private void addImageToSheet(ExcelSheet sheet) throws IOException {
//        // Load the image file
//        try (FileInputStream imageStream = new FileInputStream("profile.jpg")) {
//            byte[] bytes = IOUtils.toByteArray(imageStream);
//            int pictureIdx = sheet.getUnderlyingSheet().getWorkbook().addPicture(bytes, XSSFWorkbook.PICTURE_TYPE_JPEG);
//
//            // Create an anchor to position the image
//            XSSFClientAnchor anchor = sheet.getUnderlyingSheet().getWorkbook().getCreationHelper().createClientAnchor();
//            anchor.setCol1(sheet.getMaxColPerRow() + 1);
//            anchor.setRow1(sheet.getStartingRow());
//            anchor.setCol2(sheet.getMaxColPerRow() + 3);
//            anchor.setRow2(sheet.getStartingRow() + 7);
//
//            anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);
//
//            // Insert the image into the sheet
//            XSSFDrawing drawing = sheet.getUnderlyingSheet().createDrawingPatriarch();
//            drawing.createPicture(anchor, pictureIdx);
//
//            System.out.println("Image added");
//        }
//    }
//
//
//
//}
