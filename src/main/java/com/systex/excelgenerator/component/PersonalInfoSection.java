package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Candidate;
import com.systex.excelgenerator.utils.DataValidationHandler;
import com.systex.excelgenerator.utils.FormattingHandler;
import com.systex.excelgenerator.utils.HyperlinkHandler;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;

import java.text.DateFormat;
import java.util.Collection;

public class PersonalInfoSection extends AbstractSection<Candidate> {

    private static final Logger log = LogManager.getLogger(PersonalInfoSection.class);
    private Candidate candidate;
    private FormattingHandler formattingHandler = new FormattingHandler();
    private HyperlinkHandler hyperlinkHandler = new HyperlinkHandler();

    public PersonalInfoSection() {
        super("Personal Information");
    }

    @Override
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
        return 3; //
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

        // Fill in the data
        Row row = sheet.createOrGetRow(startRow++);
        row.createCell(startCol).setCellValue(candidate.getName());
        row = sheet.createOrGetRow(startRow++);
        row.createCell(startCol).setCellValue(candidate.getGender());

        // test data valid - gender : male/female
        DataValidationHandler dataValidationHandler = new DataValidationHandler(sheet.getXssfSheet() , row.getRowNum(), row.getRowNum(), startCol, startCol);
        String[] options = {"Male","Female"};
        dataValidationHandler.ListDataValid(options);

        row = sheet.createOrGetRow(startRow++);
        row.createCell(startCol).setCellValue(DateFormat.getDateInstance().format(candidate.getBirthday()));
        row = sheet.createOrGetRow(startRow++);
        row.createCell(startCol).setCellValue(candidate.getPhone());

        // format phone number
        row.getCell(startCol).setCellStyle(formattingHandler.TextFormatting(sheet.getWorkbook()));

        row = sheet.createOrGetRow(startRow++);
        row.createCell(startCol).setCellValue(candidate.getEmail());

        // Set Email HyperLink
        System.out.println(candidate.getEmail());
        hyperlinkHandler.setEmailLink(candidate.getEmail(), row.getCell(startCol) , sheet.getWorkbook());

        row = sheet.createOrGetRow(startRow);
        row.createCell(startCol).setCellValue(candidate.getAddress().toString());
    }

    @Override
    protected void populateFooter(ExcelSheet sheet, int startRow, int startCol){
        // implement footer logic here
    }

    private void addImageToSheet(ExcelSheet sheet) throws IOException {
        // Load the image file
        try (FileInputStream imageStream = new FileInputStream("profile.jpg")) {
            byte[] bytes = IOUtils.toByteArray(imageStream);
            int pictureIdx = sheet.getXssfSheet().getWorkbook().addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);

            // Create an anchor to position the image
            XSSFClientAnchor anchor = sheet.getXssfSheet().getWorkbook().getCreationHelper().createClientAnchor();
            anchor.setCol1(sheet.getMaxColPerRow() + 1);
            anchor.setRow1(sheet.getStartingRow());
            anchor.setCol2(sheet.getMaxColPerRow() + 3);
            anchor.setRow2(sheet.getStartingRow() + 7);

            anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);

            // Insert the image into the sheet
            XSSFDrawing drawing = sheet.getXssfSheet().createDrawingPatriarch();
            drawing.createPicture(anchor, pictureIdx);

            log.info("Image added");
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