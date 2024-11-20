package com.systex.excelgenerator.component;

import com.systex.excelgenerator.style.StyleTemplate;
import com.systex.excelgenerator.style.ExcelFormat;
import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Candidate;
import com.systex.excelgenerator.utils.DataValidationHandler;
import com.systex.excelgenerator.utils.HyperlinkHandler;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.text.DateFormat;
import java.util.Collection;

public class PersonalInfoDataSection extends AbstractDataSection<Candidate> {

    private static final Logger log = LogManager.getLogger(PersonalInfoDataSection.class);
    private Candidate candidate;
    private HyperlinkHandler hyperlinkHandler = new HyperlinkHandler();

    public PersonalInfoDataSection() {
        super("Personal Information");
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
    protected void renderHeader(ExcelSheet sheet, int startRow, int startCol) {

        String[] headers = {"Name", "Gender", "Birthday", "Phone", "Email", "Address"};

        for (String header : headers) {
            Row row = sheet.createOrGetRow(startRow++);
            Cell cell = row.createCell(startCol);
            cell.setCellValue(header);
        }
    }

    @Override
    protected void renderBody(ExcelSheet sheet, int startRow, int startCol) {

        XSSFWorkbook workbook = (XSSFWorkbook) sheet.getWorkbook();
        CellStyle cloneStyle = StyleTemplate.createCommonStyle(workbook);
        CellStyle phoneStyle = ExcelFormat.TextFormatting(workbook);

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
        Cell phoneCell = row.createCell(startCol);
        phoneCell.setCellValue(candidate.getPhone());
        phoneCell.setCellStyle(cloneStyle);
        phoneCell.setCellStyle(phoneStyle);

        row = sheet.createOrGetRow(startRow++);
        Cell emailCell = row.createCell(startCol);
        emailCell.setCellValue(candidate.getEmail());
        emailCell.setCellStyle(cloneStyle);
        row.createCell(startCol).setCellValue(candidate.getEmail());
        // Set Email HyperLink
        hyperlinkHandler.setEmailLink(candidate.getEmail(), row.getCell(startCol) , sheet.getWorkbook());

        row = sheet.createOrGetRow(startRow);
        row.createCell(startCol).setCellValue(candidate.getAddress().toString());
    }

    @Override
    protected void renderFooter(ExcelSheet sheet, int startRow, int startCol) {
        // implement footer logic here
    }

    @Override
    public void render(ExcelSheet sheet, int startRow, int startCol) {
        addSectionTitle(sheet, startRow, startCol);
        renderHeader(sheet, startRow + 1, startCol);
        renderBody(sheet, startRow + 1, startCol + 1);
        renderFooter(sheet, startRow + 1, startCol + 2);
    }
}