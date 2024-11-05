package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Education;
import org.apache.poi.ss.usermodel.Row;

public class EducationSection extends AbstractSection<Education> {

    public EducationSection() {
        super("Education");
    }

    @Override
    public boolean isEmpty() {
        return content == null || content.isEmpty();
    }

    @Override
    public int getWidth() {
        // Set the width based on the number of columns this section uses.
        return 6; // Example width, assuming we have 5 columns for education details plus one additional column to separate different section
    }

    @Override
    public int getHeight() {
        // Height based on the number of education entries
        return content.size() + 1; // +1 for the header row
    }

    protected void populateHeader(ExcelSheet sheet, int startRow, int startCol) {
        // Create header row for Education section
        Row headerRow = sheet.createOrGetRow(startRow);
        headerRow.createCell(startCol).setCellValue("School Name");
        headerRow.createCell(startCol + 1).setCellValue("Major");
        headerRow.createCell(startCol + 2).setCellValue("Grade");
        headerRow.createCell(startCol + 3).setCellValue("Start Date");
        headerRow.createCell(startCol + 4).setCellValue("End Date");
    }

    protected void populateBody(ExcelSheet sheet, int startRow, int startCol) {
        int rowNum = startRow; // Start from the row after the header

        for (Education edu : content) {
            Row row = sheet.createOrGetRow(rowNum++);
            row.createCell(startCol).setCellValue(edu.getSchoolName());
            row.createCell(startCol + 1).setCellValue(edu.getMajor());
            row.createCell(startCol + 2).setCellValue(edu.getGrade());
            row.createCell(startCol + 3).setCellValue(edu.getStartDate());
            row.createCell(startCol + 4).setCellValue(edu.getEndDate());
        }
    }

    protected void populateFooter(ExcelSheet sheet, int startRow, int startCol) {
        // implement footer logic here
    }
}