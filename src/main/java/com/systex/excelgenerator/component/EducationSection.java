package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Education;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

public class EducationSection extends Section {

    private List<Education> educations;

    public EducationSection(List<Education> educations) {
        super("Education");
        this.educations = educations;
    }

    @Override
    public int populate(XSSFSheet sheet, int rowNum) {
        addHeader(sheet, rowNum);
        rowNum++;

        Row headerRow = sheet.createRow(rowNum++);
        headerRow.createCell(0).setCellValue("School Name");
        headerRow.createCell(1).setCellValue("Major");
        headerRow.createCell(2).setCellValue("Grade");
        headerRow.createCell(3).setCellValue("Start Date");
        headerRow.createCell(4).setCellValue("End Date");

        for (Education edu : educations) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(edu.getSchoolName());
            row.createCell(1).setCellValue(edu.getMajor());
            row.createCell(2).setCellValue(edu.getGrade());
            row.createCell(3).setCellValue(edu.getStartDate());
            row.createCell(4).setCellValue(edu.getEndDate());
        }
        return rowNum;
    }
}
