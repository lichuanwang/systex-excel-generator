package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Education;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

public class EducationSection extends AbstractSection<Education> {

    private List<Education> educations;

    public EducationSection() {
        super("Education");
    }

    @Override
    protected int generateHeader(XSSFSheet sheet, int rowNum) {
        Row headerRow = sheet.createRow(rowNum++);
        headerRow.createCell(0).setCellValue("School Name");
        headerRow.createCell(1).setCellValue("Major");
        headerRow.createCell(2).setCellValue("Grade");
        headerRow.createCell(3).setCellValue("Start Date");
        headerRow.createCell(4).setCellValue("End Date");

        return rowNum;
    }

    @Override
    protected int generateData(XSSFSheet sheet, int rowNum) {
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

    @Override
    protected int generateFooter(XSSFSheet sheet, int rowNum) {
        return rowNum;
    }



    @Override
    public void setData(Education data) {
        if( educations != null ) {
            this.educations = Arrays.asList(data); // Check if this will return the same thing just like the one below
        }
    }

    @Override
    public void setData(Collection<Education> dataCollection) {
        if (dataCollection != null && !dataCollection.isEmpty()) {
            this.educations = new ArrayList<>(dataCollection);
        }
    }

    @Override
    public boolean isEmpty() {
        return educations == null || educations.isEmpty();
    }
//    @Override
//    public int populate(XSSFSheet sheet, int rowNum) {
//        addHeader(sheet, rowNum);
//        rowNum++;
//
//        Row headerRow = sheet.createRow(rowNum++);
//        headerRow.createCell(0).setCellValue("School Name");
//        headerRow.createCell(1).setCellValue("Major");
//        headerRow.createCell(2).setCellValue("Grade");
//        headerRow.createCell(3).setCellValue("Start Date");
//        headerRow.createCell(4).setCellValue("End Date");
//
//        for (Education edu : educations) {
//            Row row = sheet.createRow(rowNum++);
//            row.createCell(0).setCellValue(edu.getSchoolName());
//            row.createCell(1).setCellValue(edu.getMajor());
//            row.createCell(2).setCellValue(edu.getGrade());
//            row.createCell(3).setCellValue(edu.getStartDate());
//            row.createCell(4).setCellValue(edu.getEndDate());
//        }
//        return rowNum;
//    }
}
