package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Education;
import org.apache.poi.ss.usermodel.Row;

import java.util.*;

public class EducationSection extends AbstractSection<Education> {

    private List<Education> educations;

    public EducationSection() {
        super("Education");
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

    @Override
    public int getWidth() {
        // Set the width based on the number of columns this section uses.
        return 6; // Example width, assuming we have 5 columns for education details plus one additional column to separate different section
    }

    @Override
    public int getHeight() {
        // Height based on the number of education entries
        return educations.size() + 2; // +2 for the header row and extra row space
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

        for (Education edu : educations) {
            Row row = sheet.createOrGetRow(rowNum++);
            row.createCell(startCol).setCellValue(edu.getSchoolName());
            row.createCell(startCol + 1).setCellValue(edu.getMajor());
            row.createCell(startCol + 2).setCellValue(edu.getGrade());
            row.createCell(startCol + 3).setCellValue(edu.getStartDate());
            row.createCell(startCol + 4).setCellValue(edu.getEndDate());
        }
    }

    protected void populateFooter(ExcelSheet sheet, int startRow, int startCol) {

    }
}


//package com.systex.excelgenerator.component;
//
//import com.systex.excelgenerator.excel.ExcelSheet;
//import com.systex.excelgenerator.model.Education;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//
//import java.util.ArrayList;
//import java.util.Arrays;
//import java.util.Collection;
//import java.util.List;
//
//public class EducationSection extends AbstractSection<Education> {
//
//    private List<Education> educations;
//
//    public EducationSection() {
//        super("Education");
//    }
//
//    @Override
//    protected void populateHeader(ExcelSheet sheet) {
//        int rowNum = sheet.getStartingRow() + 1;
//        int colNum = sheet.getStartingCol();
//        Row headerRow = sheet.createOrGetRow(rowNum++);
//        headerRow.createCell(colNum++).setCellValue("School Name");
//        headerRow.createCell(colNum++).setCellValue("Major");
//        headerRow.createCell(colNum++).setCellValue("Grade");
//        headerRow.createCell(colNum++).setCellValue("Start Date");
//        headerRow.createCell(colNum).setCellValue("End Date");
//
//        sheet.setStartingRow(rowNum);
//
//        // Update the deepest row on current level
//        if (rowNum > sheet.getDeepestRowOnCurrentLevel()) {
//            sheet.setDeepestRowOnCurrentLevel(rowNum);
//        }
//    }
//
//    @Override
//    protected void populateData(ExcelSheet sheet) {
//        int rowNum = sheet.getStartingRow();
//        int colNum = sheet.getStartingCol();
//        for (Education edu : educations) {
//            Row row = sheet.createOrGetRow(rowNum++);
//            row.createCell(colNum++).setCellValue(edu.getSchoolName());
//            row.createCell(colNum++).setCellValue(edu.getMajor());
//            row.createCell(colNum++).setCellValue(edu.getGrade());
//            row.createCell(colNum++).setCellValue(edu.getStartDate());
//            row.createCell(colNum++).setCellValue(edu.getEndDate());
//        }
//
//        sheet.setStartingCol(colNum);
//
//        if(rowNum > sheet.getDeepestRowOnCurrentLevel()) {
//            sheet.setDeepestRowOnCurrentLevel(rowNum);
//        }
//
//    }
//
//    @Override
//    protected void populateFooter(ExcelSheet sheet) {
////        sheet.setStartingCol(sheet.getStartingCol() + 2);
//    }
//
//
//
//    @Override
//    public void setData(Education data) {
//        if( educations != null ) {
//            this.educations = Arrays.asList(data); // Check if this will return the same thing just like the one below
//        }
//    }
//
//    @Override
//    public void setData(Collection<Education> dataCollection) {
//        if (dataCollection != null && !dataCollection.isEmpty()) {
//            this.educations = new ArrayList<>(dataCollection);
//        }
//    }
//
//    @Override
//    public boolean isEmpty() {
//        return educations == null || educations.isEmpty();
//    }
////    @Override
////    public int populate(XSSFSheet sheet, int rowNum) {
////        addHeader(sheet, rowNum);
////        rowNum++;
////
////        Row headerRow = sheet.createRow(rowNum++);
////        headerRow.createCell(0).setCellValue("School Name");
////        headerRow.createCell(1).setCellValue("Major");
////        headerRow.createCell(2).setCellValue("Grade");
////        headerRow.createCell(3).setCellValue("Start Date");
////        headerRow.createCell(4).setCellValue("End Date");
////
////        for (Education edu : educations) {
////            Row row = sheet.createRow(rowNum++);
////            row.createCell(0).setCellValue(edu.getSchoolName());
////            row.createCell(1).setCellValue(edu.getMajor());
////            row.createCell(2).setCellValue(edu.getGrade());
////            row.createCell(3).setCellValue(edu.getStartDate());
////            row.createCell(4).setCellValue(edu.getEndDate());
////        }
////        return rowNum;
////    }
//}
