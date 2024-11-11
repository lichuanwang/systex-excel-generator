package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Experience;
import com.systex.excelgenerator.style.CustomStyle;
import com.systex.excelgenerator.style.ExcelFormat;
import com.systex.excelgenerator.style.ExcelStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;

public class ExperienceSection extends AbstractSection<Experience> {

    private List<Experience> experiences;
    private CellStyle clonedBlueStyle;

    public ExperienceSection() {
        super("Experience");
    }

    @Override
    public void setData(Experience data) {
        if( experiences != null ) {
            this.experiences = Arrays.asList(data); // Check if this will return the same thing just like the one below
        }
    }

    @Override
    public void setData(Collection<Experience> dataCollection) {
        if (dataCollection != null && !dataCollection.isEmpty()) {
            this.experiences = new ArrayList<>(dataCollection);
        }
    }

    @Override
    public boolean isEmpty() {
        return experiences == null || experiences.isEmpty();
    }

    @Override
    public int getWidth() {
        // Set the width based on the number of columns this section uses.
        return 6; // Example width, assuming we have 5 columns for education details
    }

    @Override
    public int getHeight() {
        // Height based on the number of education entries
        return experiences.size() + 2; // +2 for the header row and one extra row space
    }

    protected void populateHeader(ExcelSheet sheet, int startRow, int startCol) {
        // Create header row for Education section
        Row headerRow = sheet.createOrGetRow(startRow);
        headerRow.createCell(startCol).setCellValue("Company");
        headerRow.createCell(startCol + 1).setCellValue("Role");
        headerRow.createCell(startCol + 2).setCellValue("Description");
        headerRow.createCell(startCol + 3).setCellValue("Start Date");
        headerRow.createCell(startCol + 4).setCellValue("End Date");
    }

    protected void populateBody(ExcelSheet sheet, int startRow, int startCol) {
        XSSFWorkbook workbook = (XSSFWorkbook) sheet.getUnderlyingSheet().getWorkbook();

        // 1. 創建並clone預設樣式
        CellStyle initialStyle = CustomStyle.createSpecialStyle(workbook);
        clonedBlueStyle  = ExcelStyle.cloneStyle(workbook, initialStyle);

        // 2. 修改clone樣式的背景色為寶石藍
        clonedBlueStyle.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
        clonedBlueStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle dateStyle = ExcelFormat.DateFormatting(workbook);

        int rowNum = startRow; // Start from the row after the header

        for (Experience exp : experiences) {
            Row row = sheet.createOrGetRow(rowNum++);
            Cell jobTitleCell = row.createCell(startCol + 1);
            jobTitleCell.setCellValue(exp.getJobTitle());
            jobTitleCell.setCellStyle(clonedBlueStyle );        // 設置儲存格樣式

            row.createCell(startCol + 2).setCellValue(exp.getDescription());
            row.createCell(startCol + 3).setCellValue(exp.getStartDate());
            row.createCell(startCol + 4).setCellValue(exp.getEndDate());
            Cell dateCell =  row.createCell(startCol + 3);
            dateCell.setCellValue(exp.getStartDate());
            dateCell.setCellStyle(dateStyle);
            dateCell =  row.createCell(startCol + 4);
            dateCell.setCellValue(exp.getEndDate());
            dateCell.setCellStyle(dateStyle);
        }
    }

    protected void populateFooter(ExcelSheet sheet, int startRow, int startCol) {

    }
}

//package com.systex.excelgenerator.component;
//
//import com.systex.excelgenerator.model.Experience;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//
//import java.util.ArrayList;
//import java.util.Arrays;
//import java.util.Collection;
//import java.util.List;
//
//public class ExperienceSection extends AbstractSection<Experience> {
//
//    private List<Experience> experiences;
//
//    public ExperienceSection() {
//        super("Experience");
//    }
//
//    @Override
//    protected int generateHeader(XSSFSheet sheet, int rowNum) {
//        Row headerRow = sheet.createRow(rowNum++);
//        headerRow.createCell(0).setCellValue("Company");
//        headerRow.createCell(1).setCellValue("Role");
//        headerRow.createCell(2).setCellValue("Description");
//        headerRow.createCell(3).setCellValue("Start Date");
//        headerRow.createCell(4).setCellValue("End Date");
//
//        return rowNum;
//    }
//
//    @Override
//    protected int generateData(XSSFSheet sheet, int rowNum) {
//        for (Experience exp : experiences) {
//            Row row = sheet.createRow(rowNum++);
//            row.createCell(0).setCellValue(exp.getCompanyName());
//            row.createCell(1).setCellValue(exp.getJobTitle());
//            row.createCell(2).setCellValue(exp.getDescription());
//            row.createCell(3).setCellValue(exp.getStartDate());
//            row.createCell(4).setCellValue(exp.getEndDate());
//        }
//        return rowNum;
//    }
//
//    @Override
//    protected int generateFooter(XSSFSheet sheet, int rowNum) {
//        return rowNum;
//    }
//
//    @Override
//    public void setData(Experience data) {
//        if (experiences != null) {
//            this.experiences = Arrays.asList(data);
//        }
//    }
//
//    @Override
//    public void setData(Collection<Experience> dataCollection) {
//        if (dataCollection != null && !dataCollection.isEmpty()) {
//            this.experiences = new ArrayList<>(dataCollection);
//        }
//    }
//
//    @Override
//    public boolean isEmpty() {
//        return experiences == null || experiences.isEmpty();
//    }
//
//
////    @Override
////    public int populate(XSSFSheet sheet, int rowNum) {
////        addHeader(sheet, rowNum);
////        rowNum++;
////
////        Row headerRow = sheet.createRow(rowNum++);
////        headerRow.createCell(0).setCellValue("Company");
////        headerRow.createCell(1).setCellValue("Role");
////        headerRow.createCell(2).setCellValue("Description");
////        headerRow.createCell(3).setCellValue("Start Date");
////        headerRow.createCell(4).setCellValue("End Date");
////
////        for (Experience exp : experiences) {
////            Row row = sheet.createRow(rowNum++);
////            row.createCell(0).setCellValue(exp.getCompanyName());
////            row.createCell(1).setCellValue(exp.getJobTitle());
////            row.createCell(2).setCellValue(exp.getDescription());
////            row.createCell(3).setCellValue(exp.getStartDate());
////            row.createCell(4).setCellValue(exp.getEndDate());
////        }
////
////        return rowNum;
////    }
//}
