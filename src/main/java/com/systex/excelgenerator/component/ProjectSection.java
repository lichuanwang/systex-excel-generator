package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Project;
import com.systex.excelgenerator.style.ExcelStyleUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;

public class ProjectSection extends AbstractSection<Project> {

    private List<Project> projects;

    public ProjectSection() {
        super("Project");
    }

    @Override
    public void setData(Project data) {
        if( projects != null ) {
            this.projects = Arrays.asList(data); // Check if this will return the same thing just like the one below
        }
    }

    @Override
    public void setData(Collection<Project> dataCollection) {
        if (dataCollection != null && !dataCollection.isEmpty()) {
            this.projects = new ArrayList<>(dataCollection);
        }
    }

    @Override
    public boolean isEmpty() {
        return projects == null || projects.isEmpty();
    }

    @Override
    public int getWidth() {
        // Set the width based on the number of columns this section uses.
        return 5; // Example width, assuming we have 5 columns for project details
    }

    @Override
    public int getHeight() {
        // Height based on the number of education entries
        return projects.size() + 2; // +2 for the header row and extra row space
    }

    protected void populateHeader(ExcelSheet sheet, int startRow, int startCol) {
        // Create header row for Education section
        Row headerRow = sheet.createOrGetRow(startRow);
        headerRow.createCell(startCol).setCellValue("Project Name");
        headerRow.createCell(startCol + 1).setCellValue("Role");
        headerRow.createCell(startCol + 2).setCellValue("Description");
        headerRow.createCell(startCol + 3).setCellValue("Technologies Used");
    }

    protected void populateBody(ExcelSheet sheet, int startRow, int startCol) {
        XSSFWorkbook workbook = (XSSFWorkbook) sheet.getUnderlyingSheet().getWorkbook();
        CellStyle initialStyle = ExcelStyleUtils.createSpecialStyle(workbook);
        CellStyle royalBlueStyle = ExcelStyleUtils.cloneStyle(workbook, initialStyle);

        Font modifiedFont = workbook.createFont();
        modifiedFont.setFontHeightInPoints((short) 14);
        royalBlueStyle.setFont(modifiedFont);

        royalBlueStyle.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());
        royalBlueStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        int rowNum = startRow; // Start from the row after the header

        for (Project project : projects) {
            Row row = sheet.createOrGetRow(rowNum++);
            Cell projectNameCell = row.createCell(startCol);
            projectNameCell.setCellValue(project.getProjectName());
            projectNameCell.setCellStyle(royalBlueStyle);

            row.createCell(startCol + 1).setCellValue(project.getRole());
            row.createCell(startCol + 2).setCellValue(project.getDescription());
            row.createCell(startCol + 3).setCellValue(project.getTechnologiesUsed());
        }
    }

    protected void populateFooter(ExcelSheet sheet, int startRow, int startCol) {

    }
}




//package com.systex.excelgenerator.component;
//
//import com.systex.excelgenerator.model.Project;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//
//import java.util.ArrayList;
//import java.util.Arrays;
//import java.util.Collection;
//import java.util.List;
//
//public class ProjectSection extends AbstractSection<Project> {
//
//    private List<Project> projects;
//
//    public ProjectSection() {
//        super("Project");
//    }
//
//    @Override
//    protected int generateHeader(XSSFSheet sheet, int rowNum) {
//        Row headerRow = sheet.createRow(rowNum++);
//        headerRow.createCell(0).setCellValue("Project");
//        headerRow.createCell(1).setCellValue("Role");
//        headerRow.createCell(2).setCellValue("Description");
//        headerRow.createCell(3).setCellValue("Technology");
//        return rowNum;
//    }
//
//    @Override
//    protected int generateData(XSSFSheet sheet, int rowNum) {
//        for (Project project : projects) {
//            Row row = sheet.createRow(rowNum++);
//            row.createCell(0).setCellValue(project.getProjectName());
//            row.createCell(1).setCellValue(project.getRole());
//            row.createCell(2).setCellValue(project.getDescription());
//            row.createCell(3).setCellValue(project.getTechnologiesUsed());
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
//    public void setData(Project data) {
//        this.projects = Arrays.asList(data);
//    }
//
//    @Override
//    public void setData(Collection<Project> dataCollection) {
//        if (dataCollection != null && !dataCollection.isEmpty()) {
//            this.projects = new ArrayList<>(dataCollection);
//        }
//    }
//
//    @Override
//    public boolean isEmpty() {
//         return projects == null || projects.isEmpty();
//    }
//}
