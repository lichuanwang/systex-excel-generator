package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Project;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

public class ProjectSection extends Section {

    private List<Project> projects;

    public ProjectSection(List<Project> projects) {
        super("Project");
        this.projects = projects;
    }

    @Override
    public int populate(XSSFSheet sheet, int rowNum) {
        addHeader(sheet, rowNum);
        rowNum++;

        Row headerRow = sheet.createRow(rowNum++);
        headerRow.createCell(0).setCellValue("Project");
        headerRow.createCell(1).setCellValue("Role");
        headerRow.createCell(2).setCellValue("Description");
        headerRow.createCell(3).setCellValue("Technology");

        for (Project project : projects) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(project.getProjectName());
            row.createCell(1).setCellValue(project.getRole());
            row.createCell(2).setCellValue(project.getDescription());
            row.createCell(3).setCellValue(project.getTechnologiesUsed());
        }
        return rowNum;

    }
}
