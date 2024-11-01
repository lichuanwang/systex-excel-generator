package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Project;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

public class ProjectSection extends AbstractSection<Project> {

    private List<Project> projects;

    public ProjectSection() {
        super("Project");
    }

    @Override
    protected int generateHeader(XSSFSheet sheet, int rowNum) {
        Row headerRow = sheet.createRow(rowNum++);
        headerRow.createCell(0).setCellValue("Project");
        headerRow.createCell(1).setCellValue("Role");
        headerRow.createCell(2).setCellValue("Description");
        headerRow.createCell(3).setCellValue("Technology");
        return rowNum;
    }

    @Override
    protected int generateData(XSSFSheet sheet, int rowNum) {
        for (Project project : projects) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(project.getProjectName());
            row.createCell(1).setCellValue(project.getRole());
            row.createCell(2).setCellValue(project.getDescription());
            row.createCell(3).setCellValue(project.getTechnologiesUsed());
        }
        return rowNum;
    }

    @Override
    protected int generateFooter(XSSFSheet sheet, int rowNum) {
        return rowNum;
    }

    @Override
    public void setData(Project data) {
        this.projects = Arrays.asList(data);
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
}
