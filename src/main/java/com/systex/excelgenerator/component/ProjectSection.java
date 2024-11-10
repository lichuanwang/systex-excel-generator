package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Project;
import org.apache.poi.ss.usermodel.Row;

public class ProjectSection extends AbstractSection<Project> {

    public ProjectSection() {
        super("Project");
    }

    @Override
    public boolean isEmpty() {
        return content == null || content.isEmpty();
    }

    @Override
    public int getWidth() {
        // Set the width based on the number of columns this section uses.
        return 5; // Example width, assuming we have 5 columns for project details
    }

    @Override
    public int getHeight() {
        // Height based on the number of education entries
        return content.size() + 1; // +1 for the header row
    }

    protected void renderHeader(ExcelSheet sheet, int startRow, int startCol) {
        // Create header row for Education section
        Row headerRow = sheet.createOrGetRow(startRow);
        headerRow.createCell(startCol).setCellValue("Project Name");
        headerRow.createCell(startCol + 1).setCellValue("Role");
        headerRow.createCell(startCol + 2).setCellValue("Description");
        headerRow.createCell(startCol + 3).setCellValue("Technologies Used");
    }

    protected void renderBody(ExcelSheet sheet, int startRow, int startCol) {
        int rowNum = startRow; // Start from the row after the header

        for (Project project : content) {
            Row row = sheet.createOrGetRow(rowNum++);
            row.createCell(startCol).setCellValue(project.getProjectName());
            row.createCell(startCol + 1).setCellValue(project.getRole());
            row.createCell(startCol + 2).setCellValue(project.getDescription());
            row.createCell(startCol + 3).setCellValue(project.getTechnologiesUsed());
        }
    }

    protected void renderFooter(ExcelSheet sheet, int startRow, int startCol) {
        // implement footer logic here
    }
}