package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Project;
import com.systex.excelgenerator.utils.HyperlinkHandler;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

public class ProjectSection extends Section {

    private List<Project> projects;
    private HyperlinkHandler hyperlinkHandler;

    {
        this.hyperlinkHandler = new HyperlinkHandler();
    }

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
        headerRow.createCell(4).setCellValue("ProjectLink");

        for (Project project : projects) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(project.getProjectName());
            row.createCell(1).setCellValue(project.getRole());
            row.createCell(2).setCellValue(project.getDescription());

            // demo test : set internal link
            Cell linkCell = row.createCell(3);
            linkCell.setCellValue(project.getTechnologiesUsed());
            hyperlinkHandler.setInternalLink("Test" , linkCell , sheet.getWorkbook());
            //row.createCell(3).setCellValue(project.getTechnologiesUsed());

            // set outer link
            linkCell = row.createCell(4);
            linkCell.setCellValue("Click me!");
            hyperlinkHandler.setHyperLink(project.getProjectlink(), linkCell , sheet.getWorkbook());
        }
        return rowNum;

    }
}
