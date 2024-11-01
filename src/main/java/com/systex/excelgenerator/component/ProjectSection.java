package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Project;
import com.systex.excelgenerator.utils.ExcelUtils;
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
    public int populate(XSSFSheet sheet) {
        addHeader(sheet);
        if(sheet.getPhysicalNumberOfRows() == 0){
            relativeRow = 0;
            relativeColumn = 0;
        }
        int bodyRow = relativeRow + 1;
        int flag = relativeColumn;

        String[] headers = {"Project", "Role", "Description", "Technology"};

        Row headerRow = ExcelUtils.createOrGet(sheet, bodyRow++);
        relativeColumn = flag;
        for (String header : headers) {
            headerRow.createCell(relativeColumn++).setCellValue(header);
        }

        for (Project project : projects) {
            relativeColumn = flag;
            Row row = ExcelUtils.createOrGet(sheet, bodyRow++);
            Object[] data = {
                    project.getProjectName(),
                    project.getRole(),
                    project.getDescription(),
                    project.getTechnologiesUsed()
            };
            for (Object value : data) {
                row.createCell(relativeColumn++).setCellValue(String.valueOf(value));
            }
        }

        relativeColumn += ExcelUtils.colStride(2);
        nextRelativeRow = Math.max(relativeRow, bodyRow);
        if (relativeColumn >= maxCol) {
            relativeRow = ExcelUtils.rowStride(nextRelativeRow);
            relativeColumn = 0;
        }

        System.out.println("relativeRow: " + relativeRow);
        System.out.println("relativeColumn: " + relativeColumn);
        System.out.println("nextRelativeRow: " + nextRelativeRow);
        return relativeRow;
    }

}
