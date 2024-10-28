package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Experience;
import com.systex.excelgenerator.utils.FormattingHandler;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

public class ExperienceSection extends Section {

    private List<Experience> experiences;
    private FormattingHandler formattingHandler;

    {
        this.formattingHandler = new FormattingHandler();
    }

    public ExperienceSection(List<Experience> experiences) {
        super("Experience");
        this.experiences = experiences;
    }

    @Override
    public int populate(XSSFSheet sheet, int rowNum) {
        addHeader(sheet, rowNum);
        rowNum++;

        Row headerRow = sheet.createRow(rowNum++);
        headerRow.createCell(0).setCellValue("Company");
        headerRow.createCell(1).setCellValue("Role");
        headerRow.createCell(2).setCellValue("Description");
        headerRow.createCell(3).setCellValue("Start Date");
        headerRow.createCell(4).setCellValue("End Date");

        for (Experience exp : experiences) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(exp.getCompanyName());
            row.createCell(1).setCellValue(exp.getJobTitle());
            row.createCell(2).setCellValue(exp.getDescription());

            //  format start and end date
            Cell dateCell = row.createCell(3);
            dateCell.setCellValue(exp.getStartDate());
            dateCell.setCellStyle(formattingHandler.DateFormatting(exp.getStartDate(), sheet.getWorkbook()));

            dateCell = row.createCell(4);
            dateCell.setCellValue(exp.getEndDate());
            dateCell.setCellStyle(formattingHandler.DateFormatting(exp.getEndDate(), sheet.getWorkbook()));
        }

        return rowNum;
    }
}
