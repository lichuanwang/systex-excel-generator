package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Experience;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

public class ExperienceSection extends AbstractSection<Experience> {

    private List<Experience> experiences;

    public ExperienceSection() {
        super("Experience");
    }

    @Override
    protected int generateHeader(XSSFSheet sheet, int rowNum) {
        Row headerRow = sheet.createRow(rowNum++);
        headerRow.createCell(0).setCellValue("Company");
        headerRow.createCell(1).setCellValue("Role");
        headerRow.createCell(2).setCellValue("Description");
        headerRow.createCell(3).setCellValue("Start Date");
        headerRow.createCell(4).setCellValue("End Date");

        return rowNum;
    }

    @Override
    protected int generateData(XSSFSheet sheet, int rowNum) {
        for (Experience exp : experiences) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(exp.getCompanyName());
            row.createCell(1).setCellValue(exp.getJobTitle());
            row.createCell(2).setCellValue(exp.getDescription());
            row.createCell(3).setCellValue(exp.getStartDate());
            row.createCell(4).setCellValue(exp.getEndDate());
        }
        return rowNum;
    }

    @Override
    protected int generateFooter(XSSFSheet sheet, int rowNum) {
        return rowNum;
    }

    @Override
    public void setData(Experience data) {
        if (experiences != null) {
            this.experiences = Arrays.asList(data);
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


//    @Override
//    public int populate(XSSFSheet sheet, int rowNum) {
//        addHeader(sheet, rowNum);
//        rowNum++;
//
//        Row headerRow = sheet.createRow(rowNum++);
//        headerRow.createCell(0).setCellValue("Company");
//        headerRow.createCell(1).setCellValue("Role");
//        headerRow.createCell(2).setCellValue("Description");
//        headerRow.createCell(3).setCellValue("Start Date");
//        headerRow.createCell(4).setCellValue("End Date");
//
//        for (Experience exp : experiences) {
//            Row row = sheet.createRow(rowNum++);
//            row.createCell(0).setCellValue(exp.getCompanyName());
//            row.createCell(1).setCellValue(exp.getJobTitle());
//            row.createCell(2).setCellValue(exp.getDescription());
//            row.createCell(3).setCellValue(exp.getStartDate());
//            row.createCell(4).setCellValue(exp.getEndDate());
//        }
//
//        return rowNum;
//    }
}
