package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Education;
import com.systex.excelgenerator.utils.FormattingHandler;
import com.systex.excelgenerator.utils.FormulaHandler;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

public class EducationSection extends Section {

    private List<Education> educations;
    private FormattingHandler formattingHandler;
    private FormulaHandler formulaHandler;

    {
        this.formattingHandler = new FormattingHandler();
        this.formulaHandler = new FormulaHandler();
    }

    public EducationSection(List<Education> educations) {
        super("Education");
        this.educations = educations;
    }

    @Override
    public int populate(XSSFSheet sheet, int rowNum) {
        addHeader(sheet, rowNum);
        rowNum++;

        Row headerRow = sheet.createRow(rowNum++);
        headerRow.createCell(0).setCellValue("School Name");
        headerRow.createCell(1).setCellValue("Major");
        headerRow.createCell(2).setCellValue("Grade");
        headerRow.createCell(3).setCellValue("Start Date");
        headerRow.createCell(4).setCellValue("End Date");
        headerRow.createCell(5).setCellValue("DateInterval");

        for (Education edu : educations) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(edu.getSchoolName());
            row.createCell(1).setCellValue(edu.getMajor());
            row.createCell(2).setCellValue(edu.getGrade());

            //  format start and end date
            Cell dateCell = row.createCell(3);
            dateCell.setCellValue(edu.getStartDate());
            dateCell.setCellStyle(formattingHandler.DateFormatting(edu.getStartDate(), sheet.getWorkbook()));

            dateCell = row.createCell(4);
            dateCell.setCellValue(edu.getEndDate());
            dateCell.setCellStyle(formattingHandler.DateFormatting(edu.getEndDate(), sheet.getWorkbook()));

            // cal date interval
            dateCell = row.createCell(5);
            dateCell.setCellFormula(formulaHandler.calDataInterval(row.getRowNum() , 3 , 4));

        }
        return rowNum;
    }
}
