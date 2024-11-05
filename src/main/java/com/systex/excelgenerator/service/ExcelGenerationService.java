package com.systex.excelgenerator.service;

import com.systex.excelgenerator.builder.RecruitmentExcelBuilder;
import com.systex.excelgenerator.builder.ExcelBuilder;
import com.systex.excelgenerator.style.StyleBuilder;
import com.systex.excelgenerator.director.ExcelDirector;
import com.systex.excelgenerator.excel.ExcelFile;
import com.systex.excelgenerator.model.Candidate;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.IOException;

public class ExcelGenerationService {

    public void generateExcelForCandidate(Candidate candidate) {
        // Build the Excel content
        ExcelBuilder builder = new RecruitmentExcelBuilder(candidate);
        ExcelDirector director = new ExcelDirector(builder);
        director.constructExcelFile();

        ExcelFile excelFile = director.getExcelFile();

        // Apply custom styles to the content
        XSSFSheet sheet = excelFile.getWorkbook().getSheet(candidate.getName());
        applyStyles(sheet);

        int maxColumnIndex = 0;
        for (int rowIndex = 0; rowIndex <= sheet.getPhysicalNumberOfRows(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null && row.getLastCellNum() > maxColumnIndex) {
                maxColumnIndex = row.getLastCellNum();
            }
        }

        // Auto-size all columns up to the maximum column index
        for (int i = 0; i < maxColumnIndex; i++) {
            sheet.autoSizeColumn(i);
        }

        // Save the Excel file
        try {
            excelFile.saveToFile("candidate_info_test.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void applyStyles(XSSFSheet sheet) {
        // Example: Apply styles to the first row (header)
        Row headerRow = sheet.getRow(0);
        StyleBuilder styleBuilder = new StyleBuilder(sheet.getWorkbook());

        if (headerRow != null) {
            for (Cell cell : headerRow) {
                CellStyle headerStyle = styleBuilder.setBold(true)
                        .setFontSize((short) 14)
                        .setAlignment(HorizontalAlignment.CENTER)
                        .setBorder(BorderStyle.THIN)
                        .build();
                cell.setCellStyle(headerStyle);
            }
        }

        // Auto-size columns
        for (int i = 0; i < 5; i++) {
            sheet.autoSizeColumn(i);
        }
    }
}
