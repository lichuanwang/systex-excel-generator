package com.systex.excelgenerator.service;

import com.systex.excelgenerator.component.*;
import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Skill;
import com.systex.excelgenerator.style.StyleBuilder;
import com.systex.excelgenerator.excel.ExcelFile;
import com.systex.excelgenerator.model.Candidate;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.IOException;
import java.util.List;

public class ExcelGenerationService {

    public void generateExcelForCandidate(List<Candidate> candidates) {

        // Create a new file
        ExcelFile excelFile = new ExcelFile("Candidate Information");

        for (Candidate candidate : candidates) {
            // create a new sheet
            ExcelSheet sheet = excelFile.createSheet(candidate.getName(), 10);

            // add sections to sheet
            sheet.addSection(new PersonalInfoSection(), List.of(candidate));
            sheet.addSection(new EducationSection(), candidate.getEducationList());
            sheet.addSection(new ExperienceSection(), candidate.getExperienceList());
            sheet.addSection(new ProjectSection(), candidate.getProjects());
            sheet.addSection(new SkillSection(), candidate.getSkills());

            // add chart sections to sheet
            // 改成傳section name進去,在裡面用name找
            sheet.addChartSection(new RadarChartSection() , "Skill");


            // Apply styles to sheet
            applyStyles(sheet);

            // Auto-size all columns up to the maximum column index
            for (int i = 0; i < sheet.getMaxColPerRow(); i++) {
                XSSFSheet xssfSheet = sheet.getXssfSheet();
                xssfSheet.autoSizeColumn(i);
            }
        }

        // Save the Excel file
        try {
            excelFile.save("candidate_info_test.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void applyStyles(ExcelSheet sheet) {

        // get the xssfsheet
        XSSFSheet xssfSheet = sheet.getXssfSheet();

        Row headerRow = xssfSheet.getRow(0);
        StyleBuilder styleBuilder = new StyleBuilder(xssfSheet.getWorkbook());

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
    }
}
