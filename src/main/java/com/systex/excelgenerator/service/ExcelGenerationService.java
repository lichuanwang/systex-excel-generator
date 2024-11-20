package com.systex.excelgenerator.service;

import com.systex.excelgenerator.component.*;
import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.excel.ExcelFile;
import com.systex.excelgenerator.model.Candidate;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.IOException;
import java.util.List;

public class ExcelGenerationService {

    public void generateExcelForCandidate(List<Candidate> candidates) {

        // Create a new file
        ExcelFile excelFile = new ExcelFile("Candidate Information");

        for (Candidate candidate : candidates) {
            // create a new sheet
            ExcelSheet sheet = excelFile.createSheet(candidate.getName());

            PersonalInfoDataSection personalInfoDataSection = new PersonalInfoDataSection();
            personalInfoDataSection.setData(List.of(candidate));

            EducationDataSection educationDataSection = new EducationDataSection();
            educationDataSection.setData(candidate.getEducationList());

            ExperienceDataSection experienceDataSection = new ExperienceDataSection();
            experienceDataSection.setData(candidate.getExperienceList());

            ProjectDataSection projectDataSection = new ProjectDataSection();
            projectDataSection.setData(candidate.getProjects());

            SkillDataSection skillDataSection = new SkillDataSection();
            skillDataSection.setData(candidate.getSkills());

            // add sections to sheet
            sheet.addSection(personalInfoDataSection, "A1");
            sheet.addSection(educationDataSection, "H1");
            sheet.addSection(experienceDataSection, "A9");
            sheet.addSection(projectDataSection, "H9");
            sheet.addSection(skillDataSection, "A15");

            // add chart sections to sheet
            sheet.addChartSection("A30", new RadarChartSection(), "Skill", 6, 6);
            sheet.addChartSection("A50", new PieChartSection(), "Skill", 6, 6);
            sheet.addChartSection("A70", new BarChartSection(), "Skill",  6, 6);
            sheet.addChartSection("A90", new LineChartSection(), "Skill", 6, 6);

            autoSizeColumns(sheet);
        }

        // Save the Excel file
        try {
            excelFile.save("candidate_info_test.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void autoSizeColumns(ExcelSheet sheet) {
        XSSFSheet underlyingSheet = sheet.getXssfSheet();

        // Find the maximum number of columns in the sheet
        int maxColumnCount = getMaxColumnCount(underlyingSheet);

        // Auto-size each column up to the maximum column count
        for (int i = 0; i < maxColumnCount; i++) {
            underlyingSheet.autoSizeColumn(i);
        }
    }

    // Dynamically find the maximum number of columns in the sheet
    private int getMaxColumnCount(XSSFSheet sheet) {
        int maxColumns = 0;
        for (Row row : sheet) { // Iterate over all rows
            int lastCellNum = row.getLastCellNum(); // Get the last cell number in the row
            if (lastCellNum > maxColumns) {
                maxColumns = lastCellNum; // Update maxColumns if this row has more cells
            }
        }

        return maxColumns;
    }
}



// Determine the maximum number of columns
//            int maxColumns = 0;
//            XSSFSheet xssfSheet = sheet.getXssfSheet();
//            for (int rowIndex = 0; rowIndex <= xssfSheet.getLastRowNum(); rowIndex++) {
//                XSSFRow currentRow = xssfSheet.getRow(rowIndex);
//                if (currentRow != null && currentRow.getLastCellNum() > maxColumns) {
//                    maxColumns = currentRow.getLastCellNum();
//                }
//            }
//
//            // Autosize all columns based on the maximum column count
//            for (int columnIndex = 0; columnIndex < maxColumns; columnIndex++) {
//                xssfSheet.autoSizeColumn(columnIndex);
//            }