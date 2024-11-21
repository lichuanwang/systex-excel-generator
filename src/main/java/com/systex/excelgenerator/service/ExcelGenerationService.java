package com.systex.excelgenerator.service;

import com.systex.excelgenerator.component.*;
import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.excel.ExcelFile;
import com.systex.excelgenerator.model.Candidate;
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

            ImageDataSection imageDataSection = new ImageDataSection();
            imageDataSection.setData(candidate.getImagepath());

            // add sections to sheet
            sheet.addSection(personalInfoDataSection, "A1");
            sheet.addSection(educationDataSection, "H1");
            sheet.addSection(experienceDataSection, "A9");
            sheet.addSection(projectDataSection, "H9");
            sheet.addSection(skillDataSection, "A15");

            // add image section to sheet
            sheet.addSection(imageDataSection , "png" , "G30");

            // add chart sections to sheet
            sheet.addChartSection("A30", new RadarChartSection(), "Skill", 6, 6);
            sheet.addChartSection("A50", new PieChartSection(), "Skill", 6, 6);
            sheet.addChartSection("A70", new BarChartSection(), "Skill",  6, 6);
            sheet.addChartSection("A90", new LineChartSection(), "Skill", 6, 6);

            // Determine the maximum number of columns
            int maxColumns = 0;
            XSSFSheet xssfSheet = sheet.getXssfSheet();
            for (int rowIndex = 0; rowIndex <= xssfSheet.getLastRowNum(); rowIndex++) {
                XSSFRow currentRow = xssfSheet.getRow(rowIndex);
                if (currentRow != null && currentRow.getLastCellNum() > maxColumns) {
                    maxColumns = currentRow.getLastCellNum();
                }
            }

            // Autosize all columns based on the maximum column count
            for (int columnIndex = 0; columnIndex < maxColumns; columnIndex++) {
                xssfSheet.autoSizeColumn(columnIndex);
            }
        }

        // Save the Excel file
        try {
            excelFile.save("candidate_info_test.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
