package com.systex.excelgenerator.service;

import com.systex.excelgenerator.component.*;
import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.excel.ExcelFile;
import com.systex.excelgenerator.model.Candidate;
import com.systex.excelgenerator.utils.ExcelStyleAndSheetUtils;
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

            // Start creating each section
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
            imageDataSection.setImageType("png");

            RadarChartSection radarChartSection = new RadarChartSection();
            radarChartSection.setHeight(6);
            radarChartSection.setWidth(6);

            PieChartSection pieChartSection = new PieChartSection();
            pieChartSection.setHeight(6);
            pieChartSection.setWidth(6);

            BarChartSection barChartSection = new BarChartSection();
            barChartSection.setHeight(6);
            barChartSection.setWidth(6);

            LineChartSection lineChartSection = new LineChartSection();
            lineChartSection.setHeight(6);
            lineChartSection.setWidth(6);

            // add sections to sheet
            sheet.addSection(personalInfoDataSection, "A1");
            sheet.addSection(educationDataSection, "H60");
            sheet.addSection(experienceDataSection, "A9");
            sheet.addSection(projectDataSection, "H9");
            sheet.addSection(skillDataSection, "A15");
            sheet.addSection(imageDataSection, "Z50");
            sheet.addChartSection("B30", radarChartSection, "Skill");
            sheet.addChartSection("B50", pieChartSection, "Skill");
            sheet.addChartSection("B70", barChartSection, "Skill");
            sheet.addChartSection("B90", lineChartSection, "Skill");

            // Hide Column
            ExcelStyleAndSheetUtils.hideColumns(sheet.getXssfSheet(),10,12);

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
                int currentWidth = xssfSheet.getColumnWidth(columnIndex);
                xssfSheet.setColumnWidth(columnIndex, (int) (currentWidth * 1.1));
            }
        }

        // add protectSheet
        ExcelStyleAndSheetUtils.protectSheet(excelFile.getExelSheet("JohnDoe").getXssfSheet(), "12345");

        // Save the Excel file
        try {
            excelFile.save("candidate_info_test.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}