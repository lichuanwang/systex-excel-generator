package com.systex.excelgenerator.service;

import com.systex.excelgenerator.component.*;
import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.excel.ExcelFile;
import com.systex.excelgenerator.model.Candidate;
import com.systex.excelgenerator.utils.ExcelStyleAndSheetHandler;
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

            // create all data section that we want to add the sheet
            PersonalInfoDataSection personalInfoDataSection = new PersonalInfoDataSection();
            EducationDataSection educationDataSection = new EducationDataSection();
            ExperienceDataSection experienceDataSection = new ExperienceDataSection();
            ProjectDataSection projectDataSection = new ProjectDataSection();
            SkillDataSection skillDataSection = new SkillDataSection();
            ImageDataSection imageDataSection = new ImageDataSection();

            // provide data for each data section
            personalInfoDataSection.setData(List.of(candidate));
            educationDataSection.setData(candidate.getEducationList());
            experienceDataSection.setData(candidate.getExperienceList());
            projectDataSection.setData(candidate.getProjects());
            skillDataSection.setData(candidate.getSkills());
            imageDataSection.setData(candidate.getImagepath());

            // create each chart section
            RadarChartSection radarChartSection = new RadarChartSection();
            PieChartSection pieChartSection = new PieChartSection();
            BarChartSection barChartSection = new BarChartSection();
            LineChartSection lineChartSection = new LineChartSection();

            // set each chart section's data section reference
            radarChartSection.setDataSource(skillDataSection);
            pieChartSection.setDataSource(skillDataSection);
            barChartSection.setDataSource(skillDataSection);
            lineChartSection.setDataSource(skillDataSection);

            // set height and width for each chart section
            radarChartSection.setHeight(6);
            radarChartSection.setWidth(6);
            pieChartSection.setHeight(6);
            pieChartSection.setWidth(6);
            barChartSection.setHeight(6);
            barChartSection.setWidth(6);
            lineChartSection.setHeight(6);
            lineChartSection.setWidth(6);

            // add chart sections to sheet
            sheet.addSection(personalInfoDataSection, "A1");
            sheet.addSection(educationDataSection, "H60");
            sheet.addSection(experienceDataSection, "A9");
            sheet.addSection(projectDataSection, "H9");
            sheet.addSection(skillDataSection, "A15");
            sheet.addSection(imageDataSection , "png" , "Z50");
            sheet.addChartSection("B30", radarChartSection);
            sheet.addChartSection("B50", pieChartSection);
            sheet.addChartSection("B70", barChartSection);
            sheet.addChartSection("B90", lineChartSection);

            // Hidden column
            ExcelStyleAndSheetHandler.hideColumns(sheet.getXssfSheet(),false,10,12);

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
        ExcelStyleAndSheetHandler styleUtils = new ExcelStyleAndSheetHandler();
        styleUtils.protectSheet(excelFile.getExelSheet("JohnDoe").getXssfSheet(), "12345");

        // Save the Excel file
        try {
            excelFile.save("candidate_info_test.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}