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
            imageDataSection.setImageType("png");

            // create each chart section
            RadarChartSection radarChartSection = new RadarChartSection("Skill Radar Chart");
            PieChartSection pieChartSection = new PieChartSection("Skill Pie Chart");
            BarChartSection barChartSection = new BarChartSection("Skill Bar Chart");
            LineChartSection lineChartSection = new LineChartSection("Skill Line Chart");

            // set each chart section's data section reference
//            radarChartSection.setDataSource(skillDataSection);
//            pieChartSection.setDataSource(skillDataSection);
//            barChartSection.setDataSource(skillDataSection);
//            lineChartSection.setDataSource(skillDataSection);

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
            sheet.addSection("A1", personalInfoDataSection);
            sheet.addSection( "H60", educationDataSection);
            sheet.addSection( "A9", experienceDataSection);
            sheet.addSection( "H9", projectDataSection);
            sheet.addSection("A15", skillDataSection);
            sheet.addSection("Z50", imageDataSection);

            sheet.addSection("B30", radarChartSection, "Skill");
            sheet.addSection("B50", pieChartSection, "Skill");
            sheet.addSection("B70", barChartSection, "Skill");
            sheet.addSection("B90", lineChartSection, "Skill");

            sheet.render();

            // Hide Column
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

        // add sheet protection
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