package com.systex.excelgenerator.service;

import com.systex.excelgenerator.component.*;
import com.systex.excelgenerator.excel.ExcelSheet;
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
            sheet.addImageSection(imageDataSection , "png" , "G30");

            // add chart sections to sheet
            sheet.addChartSection("A30", new RadarChartSection(), "Skill", 6, 6);
            sheet.addChartSection("A50", new PieChartSection(), "Skill", 6, 6);
            sheet.addChartSection("A70", new BarChartSection(), "Skill",  6, 6);
            sheet.addChartSection("A90", new LineChartSection(), "Skill", 6, 6);

            // Apply styles to sheet
            applyStyles(sheet);

            // Auto-size all columns up to the maximum column index
            for (int i = 0; i < 100; i++) {
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
        Workbook wb = sheet.getWorkbook();

        if (headerRow != null) {
            for (Cell cell : headerRow) {
                CellStyle style = wb.createCellStyle();
                Font font = wb.createFont();
                font.setBold(true);
                font.setFontHeightInPoints((short) 14);
                style.setFont(font);
                style.setAlignment(HorizontalAlignment.CENTER);
                style.setBorderBottom(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setBorderTop(BorderStyle.THIN);
                cell.setCellStyle(style);
            }
        }
    }
}
