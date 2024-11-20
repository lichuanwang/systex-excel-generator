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
            ExcelSheet sheet = excelFile.createSheet(candidate.getName(), 1000, 1000);

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
            // the process of creating personalInfoSection could be a static method
            sheet.addSectionAt("F40", personalInfoDataSection);
            sheet.addSectionAt("D15", educationDataSection);
            sheet.addSectionAt("A8", experienceDataSection);
//            sheet.addSectionAt("D7", projectDataSection);
            sheet.addSectionAt("Z150", skillDataSection);
//            sheet.addSectionAt("A2", personalInfoDataSection);

            // add chart sections to sheet
            sheet.addChartSection(new RadarChartSection() , "Skill", "A20");
//            sheet.addChartSection(new PieChartSection() , "Skill");
//            sheet.addChartSection(new BarChartSection() , "Skill");
//            sheet.addChartSection(new LineChartSection() , "Skill");

            // Apply styles to sheet
            applyStyles(sheet);

            // Auto-size all columns up to the maximum column index
//            for (int i = 0; i < sheet.getMaxColPerRow(); i++) {
//                XSSFSheet xssfSheet = sheet.getXssfSheet();
//                xssfSheet.autoSizeColumn(i);
//            }
        }

        // Save the Excel file
        try {
            excelFile.save("candidate_info_test_layout.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

//    public enum SectionType {
//        PERSONAL_INFO {
//            @Override
//            public AbstractDataSection<?> getInstance() {
//                return new PersonalInfoDataSection();
//            }
//        },
//        EDUCATION {
//            @Override
//            public AbstractDataSection<?> getInstance() {
//                return new EducationDataSection();
//            }
//        };
//
//        public abstract AbstractDataSection<?> getInstance();
//    }

//    public static void main(String[] args) {
//        AbstractDataSection<?> instance = SectionType.PERSONAL_INFO.getInstance();
//        instance.setData(List.of());
//    }


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
