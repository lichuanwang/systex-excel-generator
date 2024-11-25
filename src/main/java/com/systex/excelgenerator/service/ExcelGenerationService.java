package com.systex.excelgenerator.service;

import com.systex.excelgenerator.component.*;
import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.excel.ExcelFile;
import com.systex.excelgenerator.model.Candidate;
import com.systex.excelgenerator.model.Education;
import com.systex.excelgenerator.utils.ExcelStyleAndSheetHandler;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.IOException;
import java.util.*;

public class ExcelGenerationService {

    public void generateExcel(String[] headerColValues, Map<Integer, List<Object>> dataMap) {
        EducationDataSection educationDataSection = new EducationDataSection();
        educationDataSection.setData(headerColValues, dataMap);

        ExcelFile excelFile = new ExcelFile("Candidate Information");
        ExcelSheet sheet = excelFile.createSheet("Test");
        sheet.addSection(educationDataSection, "A1");
        // 設定section中的data往右長或是往下長(header是一行還是一列) -> 使用者傳入參數要往右還是往下

        try {
            excelFile.save("candidate_info_test.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

//    public void generateExcelForCandidate(List<Candidate> candidates) {
//
//        // Create a new file
//        ExcelFile excelFile = new ExcelFile("Candidate Information");
//
//        for (Candidate candidate : candidates) {
//            // create a new sheet
//            ExcelSheet sheet = excelFile.createSheet(candidate.getName());
////            PersonalInfoDataSection personalInfoDataSection = new PersonalInfoDataSection();
////            personalInfoDataSection.setData(List.of(candidate));
//
//            EducationDataSection educationDataSection = new EducationDataSection();
//
//            List<Object> headerColumnValue = new ArrayList<>();
//            headerColumnValue.add("School Name");
//            headerColumnValue.add("Major");
//            Map<Integer, List<Object>> educationDataMap = new LinkedHashMap<>();
//            educationDataMap.put(0, headerColumnValue);
//            List<Education> eduList = candidate.getEducationList();
//            for (int i = 0; i < eduList.size(); i++) {
//                List<Object> eduData = new ArrayList<>();
//                Education edu = eduList.get(i);
//                eduData.add(edu.getSchoolName());
//                eduData.add(edu.getMajor());
//                eduData.add(edu.getGrade());
//                eduData.add(edu.getStartDate());
//                eduData.add(edu.getEndDate());
//                educationDataMap.put(i+1, eduData);
//            }
//            educationDataSection.setData(educationDataMap);
//            //        headerRow.createCell(startCol).setCellValue("School Name");
////        headerRow.createCell(startCol + 1).setCellValue("Major");
////        headerRow.createCell(startCol + 2).setCellValue("Grade");
////        headerRow.createCell(startCol + 3).setCellValue("Start Date");
////        headerRow.createCell(startCol + 4).setCellValue("End Date");
////        headerRow.createCell(startCol + 5).setCellValue("Date Interval");
//
//
////            ExperienceDataSection experienceDataSection = new ExperienceDataSection();
////            experienceDataSection.setData(candidate.getExperienceList());
////
////            ProjectDataSection projectDataSection = new ProjectDataSection();
////            projectDataSection.setData(candidate.getProjects());
////
////            SkillDataSection skillDataSection = new SkillDataSection();
////            skillDataSection.setData(candidate.getSkills());
////
////            ImageDataSection imageDataSection = new ImageDataSection();
////            imageDataSection.setData(candidate.getImagepath());
//
//            // add sections to sheet
//            //sheet.addSection(personalInfoDataSection, "A1");
//            sheet.addSection(educationDataSection, "H60");
////            sheet.addSection(experienceDataSection, "A9");
////            sheet.addSection(projectDataSection, "H9");
////            sheet.addSection(skillDataSection, "A15");
//
////            // add image section to sheet
////            sheet.addSection(imageDataSection , "png" , "Z50");
//
//            // add chart sections to sheet
////            sheet.addChartSection("B30", new RadarChartSection(), "Skill", 6, 6);
////            sheet.addChartSection("B50", new PieChartSection(), "Skill", 6, 6);
////            sheet.addChartSection("B70", new BarChartSection(), "Skill",  6, 6);
////            sheet.addChartSection("B90", new LineChartSection(), "Skill", 6, 6);
//
//            // Hidden col
//            ExcelStyleAndSheetHandler.hideColumns(sheet.getXssfSheet(),false,10,12);
//
//            // Determine the maximum number of columns
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
//                int currentWidth = xssfSheet.getColumnWidth(columnIndex);
//                xssfSheet.setColumnWidth(columnIndex, (int) (currentWidth * 1.1));
//            }
//        }
//
//        // add protectSheet
//        ExcelStyleAndSheetHandler styleUtils = new ExcelStyleAndSheetHandler();
//        styleUtils.protectSheet(excelFile.getExelSheet("JohnDoe").getXssfSheet(), "12345");
//
//        // Save the Excel file
//        try {
//            excelFile.save("candidate_info_test.xlsx");
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
}