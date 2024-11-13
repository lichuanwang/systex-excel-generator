package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Skill;
import org.apache.poi.ss.usermodel.Row;

import java.util.*;

public class SkillDataSection extends AbstractDataSection<Skill> {

    private List<Skill> skills;

    public SkillDataSection() {
        super("Skill");
    }

    @Override
    public void setData(Skill data) {
        if( skills != null ) {
            this.skills = Arrays.asList(data); // Check if this will return the same thing just like the one below
        }
    }

    @Override
    public void setData(Collection<Skill> dataCollection) {
        if (dataCollection != null && !dataCollection.isEmpty()) {
            this.skills = new ArrayList<>(dataCollection);
        }
    }

    @Override
    public boolean isEmpty() {
        return skills == null || skills.isEmpty();
    }

    @Override
    public int getWidth() {
        // Set the width based on the number of columns this section uses.
        return 4; // Example width, assuming we have 4 columns for skill details
    }

    @Override
    public int getHeight() {
        // Height based on the number of education entries
        return skills.size() + 2; // +2 for the header row and extra row space
    }

    protected void populateHeader(ExcelSheet sheet, int startRow, int startCol) {
        // Create header row for Education section
        Row headerRow = sheet.createOrGetRow(startRow);
        headerRow.createCell(startCol).setCellValue("Id");
        headerRow.createCell(startCol + 1).setCellValue("Name");
        headerRow.createCell(startCol + 2).setCellValue("Level");
    }

    protected void populateBody(ExcelSheet sheet, int startRow, int startCol) {
        int rowNum = startRow; // Start from the row after the header

        for (Skill skill : skills) {
            Row row = sheet.createOrGetRow(rowNum++);
            row.createCell(startCol).setCellValue(skill.getId());
            row.createCell(startCol + 1).setCellValue(skill.getSkillName());
            row.createCell(startCol + 2).setCellValue(skill.getLevel());
        }
    }

    protected void populateFooter(ExcelSheet sheet, int startRow, int startCol) {

    }
}




//package com.systex.excelgenerator.component;
//
//import com.systex.excelgenerator.model.Skill;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//
//import java.util.Arrays;
//import java.util.Collection;
//import java.util.List;
//
//public class SkillSection extends AbstractSection<Skill> {
//
//    private List<Skill> skills;
//
//    public SkillSection() {
//        super("Skill");
//    }
//
//    @Override
//    protected int generateHeader(XSSFSheet sheet, int rowNum) {
//        Row headerRow = sheet.createRow(rowNum++);
//        headerRow.createCell(0).setCellValue("Id");
//        headerRow.createCell(1).setCellValue("Name");
//        headerRow.createCell(2).setCellValue("Level");
//        return rowNum;
//    }
//
//    @Override
//    protected int generateData(XSSFSheet sheet, int rowNum) {
//        for (Skill skill : skills) {
//            Row row = sheet.createRow(rowNum++);
//            row.createCell(0).setCellValue(skill.getId());
//            row.createCell(1).setCellValue(skill.getSkillName());
//            row.createCell(2).setCellValue(skill.getLevel());
//        }
//        return rowNum;
//    }
//
//    @Override
//    protected int generateFooter(XSSFSheet sheet, int rowNum) {
//        return rowNum;
//    }
//
//    @Override
//    public void setData(Skill data) {
//        this.skills = Arrays.asList(data);
//    }
//
//    @Override
//    public void setData(Collection<Skill> dataCollection) {
//        this.skills = (List<Skill>) dataCollection;
//    }
//
//    @Override
//    public boolean isEmpty() {
//        return false;
//    }
////
////    @Override
////    public int populate(XSSFSheet sheet, int rowNum) {
////        addHeader(sheet, rowNum);
////        rowNum++;
////
////        Row headerRow = sheet.createRow(rowNum++);
////        headerRow.createCell(0).setCellValue("Id");
////        headerRow.createCell(1).setCellValue("Name");
////        headerRow.createCell(2).setCellValue("Level");
////
////        for (Skill skill : skills) {
////            Row row = sheet.createRow(rowNum++);
////            row.createCell(0).setCellValue(skill.getId());
////            row.createCell(1).setCellValue(skill.getSkillName());
////            row.createCell(2).setCellValue(skill.getLevel());
////        }
////        return rowNum;
////
////    }
//}
