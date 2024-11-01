package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Skill;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Arrays;
import java.util.Collection;
import java.util.List;

public class SkillSection extends AbstractSection<Skill> {

    private List<Skill> skills;

    public SkillSection() {
        super("Skill");
    }

    @Override
    protected int generateHeader(XSSFSheet sheet, int rowNum) {
        Row headerRow = sheet.createRow(rowNum++);
        headerRow.createCell(0).setCellValue("Id");
        headerRow.createCell(1).setCellValue("Name");
        headerRow.createCell(2).setCellValue("Level");
        return rowNum;
    }

    @Override
    protected int generateData(XSSFSheet sheet, int rowNum) {
        for (Skill skill : skills) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(skill.getId());
            row.createCell(1).setCellValue(skill.getSkillName());
            row.createCell(2).setCellValue(skill.getLevel());
        }
        return rowNum;
    }

    @Override
    protected int generateFooter(XSSFSheet sheet, int rowNum) {
        return rowNum;
    }

    @Override
    public void setData(Skill data) {
        this.skills = Arrays.asList(data);
    }

    @Override
    public void setData(Collection<Skill> dataCollection) {
        this.skills = (List<Skill>) dataCollection;
    }

    @Override
    public boolean isEmpty() {
        return false;
    }
//
//    @Override
//    public int populate(XSSFSheet sheet, int rowNum) {
//        addHeader(sheet, rowNum);
//        rowNum++;
//
//        Row headerRow = sheet.createRow(rowNum++);
//        headerRow.createCell(0).setCellValue("Id");
//        headerRow.createCell(1).setCellValue("Name");
//        headerRow.createCell(2).setCellValue("Level");
//
//        for (Skill skill : skills) {
//            Row row = sheet.createRow(rowNum++);
//            row.createCell(0).setCellValue(skill.getId());
//            row.createCell(1).setCellValue(skill.getSkillName());
//            row.createCell(2).setCellValue(skill.getLevel());
//        }
//        return rowNum;
//
//    }
}
