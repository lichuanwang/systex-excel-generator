package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Skill;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

public class SkillSection extends Section {

    private List<Skill> skills;

    public SkillSection(List<Skill> skills) {
        super("Skill");
        this.skills = skills;
    }

    @Override
    public int populate(XSSFSheet sheet, int rowNum) {
        addHeader(sheet, rowNum);
        rowNum++;

        Row headerRow = sheet.createRow(rowNum++);
        headerRow.createCell(0).setCellValue("Id");
        headerRow.createCell(1).setCellValue("Name");
        headerRow.createCell(2).setCellValue("Level");

        for (Skill skill : skills) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(skill.getId());
            row.createCell(1).setCellValue(skill.getSkillName());
            row.createCell(2).setCellValue(skill.getLevel());
        }
        return rowNum;

    }
}
