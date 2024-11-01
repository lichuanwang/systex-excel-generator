package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Skill;
import com.systex.excelgenerator.utils.ExcelUtils;
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
    public int populate(XSSFSheet sheet) {
        addHeader(sheet);
        if(sheet.getPhysicalNumberOfRows() == 0){
            relativeRow = 0;
            relativeColumn = 0;
        }
        int bodyRow = relativeRow + 1;
        int flag = relativeColumn;

        String[] headers = {"Id", "Name", "Level"};

        Row headerRow = ExcelUtils.createOrGet(sheet, bodyRow++);
        for (String header : headers) {
            headerRow.createCell(relativeColumn++).setCellValue(header);
        }

        for (Skill skill : skills) {
            relativeColumn = flag;
            Row row = ExcelUtils.createOrGet(sheet, bodyRow++);
            Object[] data = {
                    skill.getId(),
                    skill.getSkillName(),
                    skill.getLevel()
            };
            for (Object value : data) {
                row.createCell(relativeColumn++).setCellValue(String.valueOf(value));
            }
        }

        relativeColumn += ExcelUtils.colStride(2);
        nextRelativeRow = Math.max(relativeRow, bodyRow);

        if (relativeColumn >= maxCol) {
            relativeRow = ExcelUtils.rowStride(nextRelativeRow);
            relativeColumn = 0;
        }

        System.out.println("relativeRow: " + relativeRow);
        System.out.println("relativeColumn: " + relativeColumn);
        System.out.println("nextRelativeRow: " + nextRelativeRow);
        System.out.println("row number: " + sheet.getPhysicalNumberOfRows());
        return relativeRow;
    }
}
