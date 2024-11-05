package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Skill;
import com.systex.excelgenerator.utils.ChartHandler;
import com.systex.excelgenerator.utils.FormattingHandler;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

public class SkillSection extends Section {

    private List<Skill> skills;
    private FormattingHandler formattingHandler;
    private ChartHandler chartHandler;

    {
        this.formattingHandler = new FormattingHandler();
        this.chartHandler = new ChartHandler();
    }

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
        headerRow.createCell(1).setCellValue("Skill Name");
        headerRow.createCell(2).setCellValue("Skill Level");

        // skill開始的row (條件判斷式需要)
        int startRow = rowNum;

        for (Skill skill : skills) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(skill.getId());
            row.createCell(1).setCellValue(skill.getSkillName());
            row.createCell(2).setCellValue(skill.getLevel());
        }

        // if skill level > 2 (conditional test)
        formattingHandler.ConditionalFormatting(sheet , "2" , startRow , rowNum-1 , 2);

        // generate pie chart
        //chartHandler.genPieChart(sheet, startRow, rowNum - 1, 1, 2, rowNum+2  ,headerRow.getRowNum());

        // generate radar chart
        //chartHandler.genRadarChart(sheet, startRow, rowNum - 1, 1, 2, rowNum+2  ,headerRow.getRowNum());

        // generate bar chart
        //System.out.println(startRow+","+(rowNum-1)+",1,2,"+(rowNum+2)+","+headerRow.getRowNum());
        //chartHandler.genBarChart(sheet, startRow, rowNum - 1, 1, 2, rowNum+2  ,headerRow.getRowNum());

        // generate line chart
        chartHandler.genLineChart(sheet, startRow, rowNum - 1, 1, 2, rowNum+2  ,headerRow.getRowNum());

        return rowNum;

    }
}
