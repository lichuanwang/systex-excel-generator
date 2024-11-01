package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Education;
import com.systex.excelgenerator.model.Experience;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

public abstract class AbstractSection<T> implements Section<T> {
    protected String title;


    public AbstractSection(String title) {
        this.title = title;
    }

    // wrong naming
    protected abstract int generateHeader(XSSFSheet sheet, int rowNum);
    protected abstract int generateData(XSSFSheet sheet, int rowNum);
    protected abstract int generateFooter(XSSFSheet sheet, int rowNum);


    // problem with pass by value, should we use a rowNum or primitive type to determine in this way
    // probably using pass by reference could be better
    public int populate(XSSFSheet sheet, int rowNum) {
        addSectionTitle(sheet, rowNum);
        rowNum++; // put this in the addSectionTitle
        rowNum = generateHeader(sheet, rowNum);
        rowNum = generateData(sheet, rowNum);
        rowNum = generateFooter(sheet, rowNum);
        return rowNum;
    }


    public void addSectionTitle(XSSFSheet sheet, int rowNum) {
        Row headerRow = sheet.createRow(rowNum);
        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue(this.title);

        // Apply style if needed (e.g., bold, font size)
        CellStyle style = sheet.getWorkbook().createCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        style.setFont(font);
        headerCell.setCellStyle(style);
    }

}
