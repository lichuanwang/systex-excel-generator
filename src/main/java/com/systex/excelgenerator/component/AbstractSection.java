package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Education;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.Collection;

public abstract class AbstractSection<T> implements Section<T> {
    protected String title;
    protected Collection<T> content;

    public AbstractSection(String title) {
        this.title = title;
    }

    // wrong naming
    protected abstract void populateHeader(ExcelSheet sheet, int startRow, int startCol);
    protected abstract void populateBody(ExcelSheet sheet, int startRow, int startCol);
    protected abstract void populateFooter(ExcelSheet sheet, int startRow, int startCol);


    // problem with pass by value, should we use a rowNum or primitive type to determine in this way
    // probably using pass by reference could be better
    public void render(ExcelSheet sheet, int startRow, int startCol) {
        addSectionTitle(sheet, startRow, startCol);
        populateHeader(sheet, startRow + 1, startCol);
        populateBody(sheet, startRow + 2, startCol);
        populateFooter(sheet, startRow + 3, startCol);
    }

    public void addSectionTitle(ExcelSheet excelSheet, int startRow, int startCol) {
        Row headerRow = excelSheet.createOrGetRow(startRow);
        Cell headerCell = headerRow.createCell(startCol);
        headerCell.setCellValue(this.title);

        // Apply style if needed (e.g., bold, font size)
        CellStyle style = excelSheet.getWorkbook().createCellStyle();
        Font font = excelSheet.getWorkbook().createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        style.setFont(font);
        headerCell.setCellStyle(style);
    }

    public void setData(T data) {
        if( content != null ) {
            this.content = new ArrayList<T>(); // Check if this will return the same thing just like the one below
            this.content.add(data);
        }
    }

    public void setData(Collection<T> dataCollection) {
        if (dataCollection != null && !dataCollection.isEmpty()) {
            this.content = new ArrayList<>(dataCollection);
        }
    }
}
