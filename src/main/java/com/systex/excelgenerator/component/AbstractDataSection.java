package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Candidate;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.Collection;

public abstract class AbstractDataSection<T> implements DataSection<T> {
    protected String title;
    protected Collection<T> content;
    protected int dataStartRow;
    protected int dataEndRow;
    protected int dataStartColumn;
    protected int dataEndColumn;


    protected AbstractDataSection(String title) {
        this.title = title;
    }

    // wrong naming
    protected abstract void renderHeader(ExcelSheet sheet, int startRow, int startCol);
    protected abstract void renderBody(ExcelSheet sheet, int startRow, int startCol);
    protected abstract void renderFooter(ExcelSheet sheet, int startRow, int startCol);


    // problem with pass by value, should we use a rowNum or primitive type to determine in this way
    // probably using pass by reference could be better
    public void render(ExcelSheet sheet, int startRow, int startCol) {
        if (content == null || content.isEmpty()) {
            return;
        }
        addSectionTitle(sheet, startRow, startCol);
        renderHeader(sheet, startRow + 1, startCol);
        renderBody(sheet, startRow + 2, startCol);
        renderFooter(sheet, startRow + 3, startCol);
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

    public void setData(Collection<T> dataCollection) {
        if (dataCollection != null && !dataCollection.isEmpty()) {
            this.content = new ArrayList<>(dataCollection);
        }
    }

    public String getTitle() {
        return title;
    }

    public int getDataStartRow() {
        return dataStartRow;
    }

    public int getDataEndRow() {
        return dataEndRow;
    }

    public int getDataStartCol() {
        return dataStartColumn;
    }

    public int getDataEndCol() {
        return dataEndColumn;
    }
}
