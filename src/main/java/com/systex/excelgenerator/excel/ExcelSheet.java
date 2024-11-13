package com.systex.excelgenerator.excel;

import com.systex.excelgenerator.component.AbstractChartSection;
import com.systex.excelgenerator.component.Section;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Collection;
import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;

public class ExcelSheet {
    private final XSSFSheet xssfSheet;
    private final String sheetName;
    private Map<String, Section<?>> sectionMap = new HashMap<>();
    private int startingRow = 0;
    private int startingCol = 0;
    private int maxColPerRow;
    private int deepestRowOnCurrentLevel = 0;

    public ExcelSheet(XSSFWorkbook workbook, String sheetName, int maxColPerRow) {
        this.sheetName = sheetName;
        this.xssfSheet = workbook.createSheet(sheetName);
        this.maxColPerRow = maxColPerRow;
    }

    public String getSheetName() {
        return sheetName;
    }

    public Workbook getWorkbook() {
        return xssfSheet.getWorkbook();
    }

    // for the collection, also think about what kind of data structure might be the best for this case
    public <T> void addSection(Section<T> section, Collection<T> dataCollection) {
        // Validate that the section is not empty
        if (dataCollection == null) {
            System.out.println("Please provide data collection for your section");
            return;
        }

        // set data for specify section
        section.setData(dataCollection);

        // add section to list
        this.sectionMap.put(section.getTitle(), section);

        // Determine starting position for the section
        adjustLayoutForNewSection(section);

        // Render the section at the calculated starting position
        section.render(this, startingRow, startingCol);

        // Update layout positions after the section is rendered
        updateLayoutAfterSection(section);
    }

    private <T> void adjustLayoutForNewSection(Section<T> section) {
        // Check if adding the section would exceed maxColPerRow
        if (startingCol + section.getWidth() > maxColPerRow) {
            // Move to next row if max columns exceeded, leaving a gap
            startingRow = deepestRowOnCurrentLevel + 2;
            startingCol = 0;
        }
    }

    private <T> void updateLayoutAfterSection(Section<T> section) {
        // Update layout positions for the next section
        startingCol += section.getWidth();
        deepestRowOnCurrentLevel = Math.max(deepestRowOnCurrentLevel, startingRow + section.getHeight() + 1);
    }

    // Method to create or get a row
    public Row createOrGetRow(int rowNum) {
        Row row = xssfSheet.getRow(rowNum);
        if (row == null) {
            row = xssfSheet.createRow(rowNum);
        }
        return row;
    }

    // add chart sections
    public void addChartSection(AbstractChartSection chartSection, String sectionTitle) {
        // 傳section name進來再去查找
        // 邏輯有點死
        // 要有錯誤處理
        // set chart position
        Section<?> section = this.sectionMap.get(sectionTitle);
        chartSection.setChartPosition(startingRow, getMaxColPerRow() + 1, startingRow + 7, startingCol+12);

        // set chart data source
        chartSection.setDataSource(section);

        // render chart sections
        chartSection.render(this);
    }

    // Getter for the underlying XSSFSheet, if needed
    public XSSFSheet getXssfSheet() {
        return xssfSheet;
    }


    public int getStartingRow() {
        return startingRow;
    }

    public void setStartingRow(int startingRow) {
        this.startingRow = startingRow;
    }

    public int getStartingCol() {
        return startingCol;
    }

    public void setStartingCol(int startingCol) {
        this.startingCol = startingCol;
    }

    public int getMaxColPerRow() {
        return maxColPerRow;
    }

    public int getDeepestRowOnCurrentLevel() {
        return deepestRowOnCurrentLevel;
    }

    public void setDeepestRowOnCurrentLevel(int deepestRowOnCurrentLevel) {
        this.deepestRowOnCurrentLevel = deepestRowOnCurrentLevel;
    }

    public void setMaxColPerRow(int maxColPerRow) {
        this.maxColPerRow = maxColPerRow;
    }

    public int getMaxColPerRow(int maxColPerRow) {
        return maxColPerRow;
    }
}