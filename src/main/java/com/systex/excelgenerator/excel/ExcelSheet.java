package com.systex.excelgenerator.excel;

import com.systex.excelgenerator.component.AbstractChartSection;
import com.systex.excelgenerator.component.DataSection;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

public class ExcelSheet {
    private static final Logger log = LogManager.getLogger(ExcelSheet.class);
    private final XSSFSheet xssfSheet;
    private final String sheetName;
    private Map<String, DataSection<?>> sectionMap = new HashMap<>();
    private int startingRow = 0;
    private int startingCol = 0;
    private int maxColPerRow;
    private int deepestRowOnCurrentLevel = 0;
    private final List<SectionRange> sectionRanges = new ArrayList<>();

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

    public <T> void addSection(DataSection<T> dataSection, Collection<T> dataCollection, String dataStart) {
        // Validate that the section is not empty
        if (dataCollection == null) {
            System.out.println("Please provide data collection for your section");
            return;
        }

        // set data for specify section
        dataSection.setData(dataCollection);

        // add section to list
        this.sectionMap.put(dataSection.getTitle(), dataSection);

        // Cell is empty or not empty can add section
        if (!isCellEmpty(dataSection , dataStart)) {
            throw new IllegalArgumentException("資料重疊在"+dataStart);
        }

        // Determine starting position for the section
        adjustLayoutForNewSection(dataStart);

        // Render the section at the calculated starting position
        dataSection.render(this, startingRow, startingCol);

        // Update layout positions after the section is rendered
        updateLayoutAfterSection(dataSection);
    }

    // 判斷儲存格內是否有資料
    private <T> boolean isCellEmpty(DataSection<T> dataSection , String data) {

        CellReference dataStart = new CellReference(data);

        int startRow = dataStart.getRow();
        int startCol = dataStart.getCol();
        int endRow = startRow + dataSection.getHeight();
        int endCol = startCol + dataSection.getWidth();

        if (isCellRangeOverlap(startRow, startCol, endRow, endCol)) {
            return false;
        }

        // 如果沒有交集的話再把section的位置加入之後要比對的section range list裡面
        this.sectionRanges.add(new SectionRange(startRow, startCol, endRow, endCol));
        return true;
    }

    // 跟每個section去做比對
    private boolean isCellRangeOverlap(int startRow, int startCol, int endRow, int endCol) {
        for (SectionRange range : sectionRanges) {
            if (isOverlap(range.startRow, range.startCol, range.endRow, range.endCol, startRow, startCol, endRow, endCol)) {
                return true;
            }
        }
        return false;
    }

    // 判斷兩個範圍是否有交集
    private boolean isOverlap(int startRow1, int startCol1, int endRow1, int endCol1,
                              int startRow2, int startCol2, int endRow2, int endCol2) {
        // 最晚開始的row or col <= 最早結束的row or col
        boolean isRowOverlap = Math.max(startRow1, startRow2) <= Math.min(endRow1, endRow2);
        boolean isColOverlap = Math.max(startCol1, startCol2) <= Math.min(endCol1, endCol2);

        // row and col都有交集才是有重疊到
        return isRowOverlap && isColOverlap;
    }

    private void adjustLayoutForNewSection(String dataStart) {
        CellReference data = new CellReference(dataStart);
        startingRow = data.getRow();
        startingCol = data.getCol();

    }

    private <T> void updateLayoutAfterSection(DataSection<T> dataSection) {
        // Update layout positions for the next section
        startingCol += dataSection.getWidth();
        deepestRowOnCurrentLevel = Math.max(deepestRowOnCurrentLevel, startingRow + dataSection.getHeight() + 1);
    }

    public <T> DataSection<T> getSectionByName(String name) {

        return (DataSection<T>) sectionMap.get(name);
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
        DataSection<?> dataSection = getSectionByName(sectionTitle);

        // set chart position
        chartSection.setChartPosition(startingRow, getMaxColPerRow() + 1, startingRow + 7, startingCol+12);

        // set chart data source
        chartSection.setDataSource(dataSection);

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

    // Section 範圍記錄類
    private static class SectionRange {
        int startRow, startCol, endRow, endCol;

        public SectionRange(int startRow, int startCol, int endRow, int endCol) {
            this.startRow = startRow;
            this.startCol = startCol;
            this.endRow = endRow;
            this.endCol = endCol;
        }
    }
}