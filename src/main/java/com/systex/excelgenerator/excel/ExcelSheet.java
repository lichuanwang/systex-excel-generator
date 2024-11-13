package com.systex.excelgenerator.excel;

import com.systex.excelgenerator.component.AbstractChartSection;
import com.systex.excelgenerator.component.DataSection;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Collection;
import java.util.Map;
import java.util.TreeMap;

public class ExcelSheet {
    private final XSSFSheet xssfSheet;
    private final String sheetName;
    private Map<String, DataSection<?>> sectionMap;
    private int startingRow = 0;
    private int startingCol = 0;
    private int maxColPerRow;
    private int deepestRowOnCurrentLevel = 0;

    public ExcelSheet(XSSFWorkbook workbook, String sheetName, int maxColPerRow) {
        this.sheetName = sheetName;
        this.xssfSheet = workbook.createSheet(sheetName);
        this.maxColPerRow = maxColPerRow;
        this.sectionMap = new TreeMap<>();
    }

    public String getSheetName() {
        return sheetName;
    }

    public Workbook getWorkbook() {
        return xssfSheet.getWorkbook();
    }

    public <T> void addSection(DataSection<T> dataSection, Collection<T> dataCollection) {
        // Validate that the section is not empty
        if (dataCollection == null) {
            System.out.println("Please provide data collection for your section");
            return;
        }

        // set data for specify section
        dataSection.setData(dataCollection);

        // add section to list
        this.sectionMap.put(dataSection.getTitle(), dataSection);

        // Determine starting position for the section
        adjustLayoutForNewSection(dataSection);

        // Render the section at the calculated starting position
        dataSection.render(this, startingRow, startingCol);

        // Update layout positions after the section is rendered
        updateLayoutAfterSection(dataSection);
    }

    private <T> void adjustLayoutForNewSection(DataSection<T> dataSection) {
        // Check if adding the section would exceed maxColPerRow
        if (startingCol + dataSection.getWidth() > maxColPerRow) {
            // Move to next row if max columns exceeded, leaving a gap
            startingRow = deepestRowOnCurrentLevel + 2;
            startingCol = 0;
        }
    }

    private <T> void updateLayoutAfterSection(DataSection<T> dataSection) {
        // Update layout positions for the next section
        startingCol += dataSection.getWidth();
        deepestRowOnCurrentLevel = Math.max(deepestRowOnCurrentLevel, startingRow + dataSection.getHeight() + 1);
    }

    public DataSection<?> getSectionByName(String name) {
//        Section<T> result = (Section<T>) sectionMap.get(name);


        return sectionMap.get(name);
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
    public void addChartSection(AbstractChartSection chartsection , String sectionname) {

        // 傳section name進來再去查找
        DataSection<?> dataSection = getSectionByName(sectionname);

        // 邏輯有點死
        // 要有錯誤處理

        // set chart position
        chartsection.setChartPosition(getMaxColPerRow() + 1 , startingRow);
        //chartsection.setChartPosition(1, startingRow);
        // set chart data source
        chartsection.setDataSource(dataSection);
        // render chart sections
        chartsection.render(this);

        // 要更新每個圖表生成的位置(還沒做)
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