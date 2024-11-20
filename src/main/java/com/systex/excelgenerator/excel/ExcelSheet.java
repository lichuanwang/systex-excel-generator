package com.systex.excelgenerator.excel;

import com.systex.excelgenerator.component.AbstractChartSection;
import com.systex.excelgenerator.component.DataSection;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;

public class ExcelSheet {
    private final XSSFSheet xssfSheet;
    private final String sheetName;
    private Map<String, DataSection<?>> sectionMap = new HashMap<>();
    private final List<SectionRange> sectionRanges = new ArrayList<>();

    public ExcelSheet(XSSFWorkbook workbook, String sheetName) {
        this.sheetName = sheetName;
        this.xssfSheet = workbook.createSheet(sheetName);
    }

    public String getSheetName() {
        return sheetName;
    }

    public Workbook getWorkbook() {
        return xssfSheet.getWorkbook();
    }

    public <T> void addSection(DataSection<T> dataSection, String cellReference) {

        int[] startingPoint = parseCellReference(cellReference);

        // Cell is empty or not empty can add section
        if (isEmptyCell(dataSection, startingPoint)) {
            throw new IllegalArgumentException("資料重疊在"+cellReference);
        }

        // add section to map
        this.sectionMap.put(dataSection.getTitle(), dataSection);

        // Render the section at the calculated starting position
        dataSection.render(this, startingPoint[0], startingPoint[1]);
    }

    // add chart sections
    public <T> void addChartSection(String cellReference, AbstractChartSection chartSection, String referenceSectionTitle, int chartHeight, int chartWidth) {
        // 傳section name進來再去查找
        DataSection<T> dataSection = getSectionByName(referenceSectionTitle);

        int[] startingPoint = parseCellReference(cellReference);

        // Cell is empty or not empty can add section
        if (isEmptyCell(dataSection, startingPoint)) {
            throw new IllegalArgumentException("資料重疊在"+cellReference);
        }

        // set chart position
        chartSection.setChartPosition(startingPoint[0], startingPoint[1], startingPoint[0] + chartHeight, startingPoint[1]+chartWidth);

        // set chart data source
        chartSection.setDataSource(dataSection);

        // render chart sections
        chartSection.render(this);
    }

    // 判斷儲存格內是否有資料
    private <T> boolean isEmptyCell(DataSection<T> dataSection , int[] startingPoint) {

        int startRow = startingPoint[0];
        int startCol = startingPoint[1];
        int endRow = startRow + dataSection.getHeight();
        int endCol = startCol + dataSection.getWidth();

        if (isCellRangeOverlap(startRow, startCol, endRow, endCol)) {
            return true;
        }

        // 如果沒有交集的話再把section的位置加入之後要比對的section range list裡面
        this.sectionRanges.add(new SectionRange(startRow, startCol, endRow, endCol));
        return false;
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

    private int[] parseCellReference(String cellReference) {
        CellReference data = new CellReference(cellReference);
        int[] result = new int[2];
        result[0] = data.getRow();
        result[1] = data.getCol();
        return result;

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

    // Getter for the underlying XSSFSheet, if needed
    public XSSFSheet getXssfSheet() {
        return xssfSheet;
    }

    // Section 範圍記錄類
    private static class SectionRange {
        int startRow;
        int startCol;
        int endRow;
        int endCol;

        public SectionRange(int startRow, int startCol, int endRow, int endCol) {
            this.startRow = startRow;
            this.startCol = startCol;
            this.endRow = endRow;
            this.endCol = endCol;
        }
    }
}