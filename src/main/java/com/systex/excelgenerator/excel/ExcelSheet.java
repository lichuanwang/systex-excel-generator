package com.systex.excelgenerator.excel;

import com.systex.excelgenerator.component.AbstractChartSection;
import com.systex.excelgenerator.component.ChartSection;
import com.systex.excelgenerator.component.DataSection;
import com.systex.excelgenerator.component.Section;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.*;

public class ExcelSheet {
    private final XSSFSheet xssfSheet;
    private final String sheetName;
    private Map<String, List<Object>> sectionMap = new LinkedHashMap<>();
    private final List<ExcelSectionRange> sectionRanges = new ArrayList<>();
    private static final Logger logger = LogManager.getLogger(ExcelSheet.class);

    public ExcelSheet(XSSFWorkbook workbook, String sheetName) {
        this.sheetName = sheetName;
        this.xssfSheet = workbook.createSheet(sheetName);
    }

    public XSSFSheet getXssfSheet() {
        return xssfSheet;
    }

    public String getSheetName() {
        return sheetName;
    }

//    public Map<String, Section> getSectionMap() {
//        return this.sectionMap;
//    }

    public Workbook getWorkbook() {
        return xssfSheet.getWorkbook();
    }

    public void addSection(String cellReference, Section section) {

        int[] startingPoint = parseCellReference(cellReference);

        if (!canPlaceSection(section, startingPoint)) {
            throw new IllegalArgumentException("資料重疊在" + cellReference);
        }

        // add section to map
        ArrayList<Object> sectionData = new ArrayList<>();
        sectionData.add(section);
        sectionData.add(startingPoint);
        this.sectionMap.put(section.getTitle(), sectionData);
    }

    public <T> void addSection(String cellReference, AbstractChartSection chartSection, String referenceTitle) {
        DataSection<T> dataSection = (DataSection<T>) this.sectionMap.get(referenceTitle).get(0);
        if (dataSection == null) {
            throw new IllegalArgumentException("Reference Section Not Found: " + referenceTitle);
        }
        int[] startingPoint = parseCellReference(cellReference);
        if (!canPlaceSection(dataSection, startingPoint)) {
            throw new IllegalArgumentException("There is overlap");
        }

        List<Object> sectionData = new ArrayList<>();
        sectionData.add(chartSection);
        sectionData.add(startingPoint);
        sectionData.add(dataSection);

        this.sectionMap.put(chartSection.getTitle(), sectionData);

    }

    public void render() {
        for (Map.Entry<String, List<Object>> sectionEntry : this.sectionMap.entrySet()) {
            List<Object> sectionEntryValue = sectionEntry.getValue();
            Section section = (Section) sectionEntryValue.get(0);
            if (section instanceof ChartSection) {
                ((ChartSection) section).setDataSource((DataSection<?>) sectionEntryValue.get(2));
            }
            int[] startingPoint = (int[]) sectionEntryValue.get(1);
            section.render(this, startingPoint[0], startingPoint[1]);
        }
    }

    // 判斷儲存格內是否有資料
    private boolean canPlaceSection(Section section , int[] startingPoint) {

        int startRow = startingPoint[0];
        int startCol = startingPoint[1];
        int endRow = startRow + section.getHeight();
        int endCol = startCol + section.getWidth();

        // 跟每個Section去做比對
        for (ExcelSectionRange range : sectionRanges) {
            // 最晚開始的row or col <= 最早結束的row or col
            boolean isRowOverlap = Math.max(range.getStartRow(), startRow) <= Math.min(range.getEndRow(), endRow);
            boolean isColOverlap = Math.max(range.getStartCol(), startCol) <= Math.min(range.getEndCol(), endCol);

            // row and col都有交集才是有重疊到
            if (isRowOverlap && isColOverlap) {
                return false;
            }
        }

        // 如果沒有交集的話再把section的位置加入之後要比對的section range list裡面
        this.sectionRanges.add(new ExcelSectionRange(startRow, startCol, endRow, endCol));
        return true;
    }

    private int[] parseCellReference(String cellReference) {
        // Initialize result array with default values
        int[] result = new int[2];
        try {
            // Null or empty check
            if (cellReference == null || cellReference.trim().isEmpty()) {
                throw new IllegalArgumentException("Cell reference cannot be null or empty.");
            }

            // Parse the cell reference
            CellReference data = new CellReference(cellReference);

            // Extract row and column
            result[0] = data.getRow();
            result[1] = data.getCol();
        } catch (IllegalArgumentException e) {
            logger.info("Invalid cell reference provided: " + cellReference);
            throw e; // Re-throw the exception if needed, or handle it here
        } catch (Exception e) {
            logger.info("An error occurred while parsing cell reference: " + e.getMessage());
            throw new RuntimeException("Failed to parse cell reference: " + cellReference, e);
        }

        return result;
    }

    // Method to create or get a row
    public Row createOrGetRow(int rowNum) {
        Row row = xssfSheet.getRow(rowNum);
        if (row == null) {
            row = xssfSheet.createRow(rowNum);
        }
        return row;
    }
}