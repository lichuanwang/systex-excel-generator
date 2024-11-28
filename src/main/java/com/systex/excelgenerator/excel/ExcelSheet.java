package com.systex.excelgenerator.excel;

import com.systex.excelgenerator.component.AbstractChartSection;
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
    private Map<String, Section> sectionMap = new HashMap<>();
    private final List<ExcelSectionRange> sectionRanges = new ArrayList<>();
    private static final Logger logger = LogManager.getLogger(ExcelSheet.class);

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
        if (overlapWithOtherSections(dataSection, startingPoint)) {
            throw new IllegalArgumentException("資料重疊在"+cellReference);
        }

        // add section to map
        this.sectionMap.put(dataSection.getTitle(), dataSection);

        // Render the section at the calculated starting position
        dataSection.render(this, startingPoint[0], startingPoint[1]);
    }

    // add chart sections
    public void addChartSection(String cellReference, AbstractChartSection chartSection, String referenceDataSectionTitle) {
        // 傳section name進來再去查找
        Section dataSection = this.sectionMap.get(referenceDataSectionTitle);

        if (dataSection == null) {
            throw new IllegalArgumentException("No reference section found. Please add dataSection to the sheet first.");
        }

        if (!(dataSection instanceof DataSection)) {
            throw new IllegalArgumentException("Reference section is not a data section. Please refer to a data section when using addChartSection.");
        }

        chartSection.setDataSource((DataSection<?>) dataSection);

        int[] startingPoint = parseCellReference(cellReference);

        // Cell is empty or not empty can add section
        if (overlapWithOtherSections(chartSection, startingPoint)) {
            throw new IllegalArgumentException("資料重疊在" + cellReference);
        }

        this.sectionMap.put(dataSection.getTitle() + " " + chartSection.getTitle(), chartSection);

        // render chart sections
        chartSection.render(this, startingPoint[0], startingPoint[1]);
    }

    // 判斷儲存格內是否有資料
    private boolean overlapWithOtherSections(Section section , int[] startingPoint) {

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
                return true;
            }
        }

        // 如果沒有交集的話再把section的位置加入之後要比對的section range list裡面
        this.sectionRanges.add(new ExcelSectionRange(startRow, startCol, endRow, endCol));
        return false;
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
            // Handle invalid input format
            logger.info("Invalid cell reference provided: " + cellReference);
            throw e; // Re-throw the exception if needed, or handle it here
        } catch (Exception e) {
            // Handle any unexpected errors
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

    // Getter for the underlying XSSFSheet, if needed
    public XSSFSheet getXssfSheet() {
        return xssfSheet;
    }
}