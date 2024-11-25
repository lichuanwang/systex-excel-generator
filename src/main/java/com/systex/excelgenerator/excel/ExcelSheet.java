package com.systex.excelgenerator.excel;

import com.systex.excelgenerator.component.AbstractChartSection;
import com.systex.excelgenerator.component.DataSection;
import com.systex.excelgenerator.component.ImageDataSection;
import com.systex.excelgenerator.component.Section;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;

public class ExcelSheet {
    private final XSSFSheet xssfSheet;
    private final String sheetName;
    private Map<String, Section> sectionMap = new HashMap<>();
    private final List<ExcelSectionRange> sectionRanges = new ArrayList<>();

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
        if (!isEmptyCell(dataSection, startingPoint)) {
            throw new IllegalArgumentException("資料重疊在"+cellReference);
        }

        // add section to map
        this.sectionMap.put(dataSection.getTitle(), dataSection);

        // Render the section at the calculated starting position
        dataSection.render(this, startingPoint[0], startingPoint[1]);
    }

    // add image sections
    public <T> void addSection(ImageDataSection dataSection, String imageType , String cellReference){
        int[] startingPoint = parseCellReference(cellReference);

        // Cell is empty or not empty can add section
        if (!isEmptyCell(dataSection, startingPoint)) {
            throw new IllegalArgumentException("資料重疊在"+cellReference);
        }

        // set image type
        dataSection.setImageType(imageType);

        this.sectionMap.put(dataSection.getTitle(), dataSection);

        dataSection.render(this, startingPoint[0], startingPoint[1]);
    }

    // add chart sections
    public <T> void addChartSection(String cellReference, AbstractChartSection chartSection, String referenceSectionTitle, int chartHeight, int chartWidth) {
        // 傳section name進來再去查找
        DataSection<T> dataSection = getSectionByName(referenceSectionTitle);

        int[] startingPoint = parseCellReference(cellReference);

        chartSection.setHeight(chartHeight);
        chartSection.setWidth(chartWidth);

        // set chart position
        chartSection.setChartPosition(startingPoint[0], startingPoint[1]);

        // Cell is empty or not empty can add section
        if (!isEmptyCell(chartSection, startingPoint)) {
            throw new IllegalArgumentException("資料重疊在"+cellReference);
        }

        // set chart data source
        chartSection.setDataSource(dataSection);

         this.sectionMap.put(dataSection.getTitle() + " " + chartSection.getTitle(), chartSection);

        // render chart sections
        chartSection.render(this);
    }

    // 判斷儲存格內是否有資料
    private <T> boolean isEmptyCell(Section section , int[] startingPoint) {

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
}