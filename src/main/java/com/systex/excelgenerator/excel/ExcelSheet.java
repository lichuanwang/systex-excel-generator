package com.systex.excelgenerator.excel;

import com.systex.excelgenerator.component.AbstractChartSection;
import com.systex.excelgenerator.component.DataSection;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.Map;

public class ExcelSheet {
    private static final Logger log = LogManager.getLogger(ExcelSheet.class);

    private final XSSFSheet xssfSheet;
    private final String sheetName;
    private Map<String, DataSection<?>> sectionMap = new HashMap<>();
    private boolean[][] grid; // Grid to track cell occupancy
    private final int maxRows;
    private final int maxCols;

    public ExcelSheet(XSSFWorkbook workbook, String sheetName, int maxRows, int maxCols) {
        this.sheetName = sheetName;
        this.xssfSheet = workbook.createSheet(sheetName);
        this.grid = new boolean[maxRows][maxCols];
        this.maxRows = maxRows;
        this.maxCols = maxCols;
    }

    public String getSheetName() {
        return sheetName;
    }

    public Workbook getWorkbook() {
        return xssfSheet.getWorkbook();
    }

    public XSSFSheet getXssfSheet() {
        return xssfSheet;
    }

    public <T> void addSectionAt(String cellReference, DataSection<T> dataSection) {
        // Parse the cell reference
        int[] indices = parseCellReference(cellReference);
        int startRow = indices[0];
        int startCol = indices[1];

        int sectionHeight = dataSection.getHeight();
        int sectionWidth = dataSection.getWidth();

        // Validate placement
        if (!canPlaceSection(startRow, startCol, startRow + sectionHeight , startCol + sectionWidth )) {
            throw new IllegalArgumentException("Cannot place section at " + cellReference + ": overlaps with existing content.");
        }

        // Mark cells as occupied
        markCellsOccupied(startRow, startCol, startRow + sectionHeight , startCol + sectionWidth );

        // Render the section
        dataSection.render(this, startRow, startCol);

        // Add section to the map for tracking
        sectionMap.put(dataSection.getTitle(), dataSection);
    }

    public void addChartSection(AbstractChartSection chartSection, String sectionTitle, String cellReference) {
        // Parse cell reference
        int[] indices = parseCellReference(cellReference);
        int startRow = indices[0];
        int startCol = indices[1];

        // Define chart dimensions (7 rows, 12 columns)
        int endRow = startRow + 7;
        int endCol = startCol + 12;

        // Check if the chart can fit without overlap
        if (!canPlaceSection(startRow, startCol, endRow, endCol)) {
            throw new IllegalArgumentException("Cannot place chart at " + cellReference + ": overlaps with existing content.");
        }

        // Mark cells as occupied
        markCellsOccupied(startRow, startCol, endRow, endCol);

        // Retrieve associated data section
        DataSection<?> dataSection = getSectionByName(sectionTitle);
        if (dataSection == null) {
            throw new IllegalArgumentException("Data section with title '" + sectionTitle + "' does not exist.");
        }

        // Set chart position and data source
        chartSection.setChartPosition(startRow, startCol, endRow, endCol);
        chartSection.setDataSource(dataSection);

        // Render the chart
        chartSection.render(this);
    }

    public <T> DataSection<T> getSectionByName(String name) {
        return (DataSection<T>) sectionMap.get(name);
    }

    // Parse Excel-style cell references like "A1", "B3" into row and column indices
    private int[] parseCellReference(String cellReference) {
        String column = cellReference.replaceAll("\\d", ""); // Extract letters
        String row = cellReference.replaceAll("\\D", ""); // Extract numbers

        int colIndex = 0;
        for (int i = 0; i < column.length(); i++) {
            colIndex = colIndex * 26 + (column.charAt(i) - 'A' + 1);
        }
        colIndex--; // Convert to zero-based index

        int rowIndex = Integer.parseInt(row) - 1; // Convert to zero-based index
        return new int[]{rowIndex, colIndex};
    }

    // Check if a section can fit without overlapping existing content
    private boolean canPlaceSection(int startRow, int startCol, int endRow, int endCol) {
        for (int r = startRow; r <= endRow; r++) {
            for (int c = startCol; c <= endCol; c++) {
                if (r >= maxRows || c >= maxCols || grid[r][c]) {
                    return false; // Out of bounds or overlap detected
                }
            }
        }
        return true;
    }

    // Mark cells in the grid as occupied
    private void markCellsOccupied(int startRow, int startCol, int endRow, int endCol) {
        for (int r = startRow; r <= endRow; r++) {
            for (int c = startCol; c <= endCol; c++) {
                grid[r][c] = true;
            }
        }
    }

    // Create or get a row
    public Row createOrGetRow(int rowNum) {
        Row row = xssfSheet.getRow(rowNum);
        if (row == null) {
            row = xssfSheet.createRow(rowNum);
        }
        return row;
    }

    // Debugging utility to log the grid state
    private void logGridState(int rows, int cols) {
        StringBuilder sb = new StringBuilder("Grid State:\n");
        for (int r = 0; r < rows; r++) {
            for (int c = 0; c < cols; c++) {
                sb.append(grid[r][c] ? "X " : ". ");
            }
            sb.append("\n");
        }
        log.info(sb.toString());
    }
}
