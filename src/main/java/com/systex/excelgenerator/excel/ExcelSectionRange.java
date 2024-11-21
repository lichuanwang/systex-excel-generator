package com.systex.excelgenerator.excel;

class ExcelSectionRange {
    private final int startRow;
    private final int startCol;
    private final int endRow;
    private final int endCol;

    ExcelSectionRange(int startRow, int startCol, int endRow, int endCol) {
        this.startRow = startRow;
        this.startCol = startCol;
        this.endRow = endRow;
        this.endCol = endCol;
    }

    public int getStartRow() { return startRow; }

    public int getStartCol() { return startCol; }

    public int getEndRow() { return endRow; }

    public int getEndCol() { return endCol; }
}
