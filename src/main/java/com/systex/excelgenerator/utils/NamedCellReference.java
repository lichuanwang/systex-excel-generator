package com.systex.excelgenerator.utils;

import org.apache.poi.ss.util.CellReference;

import java.util.Objects;

public class NamedCellReference {

    private final String cellName;
    private final CellReference cellReference;

    public NamedCellReference(String cellName, int pRow, int pCol) {
        this(cellName, null, pRow, pCol, false, false);
    }

    // Apache POI CellReference Constructor 所有功能 , 多了一個cellName(?)
    public NamedCellReference(String cellName, String pSheetName, int pRow, int pCol, boolean pAbsRow, boolean pAbsCol) {
        this.cellReference = new CellReference(pSheetName, pRow, pCol, pAbsRow, pAbsCol);
        this.cellName = cellName;
    }

    public String getCellName() {
        return cellName;
    }

    public CellReference getCellReference() {
        return cellReference;
    }

    public String formatAsString() {
        return cellReference.formatAsString();
    }

    @Override
    public final boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof NamedCellReference that)) return false;

        return Objects.equals(cellName, that.cellName) && Objects.equals(cellReference, that.cellReference);
    }

    @Override
    public int hashCode() {
        int result = Objects.hashCode(cellName);
        result = 31 * result + Objects.hashCode(cellReference);
        return result;
    }
}
