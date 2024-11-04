package com.systex.excelgenerator.utils;

import org.apache.poi.ss.util.CellReference;

public abstract class CustomCellReference extends CellReference implements Namble{
    protected String cellName;

    public CustomCellReference(String cellName, int row, int col) {
        super(row, col);
        this.cellName = cellName;
    }

    @Override
    public String getCellName() {
        return cellName;
    }
}
