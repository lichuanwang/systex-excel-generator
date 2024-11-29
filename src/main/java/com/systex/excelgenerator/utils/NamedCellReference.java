package com.systex.excelgenerator.utils;

import lombok.EqualsAndHashCode;
import lombok.ToString;
import org.apache.poi.ss.util.CellReference;

import java.util.Objects;

@EqualsAndHashCode
@ToString
public class NamedCellReference {

    private final CellReference cellReference;
    private final String cellRef;

    public NamedCellReference(String cellRef){
        this.cellRef = cellRef;
        this.cellReference = null;
    }

    public NamedCellReference(int row , int col){
        this(row , col , false , false);
    }

    public NamedCellReference(int row , int col, boolean lockRow, boolean lockCol){
        this.cellRef = null;
        this.cellReference = new CellReference(row , col, lockRow, lockCol);
    }

    public String getReplacement(){
        if (cellRef != null){
            return cellRef;
        }

        Objects.requireNonNull(cellReference);
        return cellReference.formatAsString();
    }
}
