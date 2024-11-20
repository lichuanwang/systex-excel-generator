package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Candidate;

import java.util.Collection;

public interface DataSection<T> {
    void setData(Collection<T> data);
    boolean isEmpty();
    int getWidth();
    int getHeight();
    void render(ExcelSheet sheet, int startRow, int startCol);
    String getTitle();
    int getDataStartRow();
    int getDataEndRow();
    int getDataStartCol();
    int getDataEndCol();
}

