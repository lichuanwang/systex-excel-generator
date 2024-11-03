package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;

import java.util.Collection;

public interface Section<T> {

    void setData(T data);
    void setData(Collection<T> dataCollection);
    boolean isEmpty();
    int getWidth();
    int getHeight();
    void render(ExcelSheet sheet, int startRow, int startCol);
}
