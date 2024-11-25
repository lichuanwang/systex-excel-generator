package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import java.util.Collection;

public interface DataSection<T> extends Section {
    void setData(Collection<T> data);
    boolean isEmpty();
    void render(ExcelSheet sheet, int startRow, int startCol);
    int getDataStartRow();
    int getDataEndRow();
    int getDataStartCol();
    int getDataEndCol();
}

