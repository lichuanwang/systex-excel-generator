package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;

public interface Section {
    int getHeight();
    int getWidth();
    String getTitle();
    void render(ExcelSheet sheet, int startRow, int startCol);
}
