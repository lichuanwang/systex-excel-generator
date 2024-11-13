package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;

public interface ChartSection {
    void setChartPosition(int startingColumn, int startingRow, int endingColumn, int endingRow);
    void setDataSource(DataSection<?> section);
    void render(ExcelSheet excelSheet);
}
