package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;

public interface ChartSection extends Section {
    void setChartPosition(int startingColumn, int startingRow);
    void setDataSource(DataSection<?> section);
    void setHeight(int height);
    void setWidth(int width);
    void render(ExcelSheet excelSheet);
}
