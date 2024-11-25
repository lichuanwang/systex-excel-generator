package com.systex.excelgenerator.component;


public interface ChartSection extends Section {
    void setDataSource(DataSection<?> section);
    void setHeight(int height);
    void setWidth(int width);
}
