package com.systex.excelgenerator.component;

import java.util.List;
import java.util.Map;

public interface DataSection<T> extends Section {
    void setData(String[] headerColumValues, Map<Integer, List<Object>> data);
    boolean isEmpty();
    int getDataStartRow();
    int getDataEndRow();
    int getDataStartCol();
    int getDataEndCol();
}

