package com.systex.excelgenerator.component;

import java.util.Collection;

public interface DataSection<T> extends Section {
    void setData(Collection<T> data);
    boolean isEmpty();
    int getDataStartRow();
    int getDataEndRow();
    int getDataStartCol();
    int getDataEndCol();
}

