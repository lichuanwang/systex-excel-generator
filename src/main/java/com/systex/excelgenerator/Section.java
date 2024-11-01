package com.systex.excelgenerator;

import java.util.Collection;

public interface Section<T> {

    // set single data
    void setData(T data);

    // set more data
    void setData(Collection<T> datacollection);

    boolean isEmpty();
    boolean populate();
}
