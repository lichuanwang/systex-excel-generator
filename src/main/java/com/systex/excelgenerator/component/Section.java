package com.systex.excelgenerator.component;

import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Collection;

public interface Section<T> {

    void setData(T data);
    void setData(Collection<T> dataCollection);
    boolean isEmpty();
    int populate(XSSFSheet sheet, int rowNum);
}
