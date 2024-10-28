package com.systex.excelgenerator.builder;

import com.systex.excelgenerator.excel.ExcelFile;

public abstract class ExcelBuilder {
    protected ExcelFile excelFile;

    public void createNewExcelFile() {
        excelFile = new ExcelFile();
    }

    public ExcelFile getExcelFile() {
        return excelFile;
    }

    public abstract void buildHeader();
    public abstract void buildSections();
    public abstract void buildFooter();
}
