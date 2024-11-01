package com.systex.excelgenerator.builder;

import com.systex.excelgenerator.excel.ExcelFile;

import java.io.IOException;

public abstract class ExcelBuilder {
    protected ExcelFile excelFile;

    public void createNewExcelFile() {
        excelFile = new ExcelFile();
    }

    public ExcelFile getExcelFile() {
        return excelFile;
    }

    public abstract void buildHeader();
    public abstract void buildSections() throws IOException;
    public abstract void buildFooter();
}
