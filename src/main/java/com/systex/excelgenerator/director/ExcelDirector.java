package com.systex.excelgenerator.director;

import com.systex.excelgenerator.builder.ExcelBuilder;
import com.systex.excelgenerator.excel.ExcelFile;

public class ExcelDirector {
    private ExcelBuilder excelBuilder;

    public ExcelDirector(ExcelBuilder builder) {
        this.excelBuilder = builder;
    }

    public void constructExcelFile() {
        excelBuilder.createNewExcelFile();
        excelBuilder.buildHeader();
        excelBuilder.buildBody();
        excelBuilder.buildFooter();
    }

    public ExcelFile getExcelFile() {
        return excelBuilder.getExcelFile();
    }
}
