package com.systex.excelgenerator.director;

import com.systex.excelgenerator.builder.ExcelBuilder;
import com.systex.excelgenerator.excel.ExcelFile;

import java.io.IOException;

public class ExcelDirector {
    private ExcelBuilder excelBuilder;

    public ExcelDirector(ExcelBuilder builder) {
        this.excelBuilder = builder;
    }

    public void constructExcelFile() throws IOException {
        excelBuilder.createNewExcelFile();
        excelBuilder.buildHeader();
        excelBuilder.buildSections();
        excelBuilder.buildFooter();
    }

    public ExcelFile getExcelFile() {
        return excelBuilder.getExcelFile();
    }
}
