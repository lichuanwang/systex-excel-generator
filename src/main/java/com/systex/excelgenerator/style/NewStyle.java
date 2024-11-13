package com.systex.excelgenerator.style;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.io.Serializable;

public class NewStyle implements Serializable{

    private final XSSFCellStyle style;

    public NewStyle(XSSFCellStyle style) {
        this.style = style;
    }
}
