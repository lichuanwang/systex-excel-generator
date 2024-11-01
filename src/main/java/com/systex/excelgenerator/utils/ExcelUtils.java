package com.systex.excelgenerator.utils;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

public class ExcelUtils {

    public static Row createOrGet(XSSFSheet sheet, int rowNum) {
        Row row = sheet.getRow(rowNum);
        if (row == null) {
            return sheet.createRow(rowNum);
        }else {
            return row;
        }
    }

    public static int colStride(int stride){
        return stride;
    }

    public static int rowStride(int nextRelRow){
        return nextRelRow + 2;
    }






}
