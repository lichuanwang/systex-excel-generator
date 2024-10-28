package com.systex.excelgenerator.excel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelFile {
    private XSSFWorkbook workbook;

    public ExcelFile() {
        this.workbook = new XSSFWorkbook();
    }

    // Method to add a new sheet
    public XSSFSheet createSheet(String sheetName) {
        return workbook.createSheet(sheetName);
    }

    // Method to save the Excel file to a specified path
    public void saveToFile(String filePath) throws IOException {
        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
        }
    }

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }
}
