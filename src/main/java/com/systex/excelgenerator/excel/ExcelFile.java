package com.systex.excelgenerator.excel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelFile {
    private final XSSFWorkbook workbook;

    private List<ExcelSheet> sheets;


    public ExcelFile() {

        this.workbook = new XSSFWorkbook();
        this.sheets = new ArrayList<>();
    }

    // Method to add a new sheet
    public ExcelSheet createSheet(String sheetName) {
        XSSFSheet sheet = workbook.createSheet(sheetName);
        ExcelSheet excelSheet = new ExcelSheet(sheet, 5);
        sheets.add(excelSheet);
        return excelSheet;
    }

    // Method to save the Excel file to a specified path
    public final void saveToFile(String filePath) throws IOException {
        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
        }
    }

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }
}
