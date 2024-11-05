package com.systex.excelgenerator.excel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelFile {
    private final XSSFWorkbook workbook;

//    private List<ExcelSheet> sheets;

    private Map<String, ExcelSheet> sheetMap;

    public ExcelFile() {

        this.workbook = new XSSFWorkbook();
        this.sheetMap = new HashMap<>();
    }

    // Method to add a new sheet
    public ExcelSheet createSheet(String sheetName) {
        ExcelSheet excelSheet = new ExcelSheet(workbook, sheetName, 10);
        sheetMap.put(sheetName, excelSheet);
        return excelSheet;
    }

    public ExcelSheet getExelSheet(String sheetName) {
        sheetName = sheetName.trim();
        ExcelSheet excelSheet = sheetMap.get(sheetName);
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
