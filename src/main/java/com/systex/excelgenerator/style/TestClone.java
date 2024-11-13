package com.systex.excelgenerator.style;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class TestClone {
    public static void main(String[] args) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontName("Times New Roman");
        font.setColor(IndexedColors.GREEN.getIndex());
        font.setFontHeightInPoints((short)11);
        font.setItalic(true);
        font.setStrikeout(true);
        cellStyle.setFont(font);
        XSSFSheet sheet = workbook.createSheet();
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Hello World");

        CellStyle clonedStyle = workbook.createCellStyle();

        clonedStyle.cloneStyleFrom(cellStyle);

//        clonedStyle.setFillForegroundColor(IndexedColors.RED.getIndex());

        Cell cell1 = row.createCell(1);
        cell1.setCellValue("test");
        cell1.setCellStyle(clonedStyle);

        System.out.println(cellStyle == clonedStyle);
        System.out.println(cellStyle.equals(clonedStyle));

        try (FileOutputStream out = new FileOutputStream("test.xlsx")) {
            workbook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
