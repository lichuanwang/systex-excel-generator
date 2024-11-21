package com.systex.excelgenerator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class test {
    public static void main(String[] args) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Protected Sheet");

        // 创建一些示例数据
        for (int i = 0; i < 10; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 5; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue("Data " + i + "," + j);
            }
        }

        // 将所有单元格设置为未锁定
        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            for (int columnIndex = 0; columnIndex < sheet.getRow(rowIndex).getLastCellNum(); columnIndex++) {
                Cell cell = row.getCell(columnIndex);
                if (cell == null) cell = row.createCell(columnIndex);
                CellStyle style = workbook.createCellStyle();
                style.setLocked(false); // 设置为未锁定
                cell.setCellStyle(style);
            }
        }

        // 锁定特定范围（例如A1到C3）
        for (int rowIndex = 0; rowIndex <= 2; rowIndex++) { // 行索引0到2
            Row row = sheet.getRow(rowIndex);
            for (int columnIndex = 0; columnIndex <= 2; columnIndex++) { // 列索引0到2
                Cell cell = row.getCell(columnIndex);
                if (cell == null) cell = row.createCell(columnIndex);
                CellStyle lockedStyle = workbook.createCellStyle();
                lockedStyle.setLocked(true); // 设置为锁定
                cell.setCellStyle(lockedStyle);
            }
        }

        // 启用保护
        sheet.protectSheet("password");

        // 保存Excel文件
        try (FileOutputStream fileOut = new FileOutputStream("LockedCellsExample.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}