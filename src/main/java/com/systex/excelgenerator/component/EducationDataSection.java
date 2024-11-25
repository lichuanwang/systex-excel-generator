package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Education;
import com.systex.excelgenerator.utils.ExcelStyleAndSheetHandler;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import com.systex.excelgenerator.utils.FormulaHandler;
import com.systex.excelgenerator.utils.NamedCellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.*;

public class EducationDataSection extends AbstractDataSection<Education> {

    private FormulaHandler formulaHandler = new FormulaHandler();

    public EducationDataSection() {
        super("Education");
    }

    @Override
    public boolean isEmpty() {
        return content == null || content.isEmpty();
    }

    @Override
    public int getWidth() {
        // Set the width based on the number of columns this section uses.
        return 6;
    }

    @Override
    public int getHeight() {
        // Height based on the number of education entries
        return content.size() + 1; // +1 for the header row
    }

    protected void renderHeader(ExcelSheet sheet, int startRow, int startCol) {
        // Create header row for Education section
        Row headerRow = sheet.createOrGetRow(startRow);
        int startingColumnIndex = startCol;

//        for (int i = 0; i < headerColumnValue.length; i++) {
//            headerRow.createCell(startingColumnIndex++).setCellValue(headerColumnValue[i]);
//        }
//        headerRow.createCell(startCol).setCellValue("School Name");
//        headerRow.createCell(startCol + 1).setCellValue("Major");
//        headerRow.createCell(startCol + 2).setCellValue("Grade");
//        headerRow.createCell(startCol + 3).setCellValue("Start Date");
//        headerRow.createCell(startCol + 4).setCellValue("End Date");
//        headerRow.createCell(startCol + 5).setCellValue("Date Interval");
    }

    protected void renderBody(ExcelSheet sheet, int startRow, int startCol) {
        XSSFWorkbook workbook = (XSSFWorkbook) sheet.getWorkbook();
        CellStyle dateStyle = ExcelStyleAndSheetHandler.dateFormatting(workbook); // Reuse a single date style
        int rowNum = startRow;

        for (Map.Entry<Integer, List<Object>> entry : content.entrySet()) {
            Row row = sheet.createOrGetRow(rowNum++);
            List<Object> data = entry.getValue();
            int colNum = startCol;

            for (Object value : data) {
                Cell cell = row.createCell(colNum++);
                setCellValue(cell, value, dateStyle);
            }
        }
    }

    private void setCellValue(Cell cell, Object value, CellStyle dateStyle) {
        if (value == null) {
            cell.setCellValue("");
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof LocalDate) {
            cell.setCellValue((LocalDate) value);
        } else if (value instanceof LocalDateTime) {
            cell.setCellValue((LocalDateTime) value);
        } else if (value instanceof RichTextString) {
            cell.setCellValue((RichTextString) value);
        } else if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    protected void renderFooter(ExcelSheet sheet, int startRow, int startCol) {
        // implement footer logic here
    }
}