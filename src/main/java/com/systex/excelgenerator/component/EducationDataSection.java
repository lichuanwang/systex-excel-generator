package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Candidate;
import com.systex.excelgenerator.model.Education;
import com.systex.excelgenerator.style.ExcelFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import com.systex.excelgenerator.utils.FormulaHandler;
import com.systex.excelgenerator.utils.NamedCellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashSet;
import java.util.Set;

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
        return 6; // Example width, assuming we have 5 columns for education details plus one additional column to separate different section
    }

    @Override
    public int getHeight() {
        // Height based on the number of education entries
        return content.size() + 1; // +1 for the header row
    }

    protected void renderHeader(ExcelSheet sheet, int startRow, int startCol) {
        // Create header row for Education section
        Row headerRow = sheet.createOrGetRow(startRow);
        headerRow.createCell(startCol).setCellValue("School Name");
        headerRow.createCell(startCol + 1).setCellValue("Major");
        headerRow.createCell(startCol + 2).setCellValue("Grade");
        headerRow.createCell(startCol + 3).setCellValue("Start Date");
        headerRow.createCell(startCol + 4).setCellValue("End Date");
        headerRow.createCell(startCol + 5).setCellValue("Date Interval");
    }

    protected void renderBody(ExcelSheet sheet, int startRow, int startCol) {
        XSSFWorkbook workbook = (XSSFWorkbook) sheet.getWorkbook();
        CellStyle dateStyle = ExcelFormat.DateFormatting(workbook);

        int rowNum = startRow; // Start from the row after the header

        for (Education edu : content) {
            Row row = sheet.createOrGetRow(rowNum++);
            row.createCell(startCol).setCellValue(edu.getSchoolName());
            row.createCell(startCol + 1).setCellValue(edu.getMajor());
            row.createCell(startCol + 2).setCellValue(edu.getGrade());
            row.createCell(startCol + 3).setCellValue(edu.getStartDate());

            Cell dateCell =  row.createCell(startCol + 3);
            dateCell.setCellValue(edu.getStartDate());
            dateCell.setCellStyle(dateStyle);
            dateCell =  row.createCell(startCol + 4);
            dateCell.setCellValue(edu.getEndDate());
            dateCell.setCellStyle(dateStyle);

            // 計算時間區間(解析公式)
            // 輸入公式
            String formula = """
                    IF(DATEDIF(${startCellRef},${endCellRef},"y")=0,"",
                    DATEDIF(${startCellRef},${endCellRef},"y")&"年")&
                    DATEDIF(${startCellRef},${endCellRef},"ym")&"個月"
                    """;

            // 要替換的佔位符set
            Set<NamedCellReference> replaceSet = new HashSet<>();
            replaceSet.add(new NamedCellReference("startCellRef" , row.getRowNum() , startCol + 3));
            replaceSet.add(new NamedCellReference("endCellRef" , row.getRowNum() , startCol + 4));

            // 計算過後的時間區間的值
            row.createCell(startCol + 5).setCellFormula(formulaHandler.parseFormula2(replaceSet , formula));
        }
    }

    protected void renderFooter(ExcelSheet sheet, int startRow, int startCol) {
        // implement footer logic here
    }
}