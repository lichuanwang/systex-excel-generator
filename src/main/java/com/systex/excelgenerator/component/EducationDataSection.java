package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Education;
import com.systex.excelgenerator.utils.ExcelStyleAndSheetUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import com.systex.excelgenerator.utils.FormulaUtil;
import com.systex.excelgenerator.utils.NamedCellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

public class EducationDataSection extends AbstractDataSection<Education> {

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
        headerRow.createCell(startCol).setCellValue("School Name");
        headerRow.createCell(startCol + 1).setCellValue("Major");
        headerRow.createCell(startCol + 2).setCellValue("Grade");
        headerRow.createCell(startCol + 3).setCellValue("Start Date");
        headerRow.createCell(startCol + 4).setCellValue("End Date");
        headerRow.createCell(startCol + 5).setCellValue("Date Interval");
    }

    protected void renderBody(ExcelSheet sheet, int startRow, int startCol) {
        XSSFWorkbook workbook = (XSSFWorkbook) sheet.getWorkbook();
        CellStyle dateStyle = ExcelStyleAndSheetUtils.dateFormatting(workbook);

        int rowNum = startRow; // Start from the row after the header

        for (Education edu : content) {
            Row row = sheet.createOrGetRow(rowNum++);
            row.createCell(startCol).setCellValue(edu.getSchoolName());
            row.createCell(startCol + 1).setCellValue(edu.getMajor());
            row.createCell(startCol + 2).setCellValue(edu.getGrade());
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

            Map<String , NamedCellReference> replacemap = new HashMap<>();
            replacemap.put("startCellRef" , new NamedCellReference("K62"));
            replacemap.put("endCellRef" , new NamedCellReference(row.getRowNum() , startCol + 4  , true , true));

            row.createCell(startCol + 5).setCellFormula(FormulaUtil.parseFormula(formula , replacemap));

            formula = "${SheetName}!${CellRef}";

            replacemap.clear();
            replacemap.put("SheetName" , new NamedCellReference("JohnDoe"));
            replacemap.put("CellRef" , new NamedCellReference("A6"));

            row.createCell(startCol + 6).setCellFormula(FormulaUtil.parseFormula(formula , replacemap));

            // 要替換的佔位符set
//            Set<NamedCellReference> replaceSet = new HashSet<>();
//            replaceSet.add(new NamedCellReference("startCellRef" , row.getRowNum() , startCol + 3));
//            replaceSet.add(new NamedCellReference("endCellRef" , row.getRowNum() , startCol + 4));

//            Map<String , String> replacemap = new HashMap<>();
//            replacemap.put("startCellRef" , "K62");
//            replacemap.put("endCellRef" , "L62");

//            Map<String , Map<Integer , Integer >> replacemap2 = new HashMap<>();
//            replacemap2.put("startCellRef" , )

            // 計算過後的時間區間的值
            //row.createCell(startCol + 5).setCellFormula(FormulaUtil.parseFormula2(replaceSet , formula));
//            row.createCell(startCol + 5).setCellFormula(FormulaUtil.parseFormula(formula , replacemap));
        }
    }

    protected void renderFooter(ExcelSheet sheet, int startRow, int startCol) {
        // implement footer logic here
    }
}