package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Experience;
import com.systex.excelgenerator.utils.FormattingHandler;
import com.systex.excelgenerator.utils.FormulaHandler;
import com.systex.excelgenerator.utils.NamedCellReference;
import org.apache.poi.ss.usermodel.Row;

import java.util.HashSet;
import java.util.Set;

public class ExperienceDataSection extends AbstractDataSection<Experience> {

    private FormattingHandler formattingHandler = new FormattingHandler();
    private FormulaHandler formulaHandler = new FormulaHandler();

    public ExperienceDataSection() {
        super("Experience");
    }

    @Override
    public boolean isEmpty() {
        return content == null || content.isEmpty();
    }

    @Override
    public int getWidth() {
        // Set the width based on the number of columns this section uses.
        return 6; // Example width, assuming we have 5 columns for education details
    }

    @Override
    public int getHeight() {
        // Height based on the number of education entries
        return content.size() + 1; // +1 for the header row
    }

    protected void renderHeader(ExcelSheet sheet, int startRow, int startCol) {
        // Create header row for Education section
        Row headerRow = sheet.createOrGetRow(startRow);
        headerRow.createCell(startCol).setCellValue("Company");
        headerRow.createCell(startCol + 1).setCellValue("Role");
        headerRow.createCell(startCol + 2).setCellValue("Description");
        headerRow.createCell(startCol + 3).setCellValue("Start Date");
        headerRow.createCell(startCol + 4).setCellValue("End Date");
        headerRow.createCell(startCol + 5).setCellValue("Date Interval");
    }

    protected void renderBody(ExcelSheet sheet, int startRow, int startCol) {
        int rowNum = startRow; // Start from the row after the header

        for (Experience exp : content) {
            Row row = sheet.createOrGetRow(rowNum++);
            row.createCell(startCol).setCellValue(exp.getCompanyName());
            row.createCell(startCol + 1).setCellValue(exp.getJobTitle());
            row.createCell(startCol + 2).setCellValue(exp.getDescription());
            row.createCell(startCol + 3).setCellValue(exp.getStartDate());

            // format date
            row.getCell(startCol + 3).setCellStyle(formattingHandler.DateFormatting(exp.getStartDate() , sheet.getWorkbook()));

            row.createCell(startCol + 4).setCellValue(exp.getEndDate());

            // format date
            row.getCell(startCol + 4).setCellStyle(formattingHandler.DateFormatting(exp.getStartDate() , sheet.getWorkbook()));

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