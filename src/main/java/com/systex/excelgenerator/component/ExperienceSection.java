package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Experience;
import com.systex.excelgenerator.utils.FormattingHandler;
import com.systex.excelgenerator.utils.FormulaHandler;
import com.systex.excelgenerator.utils.NamedCellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.*;

public class ExperienceSection extends Section {

    private List<Experience> experiences;
    private FormattingHandler formattingHandler;
    private FormulaHandler formulaHandler;

    {
        this.formattingHandler = new FormattingHandler();
        this.formulaHandler = new FormulaHandler();
    }

    public ExperienceSection(List<Experience> experiences) {
        super("Experience");
        this.experiences = experiences;
    }

    @Override
    public int populate(XSSFSheet sheet, int rowNum) {
        addHeader(sheet, rowNum);
        rowNum++;

        Row headerRow = sheet.createRow(rowNum++);
        headerRow.createCell(0).setCellValue("Company");
        headerRow.createCell(1).setCellValue("Role");
        headerRow.createCell(2).setCellValue("Description");
        headerRow.createCell(3).setCellValue("Start Date");
        headerRow.createCell(4).setCellValue("End Date");
        headerRow.createCell(5).setCellValue("DateInterval");

        for (Experience exp : experiences) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(exp.getCompanyName());
            row.createCell(1).setCellValue(exp.getJobTitle());
            row.createCell(2).setCellValue(exp.getDescription());

            //  format start and end date
            Cell dateCell = row.createCell(3);
            dateCell.setCellValue(exp.getStartDate());
            dateCell.setCellStyle(formattingHandler.DateFormatting(exp.getStartDate(), sheet.getWorkbook()));

            dateCell = row.createCell(4);
            dateCell.setCellValue(exp.getEndDate());
            dateCell.setCellStyle(formattingHandler.DateFormatting(exp.getEndDate(), sheet.getWorkbook()));

            // cal date interval
            dateCell = row.createCell(5);
            //dateCell.setCellFormula(formulaHandler.calDataInterval(row.getRowNum() , 3 , 4));

            // test parse formula
            String formula1 = """
                    IF(DATEDIF(${startCellRef},${endCellRef},"y")=0,"",
                    DATEDIF(${startCellRef},${endCellRef},"y")&"年")&
                    DATEDIF(${startCellRef},${endCellRef},"ym")&"個月"
                    """;

            String formula2 = """
                    SUM(${startCellRef1}:${endCellRef1})+SUM(${startCellRef2}:${endCellRef2})
                    """;

            System.out.println(formula1);
            System.out.println(formula2);

            // test parse formula1
            Map<String , String> fmap1 = new HashMap<>();
            fmap1.put("startCellRef","D22");
            fmap1.put("endCellRef","E22");

            dateCell.setCellFormula(formulaHandler.parseFormula1(fmap1 , formula1));

            // test parse formula2
            // test parse more params in formula
            dateCell = row.createCell(6);
            dateCell.setCellValue(2);
            dateCell = row.createCell(7);
            dateCell.setCellValue(6);
            dateCell = row.createCell(8);
            dateCell.setCellValue(12);
            dateCell = row.createCell(9);
            dateCell.setCellValue(5);

            // Set<? extends CellReference>
            // 使用者要自己建NamedCellReference Class..?
            Set<NamedCellReference> set1 = new HashSet<>();
            set1.add(new NamedCellReference("startCellRef1" , row.getRowNum() , 6));
            set1.add(new NamedCellReference("endCellRef1" , row.getRowNum() , 7));
            set1.add(new NamedCellReference("startCellRef2" , row.getRowNum() , 8));
            set1.add(new NamedCellReference("endCellRef2" , row.getRowNum() , 9));

            dateCell = row.createCell(10);
            dateCell.setCellFormula(formulaHandler.parseFormula2(set1 , formula2));
        }

        return rowNum;
    }
}
