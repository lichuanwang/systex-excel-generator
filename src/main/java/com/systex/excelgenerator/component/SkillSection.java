package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import com.systex.excelgenerator.model.Skill;
import com.systex.excelgenerator.utils.DataValidationHandler;
import org.apache.poi.ss.usermodel.Row;

public class SkillSection extends AbstractSection<Skill> {

    public SkillSection() {
        super("Skill");
    }

    @Override
    public boolean isEmpty() {
        return content == null || content.isEmpty();
    }

    @Override
    public int getWidth() {
        // Set the width based on the number of columns this section uses.
        return 4; // Example width, assuming we have 4 columns for skill details
    }

    @Override
    public int getHeight() {
        // Height based on the number of education entries
        return content.size() + 1; // +1 for the header row
    }

    protected void populateHeader(ExcelSheet sheet, int startRow, int startCol) {
        // Create header row for Education section
        Row headerRow = sheet.createOrGetRow(startRow);
        headerRow.createCell(startCol).setCellValue("Id");
        headerRow.createCell(startCol + 1).setCellValue("Name");
        headerRow.createCell(startCol + 2).setCellValue("Level");
    }

    protected void populateBody(ExcelSheet sheet, int startRow, int startCol) {
        int rowNum = startRow; // Start from the row after the header

        for (Skill skill : content) {
            Row row = sheet.createOrGetRow(rowNum++);
            row.createCell(startCol).setCellValue(skill.getId());
            row.createCell(startCol + 1).setCellValue(skill.getSkillName());

            // test data valid
            DataValidationHandler dataValidationHandler = new DataValidationHandler(sheet.getXssfSheet()
                    , row.getRowNum() , row.getRowNum() , startCol + 2 , startCol + 2);
            dataValidationHandler.IntegerDataValid("between" , "0" , "5");

            row.createCell(startCol + 2).setCellValue(skill.getLevel());
        }
    }

    protected void populateFooter(ExcelSheet sheet, int startRow, int startCol) {
        // implement footer logic here
    }
}
