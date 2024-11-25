//package com.systex.excelgenerator.component;
//
//import com.systex.excelgenerator.excel.ExcelSheet;
//import com.systex.excelgenerator.model.Skill;
//import com.systex.excelgenerator.utils.DataValidationHandler;
//import org.apache.poi.ss.usermodel.Row;
//import com.systex.excelgenerator.utils.FormattingAndFilter;
//
//public class SkillDataSection extends AbstractDataSection<Skill> {
//
//    private FormattingAndFilter formattingAndfilter = new FormattingAndFilter();
//    public SkillDataSection() {
//        super("Skill");
//    }
//
//    @Override
//    public boolean isEmpty() {
//        return content == null || content.isEmpty();
//    }
//
//    @Override
//    public int getWidth() {
//        // Set the width based on the number of columns this section uses.
//        return 4; // Example width, assuming we have 4 columns for skill details
//    }
//
//    @Override
//    public int getHeight() {
//        // Height based on the number of education entries
//        return content.size() + 1; // +1 for the header row
//    }
//
//    protected void renderHeader(ExcelSheet sheet, int startRow, int startCol) {
//        // Create header row for Education section
//        Row headerRow = sheet.createOrGetRow(startRow);
//        headerRow.createCell(startCol++).setCellValue("Id");
//        this.dataStartColumn = startCol;
//        headerRow.createCell(startCol++).setCellValue("Name");
//        headerRow.createCell(startCol).setCellValue("Level");
//        this.dataEndColumn = startCol;
//    }
//
//    protected void renderBody(ExcelSheet sheet, int startRow, int startCol) {
//        int rowNum = startRow; // Start from the row after the header
//        this.dataStartRow = startRow;
//        for (Skill skill : content) {
//            int colNum = startCol;
//            Row row = sheet.createOrGetRow(rowNum++);
//            row.createCell(colNum++).setCellValue(skill.getId());
//            row.createCell(colNum++).setCellValue(skill.getSkillName());
//
//            // test data valid , set skill level between 0-5
//            DataValidationHandler dataValidationHandler = new DataValidationHandler(sheet.getXssfSheet()
//                    , row.getRowNum() , row.getRowNum() , startCol + 2 , startCol + 2);
//            dataValidationHandler.IntegerDataValid("between" , "0" , "5");
//
//            row.createCell(colNum).setCellValue(skill.getLevel());
//            // if skill level > 2 (conditional test)
//            formattingAndfilter.ConditionalFormatting(sheet.getXssfSheet() , "2"
//                    , row.getRowNum() , row.getRowNum() , startCol + 2);
//        }
//        this.dataEndRow = rowNum - 1;
//    }
//
//    protected void renderFooter(ExcelSheet sheet, int startRow, int startCol) {
//        // implement footer logic here
//    }
//}
