//package com.systex.excelgenerator.component;
//
//import com.systex.excelgenerator.excel.ExcelSheet;
//import com.systex.excelgenerator.model.Project;
//import com.systex.excelgenerator.style.StyleTemplate;
//import com.systex.excelgenerator.utils.HyperlinkHandler;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.CellStyle;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//
//public class ProjectDataSection extends AbstractDataSection<Project> {
//
//    private HyperlinkHandler hyperlinkHandler = new HyperlinkHandler();
//
//    public ProjectDataSection() {
//        super("Project");
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
//        return 5; // Example width, assuming we have 5 columns for project details
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
//        headerRow.createCell(startCol).setCellValue("Project Name");
//        headerRow.createCell(startCol + 1).setCellValue("Role");
//        headerRow.createCell(startCol + 2).setCellValue("Description");
//        headerRow.createCell(startCol + 3).setCellValue("Technologies Used");
//    }
//
//    protected void renderBody(ExcelSheet sheet, int startRow, int startCol) {
//        int rowNum = startRow; // Start from the row after the header
//
//        for (Project project : content) {
//            Row row = sheet.createOrGetRow(rowNum++);
//            row.createCell(startCol).setCellValue(project.getProjectName());
//
//            // Set Outer HyperLink
//            hyperlinkHandler.setHyperLink("https://github.com/ruanyanamy/systex-excel-generator"
//                    , row.getCell(startCol) , sheet.getWorkbook());
//
//            row.createCell(startCol + 1).setCellValue(project.getRole());
//
//            // Set Internal HyperLink
//            hyperlinkHandler.setInternalLink(sheet.getSheetName(), row.getCell(startCol + 1) , sheet.getWorkbook());
//
//            row.createCell(startCol + 2).setCellValue(project.getDescription());
//            XSSFWorkbook workbook = (XSSFWorkbook) sheet.getWorkbook();
//            CellStyle initialStyle = StyleTemplate.createCommonStyle(workbook);
//            Cell technologiesUsed = row.createCell(startCol + 1);
//            technologiesUsed.setCellValue(project.getTechnologiesUsed());
//            technologiesUsed.setCellStyle(initialStyle);
//        }
//    }
//
//    protected void renderFooter(ExcelSheet sheet, int startRow, int startCol) {
//        // implement footer logic here
//    }
//}