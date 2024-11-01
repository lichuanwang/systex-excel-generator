package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Experience;
import com.systex.excelgenerator.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

public class ExperienceSection extends Section {

    private List<Experience> experiences;

    public ExperienceSection(List<Experience> experiences) {
        super("Experience");
        this.experiences = experiences;
    }

    @Override
    public int populate(XSSFSheet sheet) {
        addHeader(sheet);
        if(sheet.getPhysicalNumberOfRows() == 0){
            relativeRow = 0;
            relativeColumn = 0;
        }

        int bodyRow = relativeRow + 1;
        int flag = relativeColumn;

        String[] headers = {"Company", "Role", "Description", "Start Date", "End Date"};

        Row headerRow = ExcelUtils.createOrGet(sheet, bodyRow++);
        relativeColumn = flag;
        for (String header : headers) {
            headerRow.createCell(relativeColumn++).setCellValue(header);
        }

        for (Experience exp : experiences) {
            relativeColumn = flag;
            Row row = ExcelUtils.createOrGet(sheet, bodyRow++);
            Object[] data = {
                    exp.getCompanyName(),
                    exp.getJobTitle(),
                    exp.getDescription(),
                    exp.getStartDate(),
                    exp.getEndDate()
            };
            for (Object value : data) {
                row.createCell(relativeColumn++).setCellValue(String.valueOf(value));
            }
        }

        relativeColumn += ExcelUtils.colStride(2);
        nextRelativeRow = Math.max(relativeRow, bodyRow);
        if (relativeColumn >= maxCol) {
            relativeRow = ExcelUtils.rowStride(nextRelativeRow);
            relativeColumn = 0;
        }

        System.out.println("relativeRow: " + relativeRow);
        System.out.println("relativeColumn: " + relativeColumn);
        System.out.println("nextRelativeRow: " + nextRelativeRow);

        return relativeRow;
    }
}
