package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Education;
import com.systex.excelgenerator.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.text.SimpleDateFormat;
import java.util.List;

public class EducationSection extends Section {

    private List<Education> educations;

    public EducationSection(List<Education> educations) {
        super("Education");
        this.educations = educations;
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

        String[] headers = {"School Name", "Major", "Grade", "Start Date", "End Date"};
        Row headerRow = ExcelUtils.createOrGet(sheet, bodyRow++);
        for (String header : headers) {
            headerRow.createCell(relativeColumn++).setCellValue(header);
        }

        for (Education edu : educations) {
            relativeColumn = flag;
            Row row = ExcelUtils.createOrGet(sheet, bodyRow++);
            Object[] data = {
                    edu.getSchoolName(),
                    edu.getMajor(),
                    edu.getGrade(),
                    edu.getStartDate(),
                    edu.getEndDate()
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
        System.out.println("row number: " + sheet.getPhysicalNumberOfRows());
        return relativeRow;
    }



}
