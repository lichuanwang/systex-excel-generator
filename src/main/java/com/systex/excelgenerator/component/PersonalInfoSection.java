package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Candidate;
import com.systex.excelgenerator.component.Section;
import com.systex.excelgenerator.utils.ExcelUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.text.SimpleDateFormat;

public class PersonalInfoSection extends Section {

    private Candidate candidate;

    public PersonalInfoSection(Candidate candidate) {
        super("Personal Information");
        this.candidate = candidate;
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

        // Define headers and corresponding data
        String[] headers = {"Name", "Gender", "Birthday", "Phone", "Email", "Address"};
        Object[] data = {
                candidate.getName(),
                candidate.getGender(),
                SimpleDateFormat.getDateInstance().format(candidate.getBirthday()),
                candidate.getPhone(),
                candidate.getEmail(),
                candidate.getAddress().toString()
        };


        Row headerRow = ExcelUtils.createOrGet(sheet, bodyRow++);
        relativeColumn = flag;
        for (String header : headers) {
            headerRow.createCell(relativeColumn++).setCellValue(header);
        }

        Row dataRow = ExcelUtils.createOrGet(sheet, bodyRow++);
        relativeColumn = flag;
        for (Object value : data) {
            dataRow.createCell(relativeColumn++).setCellValue(String.valueOf(value));
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
