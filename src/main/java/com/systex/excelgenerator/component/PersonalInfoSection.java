package com.systex.excelgenerator.component;

import com.systex.excelgenerator.model.Candidate;
import com.systex.excelgenerator.component.Section;
import com.systex.excelgenerator.utils.DataValidationHandler;
import com.systex.excelgenerator.utils.FormattingHandler;
import com.systex.excelgenerator.utils.ImageHandler;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.IOException;
import java.text.SimpleDateFormat;

public class PersonalInfoSection extends Section {

    private Candidate candidate;
    private FormattingHandler formattingHandler;
    private ImageHandler imageHandler;
    private DataValidationHandler dataValidationHandler;

    {
        this.formattingHandler = new FormattingHandler();
        this.imageHandler = new ImageHandler();
    }

    public PersonalInfoSection(Candidate candidate) {
        super("Personal Information");
        this.candidate = candidate;
    }

    @Override
    public int populate(XSSFSheet sheet, int rowNum){
        addHeader(sheet, rowNum);
        rowNum++;

        // 資料開始的第一個row (test : 凍結名字儲存格)
        int startRow = rowNum;

        Row row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Name");
        row.createCell(1).setCellValue(candidate.getName());

        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Gender");
        row.createCell(1).setCellValue(candidate.getGender());

        // test data valid - gender : male/female
        dataValidationHandler = new DataValidationHandler(sheet , row.getRowNum() , row.getRowNum() , 1 , 1);
        String[] options = {"Male","Female"};
        dataValidationHandler.ListDataValid(options);

        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Birthday");
        row.createCell(1).setCellValue(SimpleDateFormat.getDateInstance().format(candidate.getBirthday()));

        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Phone");

        // format phone number
        Cell personelCell = row.createCell(1);
        personelCell.setCellValue(candidate.getPhone());
        personelCell.setCellStyle(formattingHandler.TextFormatting(sheet.getWorkbook()));

        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Email");
        row.createCell(1).setCellValue(candidate.getEmail());

        row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue("Address");
        row.createCell(1).setCellValue(candidate.getAddress().toString());

        // test image
        imageHandler.insertImage(sheet , 2 , startRow , "C:\\Users\\2400823\\Downloads\\test.jpg");

        // test 凍結儲存格(name)
        formattingHandler.freezeCell(sheet , startRow-1,startRow+1);

        return rowNum;
    }
}
