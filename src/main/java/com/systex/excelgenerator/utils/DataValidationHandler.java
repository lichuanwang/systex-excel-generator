package com.systex.excelgenerator.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataValidationHandler {

    private DataValidationHelper validationHelper;
    private CellRangeAddressList addressList;
    private int firstRow;
    private int lastRow;
    private int firstCol;
    private int lastCol;

    public DataValidationHandler(int firstRow , int lastRow , int firstCol , int lastCol) {
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstCol = firstCol;
        this.lastCol = lastCol;

        // 共用workbook , sheet
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();

        this.validationHelper = sheet.getDataValidationHelper();

        // 設定資料範圍
        addressList = new CellRangeAddressList(firstRow , lastRow , firstCol , lastCol);
    }

    public void DataValidation(){
        // 設定cell資料輸入限制
        DataValidationConstraint constraint = validationHelper.createIntegerConstraint(
                DataValidationConstraint.OperatorType.BETWEEN, "1", "100");

        // final setting
        applyValidation(constraint);
    }

    public void setDropDownMenu(String[] Options){
        // 設定下拉選單的選項
        DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(Options);

        // final setting
        applyValidation(constraint);
    }

    public void applyValidation(DataValidationConstraint constraint){
        // 共用workbook , sheet
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();

        DataValidation dataValidation = validationHelper.createValidation(constraint, addressList);
        sheet.addValidationData(dataValidation);
    }
}
