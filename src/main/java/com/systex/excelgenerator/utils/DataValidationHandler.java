package com.systex.excelgenerator.utils;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressList;

public class DataValidationHandler {
    private DataValidationHelper validationHelper;
    private CellRangeAddressList addressList;
    private DataValidationConstraint constraint;
    private Sheet sheet;
    private int firstRow;
    private int lastRow;
    private int firstCol;
    private int lastCol;

    public DataValidationHandler(Sheet sheet , int firstRow , int lastRow , int firstCol , int lastCol) {
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstCol = firstCol;
        this.lastCol = lastCol;
        this.sheet = sheet;
        this.validationHelper = sheet.getDataValidationHelper();

        // 設定資料範圍
        addressList = new CellRangeAddressList(firstRow , lastRow , firstCol , lastCol);
    }

    // 整數資料驗證
    public void IntegerDataValid(String operator , String startRange , String endRange) {
        // 設定cell資料輸入限制
        constraint = validationHelper.createIntegerConstraint(ConvertOperator(operator), startRange, endRange);

        // final setting
        applyValidation();
    }

    // 數字資料驗證(可包含小數...)
    public void NumericDataValid(String operator , String startRange , String endRange) {
        // validationtype預設為decimal(因為上面有整數的驗證了)
        constraint = validationHelper.createNumericConstraint(2 , ConvertOperator(operator) , startRange, endRange);

        // final setting
        applyValidation();
    }

    // 下拉選單資料驗證
    public void ListDataValid(String[] Options){
        // 設定下拉選單的選項
        constraint = validationHelper.createExplicitListConstraint(Options);

        // final setting
        applyValidation();
    }

    // 日期資料驗證
    public void DateValid(String operator , String startDate , String endDate){
        // 預設日期格式為"yyyy/MM/dd"(可更改)
        constraint = validationHelper.createDateConstraint(ConvertOperator(operator)
                , startDate , endDate , "yyyy/MM/dd");

        // final setting
        applyValidation();
    }

    // convert operator
    public int ConvertOperator(String operator){
        switch (operator.toUpperCase()) {
            case "BETWEEN":
                return DataValidationConstraint.OperatorType.BETWEEN;
            case "NOT BETWEEN":
                return DataValidationConstraint.OperatorType.NOT_BETWEEN;
            case "EQUAL":
                return DataValidationConstraint.OperatorType.EQUAL;
            case "NOT EQUAL":
                return DataValidationConstraint.OperatorType.NOT_EQUAL;
            case "GREATER THAN":
                return DataValidationConstraint.OperatorType.GREATER_THAN;
            case "LESS THAN":
                return DataValidationConstraint.OperatorType.LESS_THAN;
            case "GREATER OR EQUAL":
                return DataValidationConstraint.OperatorType.GREATER_OR_EQUAL;
            case "LESS OR EQUAL":
                return DataValidationConstraint.OperatorType.LESS_OR_EQUAL;
            default:
                throw new IllegalArgumentException("This is undefined operator: " + operator);
        }
    }

    public void applyValidation(){

        DataValidation dataValidation = validationHelper.createValidation(constraint, addressList);
        // 設定當使用者輸入不符合選項時顯示錯誤訊息
        dataValidation.setShowErrorBox(true);
        sheet.addValidationData(dataValidation);
    }
}
