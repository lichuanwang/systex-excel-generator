package com.systex.excelgenerator.utils;

import com.systex.excelgenerator.excel.ExcelSheet;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;

public class DataValidationUtil {

    private static void addValidation(DataValidationHelper validationHelper, ExcelSheet sheet, int startRow,
                                      int endRow, int startCol, int endCol, DataValidationConstraint constraint) {
        CellRangeAddressList addressList = new CellRangeAddressList(startRow, endRow, startCol, endCol);
        DataValidation dataValidation = validationHelper.createValidation(constraint, addressList);
        // 設定當使用者輸入不符合選項時顯示錯誤訊息
        dataValidation.setShowErrorBox(true);
        sheet.getXssfSheet().addValidationData(dataValidation);
    }

    // 整數資料驗證
    public static void validIntegerData(ExcelSheet sheet , int startRow , int endRow , int startCol , int endCol ,
                                        String operator , String startRange , String endRange) {
        DataValidationHelper validationHelper = sheet.getXssfSheet().getDataValidationHelper();
        // 設定資料限制
        DataValidationConstraint constraint = validationHelper.createIntegerConstraint(
                Operator.fromString(operator).getOperatorType(), startRange, endRange);
        //套用限制
        addValidation(validationHelper, sheet, startRow, endRow, startCol, endCol, constraint);
    }

    // 數字資料驗證(可包含小數...)
    public static void validNumericData(ExcelSheet sheet, int startRow , int endRow , int startCol , int endCol ,
                                        String operator , String startRange , String endRange) {
        DataValidationHelper validationHelper = sheet.getXssfSheet().getDataValidationHelper();
        // 設定資料限制,validationtype預設為decimal(因為上面有整數的驗證了)
        DataValidationConstraint constraint = validationHelper.createNumericConstraint(
                    2 , Operator.fromString(operator).getOperatorType(), startRange, endRange);
        //套用限制
        addValidation(validationHelper, sheet, startRow, endRow, startCol, endCol, constraint);
    }

    // 下拉選單資料驗證
    public static void validListData(ExcelSheet sheet , int startRow , int endRow ,
                                     int startCol , int endCol, String[] options){
        DataValidationHelper validationHelper = sheet.getXssfSheet().getDataValidationHelper();
        // 設定下拉選單的選項
        DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(options);
        //套用限制
        addValidation(validationHelper, sheet, startRow, endRow, startCol, endCol, constraint);
    }

    public enum Operator {
        BETWEEN(DataValidationConstraint.OperatorType.BETWEEN, "BETWEEN"),
        NOT_BETWEEN(DataValidationConstraint.OperatorType.NOT_BETWEEN, "NOT BETWEEN"),
        EQUAL(DataValidationConstraint.OperatorType.EQUAL, "EQUAL"),
        NOT_EQUAL(DataValidationConstraint.OperatorType.NOT_EQUAL, "NOT EQUAL"),
        GREATER_THAN(DataValidationConstraint.OperatorType.GREATER_THAN, "GREATER THAN"),
        LESS_THAN(DataValidationConstraint.OperatorType.LESS_THAN, "LESS THAN"),
        GREATER_OR_EQUAL(DataValidationConstraint.OperatorType.GREATER_OR_EQUAL, "GREATER OR EQUAL"),
        LESS_OR_EQUAL(DataValidationConstraint.OperatorType.LESS_OR_EQUAL, "LESS OR EQUAL");

        private final int operatorType;
        private final String displayName;

        Operator(int operatorType, String displayName) {
            this.operatorType = operatorType;
            this.displayName = displayName;
        }

        public int getOperatorType() {
            return operatorType;
        }

        public String getDisplayName() {
            return displayName;
        }

        public static Operator fromString(String operator) {
            for (Operator op : Operator.values()) {
                if (op.displayName.equalsIgnoreCase(operator)) {
                    return op;
                }
            }
            throw new IllegalArgumentException("無此運算子: " + operator);
        }
    }
}


