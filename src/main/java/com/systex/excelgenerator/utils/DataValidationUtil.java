package com.systex.excelgenerator.utils;

import com.systex.excelgenerator.excel.ExcelSheet;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;

/**
 * 工具類，提供與Excel資料驗證相關的工具方法，支援整數、數字及下拉選單等驗證類型。
 */
public class DataValidationUtil {

    /**
     * 添加通用的資料驗證。
     *
     * @param validationHelper 資料驗證輔助工具
     * @param sheet            應用資料驗證的工作表
     * @param startRow         資料驗證應用開始的列數
     * @param endRow           資料驗證應用結束的列數
     * @param startCol         資料驗證應用開始的行數
     * @param endCol           資料驗證應用結束的行數
     * @param constraint       資料驗證約束條件
     */
    private static void addValidation(DataValidationHelper validationHelper, ExcelSheet sheet, int startRow,
                                      int endRow, int startCol, int endCol, DataValidationConstraint constraint) {
        CellRangeAddressList addressList = new CellRangeAddressList(startRow, endRow, startCol, endCol);
        DataValidation dataValidation = validationHelper.createValidation(constraint, addressList);
        // 設定當使用者輸入不符合選項時顯示錯誤訊息
        dataValidation.setShowErrorBox(true);
        sheet.getXssfSheet().addValidationData(dataValidation);
    }

    // 整數資料驗證
    /**
     * 為指定範圍的儲存格添加整數驗證。
     *
     * @param sheet      應用資料驗證的工作表
     * @param startRow   資料驗證應用開始的列數
     * @param endRow     資料驗證應用結束的列數
     * @param startCol   資料驗證應用開始的行數
     * @param endCol     資料驗證應用結束的行數
     * @param operator   運算子
     * @param startRange 指定數值開始的範圍
     * @param endRange   指定數值結束的範圍
     */
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
    /**
     * 為指定範圍的儲存格添加整數驗證。
     *
     * @param sheet      應用資料驗證的工作表
     * @param startRow   資料驗證應用開始的列數
     * @param endRow     資料驗證應用結束的列數
     * @param startCol   資料驗證應用開始的行數
     * @param endCol     資料驗證應用結束的行數
     * @param operator   運算子
     * @param startRange 指定數值開始的範圍
     * @param endRange   指定數值結束的範圍
     */
    public static void validNumericData(ExcelSheet sheet, int startRow , int endRow , int startCol , int endCol ,
                                        String operator , String startRange , String endRange) {
        DataValidationHelper validationHelper = sheet.getXssfSheet().getDataValidationHelper();
        // 設定資料限制,validationtype預設為decimal(因為上面有整數的驗證了)
        DataValidationConstraint constraint = validationHelper.createNumericConstraint(
                    2 , Operator.fromString(operator).getOperatorType(), startRange, endRange);
        //套用限制
        addValidation(validationHelper, sheet, startRow, endRow, startCol, endCol, constraint);
    }
    /**
     * 為指定範圍的儲存格添加整數驗證。
     *
     * @param sheet      應用資料驗證的工作表
     * @param startRow   資料驗證應用開始的列數
     * @param endRow     資料驗證應用結束的列數
     * @param startCol   資料驗證應用開始的行數
     * @param endCol     資料驗證應用結束的行數
     * @param options    選指定下拉式選單中的項目
     */
    // 下拉選單資料驗證
    public static void validListData(ExcelSheet sheet , int startRow , int endRow ,
                                     int startCol , int endCol, String[] options){
        DataValidationHelper validationHelper = sheet.getXssfSheet().getDataValidationHelper();
        // 設定下拉選單的選項
        DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(options);
        //套用限制
        addValidation(validationHelper, sheet, startRow, endRow, startCol, endCol, constraint);
    }
    /**
     * 定義運算子，用於資料驗證。
     */
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

        /**
         * 取得運算子的類型。
         *
         * @return Operator類型
         */
        public int getOperatorType() {
            return operatorType;
        }

        /**
         * 取得運算子的顯示名稱。
         *
         * @return 顯示名稱
         */
        public String getDisplayName() {
            return displayName;
        }

        /**
         * 根據名稱解析運算子。
         *
         * @param operator 運算子名稱
         * @return 對應的運算子
         * @throws IllegalArgumentException 若運算子名稱無效
         */
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
