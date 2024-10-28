package com.systex.excelgenerator.utils;

import org.apache.poi.ss.util.CellReference;

public class FormulaHandler {

    // 計算時間的區間
    public String calDataInterval(int rowNum, int colNum1, int colNum2) {
        // date start and end
        String startCellRef = new CellReference(rowNum, colNum1).formatAsString();
        String endCellRef = new CellReference(rowNum, colNum2).formatAsString();

        // set formula
        String formula = (
                "IF(DATEDIF(" + startCellRef + ", " + endCellRef + ", \"y\")=0, \"\", " +
                "DATEDIF(" + startCellRef + ", " + endCellRef + ", \"y\") & \"年\") & " +
                "DATEDIF(" + startCellRef + ", " + endCellRef + ", \"ym\") & \"個月\"");

        return formula;
    }
}
