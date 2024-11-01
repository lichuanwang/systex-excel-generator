package com.systex.excelgenerator.utils;

import org.apache.poi.ss.util.CellReference;

import java.util.Map;
import java.util.Set;

public class FormulaHandler {

    // 解析公式
    public String parseFormula(Map<String,Integer> parameters , String formula){
        // 解析文字 [使用者輸入公式去抓取替換的參數]
        String template = """
                IF(DATEDIF(${startCell},${endCell},\"y\")=0
                """;
        return "";
    }

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


        // 解析文字 [使用者輸入公式去抓取替換的參數]
        String formulaTemplate = """
                SUM(${startCellRef1}:F${endCellRef1})+SUM(${startCellRef2}:F${endCellRef2})
                """;

        //Set<? extends CellReference>

        Map<String, String> parameters;

        return formula;
    }

    // 取用另外一個sheet的data
}
