package com.systex.excelgenerator.utils;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.util.CellReference;

import java.util.Map;
import java.util.Set;

public class FormulaHandler {

    // 解析公式
    // 使用者知道資料的範圍(Excel中)
    public String parseFormula1(Map<String,String> parameters , String formula){

        // 解析文字 [使用者輸入公式去抓取替換的參數]
        String template = formula;

        // 替換${},把它變成Excel中的格子(ex:A4)
        for (Map.Entry<String, String> entry : parameters.entrySet()){
            template = template.replace("${"+entry.getKey()+"}" , entry.getValue());
        }

        return template;
    }

    // 解析公式
    // 使用者不知道資料的範圍只知道資料是第幾個row和第幾個column
    public <T extends CustomCellReference> String parseFormula2(Set<T> cellRefs , String formula){
        String template = formula;

        for (T cellRef : cellRefs) {
            template = template.replace("${" + cellRef.getCellName() + "}", cellRef.formatAsString());
        }
        return template;
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
