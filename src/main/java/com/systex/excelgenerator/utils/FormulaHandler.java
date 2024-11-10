package com.systex.excelgenerator.utils;

import java.util.Map;
import java.util.Set;

public class FormulaHandler {

    // 解析公式
    // 使用者知道資料的範圍(Excel中)
    public String parseFormula1(Map<String, String> parameters, String formula){
        String template = "";
        if (formula == null) {
            return template;
        } else{
            template = formula;
        }

        // 解析文字 [使用者輸入公式去抓取替換的參數]
        // 替換${},把它變成Excel中的格子(ex:A4)
        for (Map.Entry<String, String> entry : parameters.entrySet()){
            String target = "${" + entry.getKey() + "}";
            String replacement = entry.getValue();
            template = template.replace(target , replacement); //如果使用formula每次只會替換一個,但下一次替換的時候上一次替換的就不會被替換
        }

        return template;
    }

    // 解析公式
    // 使用者不知道資料的範圍只知道資料是第幾個row和第幾個column
    public String parseFormula2(Set<NamedCellReference> cellRefs , String formula){
        String template = formula;

        for (NamedCellReference cellRef : cellRefs) {
            String target = "${" + cellRef.getCellName() + "}";
            String replacement = cellRef.formatAsString();
            template = template.replace(target, replacement);
        }
        return template;
    }
}
