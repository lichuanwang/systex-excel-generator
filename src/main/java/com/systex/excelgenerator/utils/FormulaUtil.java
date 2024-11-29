package com.systex.excelgenerator.utils;

import java.util.Map;

public class FormulaUtil {

    private FormulaUtil() {}

    public static String parseFormula(String formula , Map<String , NamedCellReference> parameters){
        String template = formula;

        for (Map.Entry<String , NamedCellReference> entry : parameters.entrySet()){
            String target = "${" + entry.getKey() + "}";
            if (entry.getValue() == null){
                throw new IllegalArgumentException("請輸入要替換的儲存格座標或是行列數");
            }
            String replacement = entry.getValue().getReplacement();
            template = template.replace(target , replacement);
        }

        return template;
    }
}
