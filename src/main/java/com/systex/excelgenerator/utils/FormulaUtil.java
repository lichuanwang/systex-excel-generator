package com.systex.excelgenerator.utils;

import java.util.Map;

/**
 * 工具類，用於解析Excel公式並將自定義參數替換為對應的儲存格座標。
 */
public class FormulaUtil {

    /**
     * 私有構造方法，防止實例化工具類。
     */
    private FormulaUtil() {}

    /**
     * 解析公式模板，將其中的占位符替換為指定的儲存格引用。
     * <p>
     * 公式中的占位符應以 `${parameterName}` 的形式出現，並在參數映射表中找到對應的值。
     * 如果對應的 `NamedCellReference` 為null，則會拋出異常。
     * </p>
     *
     * @param formula    使用者自行輸入的公式
     * @param parameters 占位符名稱與對應儲存格引用的映射
     * @return 替換後的公式字符串
     * @throws IllegalArgumentException 如果 formula 或 parameters 為null，或映射中的值為null
     *                                  或參數中的占位符無法匹配
     */
    public static String parse(String formula , Map<String , NamedCellReference> parameters){
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