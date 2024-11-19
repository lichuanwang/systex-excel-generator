package com.systex.excelgenerator.component;

import org.apache.poi.xddf.usermodel.PresetColor;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;

public class LineChartSection extends AbstractChartSection{

    @Override
    protected XDDFChartData createChartData(XSSFChart chart) {

        // 設定類別軸和數值軸
        XDDFCategoryAxis xAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis yAxis = chart.createValueAxis(AxisPosition.LEFT);

        // 設定圖表類型為折線圖
        XDDFChartData Linedata = chart.createData(ChartTypes.LINE, xAxis, yAxis);

        return Linedata;
    }

    @Override
    protected void setChartItems(XSSFChart chart, XDDFChartData data) {
        data.setVaryColors(true);
        // 顯示圖表圖例
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT); // 圖表圖例顯示在右邊
    }
}
