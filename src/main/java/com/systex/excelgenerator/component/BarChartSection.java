package com.systex.excelgenerator.component;

import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;

public class BarChartSection extends AbstractChartSection {

    @Override
    protected XDDFChartData createChartData(XSSFChart chart) {
        // 設定類別軸和數值軸
        XDDFCategoryAxis xAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis yAxis = chart.createValueAxis(AxisPosition.LEFT);

        // 設為3DBarChart (XDDF'Bar3DChartData'比XDDF'ChartData'有更多細節設定)
        XDDFBar3DChartData barChartData = (XDDFBar3DChartData) chart.createData(ChartTypes.BAR3D, xAxis, yAxis);
        // 可以改成直條圖(改方向)
        barChartData.setBarDirection(BarDirection.COL);

        return barChartData;
    }

    @Override
    protected void setChartItems(XSSFChart chart, XDDFChartData data) {
        data.setVaryColors(true);
        // 顯示圖表圖例
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT); // 圖表圖例顯示在右邊
    }
}
