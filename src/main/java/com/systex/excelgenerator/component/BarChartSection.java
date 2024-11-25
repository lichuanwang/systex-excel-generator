package com.systex.excelgenerator.component;

import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;

import java.util.ArrayList;
import java.util.List;

public class BarChartSection extends AbstractChartSection {

    public BarChartSection() {
        super("Bar Chart");
    }

    @Override
    protected List<Object> generateChartData() {
        List<Object> data = new ArrayList<>();

        XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);

        data.add(ChartTypes.BAR3D);
        data.add(categoryAxis);
        data.add(valueAxis);

        return data;
    }

    @Override
    protected void addAdditionalChartFeature(XSSFChart chart, XDDFChartData data) {
        data.setVaryColors(true);
        ((XDDFBar3DChartData) data).setBarDirection(BarDirection.COL);
        // 顯示圖表圖例
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT); // 圖表圖例顯示在右邊
    }
}
