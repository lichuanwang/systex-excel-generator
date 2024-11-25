package com.systex.excelgenerator.component;

import org.apache.poi.xddf.usermodel.PresetColor;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;

import java.util.ArrayList;
import java.util.List;

public class LineChartSection extends AbstractChartSection{

    public LineChartSection() {
        super("Line Chart");
    }

    @Override
    protected List<Object> generateChartData(){
        List<Object> data = new ArrayList<>();
        XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);

        data.add(ChartTypes.LINE);
        data.add(categoryAxis);
        data.add(valueAxis);

        return data;
    }

    @Override
    protected void addAdditionalChartFeature(XSSFChart chart, XDDFChartData data) {
        data.setVaryColors(true);
        // 顯示圖表圖例
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT); // 圖表圖例顯示在右邊
    }
}
