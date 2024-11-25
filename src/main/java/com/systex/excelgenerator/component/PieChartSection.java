package com.systex.excelgenerator.component;

import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;

import java.util.ArrayList;
import java.util.List;

public class PieChartSection extends AbstractChartSection {

    public PieChartSection() {
        super("Pie Chart");
    }

    @Override
    protected List<Object> generateChartData() {
        List<Object> data = new ArrayList<>();

        data.add(ChartTypes.PIE);
        data.add(null);
        data.add(null);

        return data;
    }

    @Override
    protected void addAdditionalChartFeature(XSSFChart chart, XDDFChartData data) {
        data.setVaryColors(true);
        // 顯示圖表圖例
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT);

        CTDLbls dLbls = chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewDLbls();
        dLbls.addNewShowCatName().setVal(true);     // 顯示類別名稱
        dLbls.addNewShowVal().setVal(false);        // 不顯示值
        dLbls.addNewShowSerName().setVal(false);    // 不顯示數列名稱
        dLbls.addNewShowPercent().setVal(true);     // 顯示百分比
        dLbls.addNewShowLeaderLines().setVal(true); // 顯示引導線
    }
}
