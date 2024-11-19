package com.systex.excelgenerator.component;

import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;

public class PieChartSection extends AbstractChartSection {

    @Override
    protected XDDFChartData createChartData(XSSFChart chart) {
        XDDFChartData pieData = chart.createData(ChartTypes.PIE, null , null);
        return pieData;
    }

    @Override
    protected void setChartItems(XSSFChart chart, XDDFChartData data) {
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
