//package com.systex.excelgenerator.component;
//
//import org.apache.poi.ss.formula.functions.T;
//import org.apache.poi.xddf.usermodel.PresetColor;
//import org.apache.poi.xddf.usermodel.XDDFColor;
//import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
//import org.apache.poi.xddf.usermodel.chart.*;
//import org.apache.poi.xssf.usermodel.XSSFChart;
//import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;
//
//import java.util.Collection;
//
//public class PieChartSection<T> extends AbstractChartSection<T> {
//    // set chart type
//    @Override
//    protected XDDFChartData createChartData(XSSFChart chart) {
//        XDDFChartData pieData = chart.createData(ChartTypes.PIE, null , null);
//        return pieData;
//    }
//
//    @Override
//    protected XDDFChartData createChartData(XSSFChart chart, XDDFCategoryAxis categoryAxis, XDDFValueAxis valueAxis) {
//        return null;
//    }
//
//    @Override
//    protected void setChartItems(XSSFChart chart, XDDFChartData data) {
//        XDDFChartLegend legend = chart.getOrAddLegend();
//        legend.setPosition(LegendPosition.RIGHT);
//
//        // 顯示圖表圖例
//        CTDLbls dLbls = chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewDLbls();
//        dLbls.addNewShowCatName().setVal(true);     // 顯示類別名稱
//        dLbls.addNewShowVal().setVal(false);        // 不顯示值
//        dLbls.addNewShowSerName().setVal(false);    // 不顯示數列名稱
//        dLbls.addNewShowPercent().setVal(true);     // 不顯示百分比
//        dLbls.addNewShowLeaderLines().setVal(true); // 顯示引導線
//
//        data.setVaryColors(true);
//    }
//
//    @Override
//    public void setData(Collection<T> data) {
//
//    }
//
//    @Override
//    public boolean isEmpty() {
//        return false;
//    }
//}
