package com.systex.excelgenerator.component;

import org.apache.poi.xddf.usermodel.PresetColor;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;

public class RadarChartSection extends AbstractChartSection {

    @Override
    protected XDDFChartData createChartData(XSSFChart chart) {
        // 設定圖表的軸
        XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
        valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        XDDFRadarChartData radarData = (XDDFRadarChartData) chart.createData(ChartTypes.RADAR, categoryAxis, valueAxis);
        radarData.setStyle(RadarStyle.FILLED);

        return radarData;
    }

    @Override
    protected void setChartItems(XSSFChart chart, XDDFChartData data) {
        // 設定為填充式雷達圖
        ((XDDFRadarChartData) data).setStyle(RadarStyle.FILLED);

        XDDFSolidFillProperties fillProperties = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.ORANGE)); //
        ((XDDFRadarChartData.Series) data.getSeries().get(0)).setFillProperties(fillProperties);

        chart.getCTChart().getPlotArea().getCatAxArray(0).addNewMajorGridlines();
        chart.getCTChart().getPlotArea().getValAxArray(0).addNewMajorGridlines();
    }
}
