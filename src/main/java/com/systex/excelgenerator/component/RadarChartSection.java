package com.systex.excelgenerator.component;

import org.apache.poi.xddf.usermodel.PresetColor;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;

import java.util.ArrayList;
import java.util.List;

public class RadarChartSection extends AbstractChartSection {

    @Override
    protected List<Object> generateChartData() {
        List<Object> data = new ArrayList<>();

        XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
        valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        XDDFRadarChartData radarData = (XDDFRadarChartData) chart.createData(ChartTypes.RADAR, categoryAxis, valueAxis);
        radarData.setStyle(RadarStyle.FILLED);

        data.add(ChartTypes.RADAR);
        data.add(categoryAxis);
        data.add(valueAxis);

        return data;
    }

    @Override
    protected void addAdditionalChartFeature(XSSFChart chart, XDDFChartData data) {
        // 設定為填充式雷達圖
        ((XDDFRadarChartData) data).setStyle(RadarStyle.FILLED);

        XDDFSolidFillProperties fillProperties = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.ORANGE)); //
        ((XDDFRadarChartData.Series) data.getSeries().get(0)).setFillProperties(fillProperties);

        chart.getCTChart().getPlotArea().getCatAxArray(0).addNewMajorGridlines();
        chart.getCTChart().getPlotArea().getValAxArray(0).addNewMajorGridlines();
    }
}
