package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.PresetColor;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Collection;

public class RadarChartSection<T> extends AbstractChartSection<T> {

    // set chart type
    @Override
    protected XDDFChartData createChartData(XSSFChart chart, XDDFCategoryAxis categoryAxis, XDDFValueAxis valueAxis) {
        XDDFRadarChartData radarData = (XDDFRadarChartData) chart.createData(ChartTypes.RADAR, categoryAxis, valueAxis);
        radarData.setStyle(RadarStyle.FILLED);
        return radarData;
    }

    @Override
    protected void setChartItems(XSSFChart chart, XDDFChartData data) {
        // 設定為填充式雷達圖
        ((XDDFRadarChartData) data).setStyle(RadarStyle.FILLED);

        XDDFSolidFillProperties fillProperties = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.ORANGE));
        ((XDDFRadarChartData.Series) data.getSeries().get(0)).setFillProperties(fillProperties);

        chart.getCTChart().getPlotArea().getCatAxArray(0).addNewMajorGridlines();
        chart.getCTChart().getPlotArea().getValAxArray(0).addNewMajorGridlines();
    }

    @Override
    public void setData(Collection<T> data) {

    }

    @Override
    public boolean isEmpty() {
        return false;
    }

}
