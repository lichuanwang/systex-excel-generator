package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;

public abstract class AbstractChartSection implements ChartSection {

    protected int col1;
    protected int row1;
    protected int col2;
    protected int row2;
    protected int dataFirstRow;
    protected int dataLastRow;
    protected int xAxisCol;
    protected int yAxisCol;

    // 設定圖表的位置
    public void setChartPosition(int startingRow, int startingColumn, int endingRow, int endingColumn) {
        // default size 給使用者col2 , row2
        this.row1 = startingRow;
        this.col1 = startingColumn;

        this.col2 = col1 + 7;
        this.row2 = endingRow + 15;
    }

    public void setDataSource(Section<?> section) {
        this.dataFirstRow = section.getDataStartRow();
        this.dataLastRow = section.getDataEndRow();
        this.xAxisCol = section.getDataStartCol();
        this.yAxisCol = section.getDataEndCol();
    }

    // 決定是甚麼圖表類型
    protected abstract XDDFChartData createChartData(XSSFChart chart, XDDFCategoryAxis categoryAxis, XDDFValueAxis valueAxis);

    // 各個圖表的特有設定
    protected abstract void setChartItems(XSSFChart chart, XDDFChartData data);

    // 各個圖表共通有的東西

    public void render(ExcelSheet sheet){

        // 設定sheet中的畫布
        XSSFDrawing drawing = sheet.getXssfSheet().createDrawingPatriarch();

        // 設定圖表位置
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0,0,0,0, col1 , row1 , col2 , row2));

        System.out.println(dataFirstRow+","+dataLastRow);

        // 選定資料範圍類別的資料來源
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                sheet.getXssfSheet(), new CellRangeAddress(dataFirstRow, dataLastRow, xAxisCol, xAxisCol));

        // 選定資料範圍數值的資料來源
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                sheet.getXssfSheet(), new CellRangeAddress(dataFirstRow, dataLastRow, yAxisCol, yAxisCol));

        // 設定圖表的軸
        XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
        valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        // 創建具體的圖表數據並配置
        XDDFChartData data = createChartData(chart, categoryAxis, valueAxis);
        data.addSeries(categories, values);

        // 各圖表特有的設定
        setChartItems(chart, data);

        // 顯示圖表
        chart.plot(data);
    }
}
