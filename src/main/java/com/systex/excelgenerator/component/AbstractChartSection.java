package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;

import java.util.List;

public abstract class AbstractChartSection implements ChartSection {

    protected int height;
    protected int width;
    protected int dataFirstRow;
    protected int dataLastRow;
    protected int xAxisCol;
    protected int yAxisCol;
    protected String chartTitle;
    protected XSSFChart chart;
    protected Section referenceSection;

    protected AbstractChartSection(String chartTitle) {
        this.chartTitle = chartTitle;
    }

    public void setDataSource(DataSection<?> dataSection) {
        this.dataFirstRow = dataSection.getDataStartRow();
        this.dataLastRow = dataSection.getDataEndRow();
        this.xAxisCol = dataSection.getDataStartCol();
        this.yAxisCol = dataSection.getDataEndCol();
        this.referenceSection = dataSection;
    }

    public Section getReferenceSection() {
        return referenceSection;
    }

    public void setHeight(int height) {
        this.height = height;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public int getHeight() {
        return height;
    }

    public int getWidth() {
        return width;
    }

    public String getTitle() {
        return chartTitle;
    }

    protected abstract List<Object> generateChartData();

    // 各個圖表的特有設定
    protected abstract void addAdditionalChartFeature(XSSFChart chart, XDDFChartData data);

    public void render(ExcelSheet sheet, int startRow, int startCol){

        // 設定sheet中的畫布
        XSSFDrawing drawing = sheet.getXssfSheet().createDrawingPatriarch();

        // 設定圖表位置
        chart = drawing.createChart(drawing.createAnchor(0,0,0,0, startCol, startRow, startCol + width, startRow + height));

        // 選定資料範圍類別的資料來源
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                sheet.getXssfSheet(), new CellRangeAddress(dataFirstRow, dataLastRow, xAxisCol, xAxisCol));

        // 選定資料範圍數值的資料來源
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                sheet.getXssfSheet(), new CellRangeAddress(dataFirstRow, dataLastRow, yAxisCol, yAxisCol));

        // 創建具體的圖表數據並配置
        List<Object> chartData = generateChartData();

        XDDFChartData chosenChartData = chart.createData((ChartTypes) chartData.get(0),
                (XDDFChartAxis) chartData.get(1), (XDDFValueAxis) chartData.get(2));

        // bar chart如果沒有用series設定標題會出錯
        XDDFChartData.Series series = chosenChartData.addSeries(categories, values);
        series.setTitle(chartTitle,null);

        // 各圖表特有的設定
        addAdditionalChartFeature(chart, chosenChartData);

        // 顯示圖表
        chart.plot(chosenChartData);
    }
}
