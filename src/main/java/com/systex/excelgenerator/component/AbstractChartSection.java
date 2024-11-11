package com.systex.excelgenerator.component;

import com.systex.excelgenerator.excel.ExcelSheet;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;

public abstract class AbstractChartSection<T> implements Section<T> {

    protected int col1, row1, col2, row2;
    protected int dataFirstRow, dataLastRow, xAxisCol, yAxisCol;

    // set data
//    public void setData(T data) {
//        if(content != null) {
//            this.content = new ArrayList<>();
//            this.content.add(data);
//        }
//    }
//
//    public void setData(Collection<T> dataCollection) {
//        if (dataCollection != null && !dataCollection.isEmpty()) {
//            this.content = new ArrayList<>(dataCollection);
//        }
//    }
//
//    public boolean isEmpty() {
//        return content == null || content.isEmpty();
//    }
    public int getWidth(){
        return 7;
    }

    public int getHeight(){
        return 15;
    }

    // 設定圖表的位置
    public void setChartPosition(int col1, int row1) {
        this.col1 = col1;
        this.row1 = row1;
        this.col2 = col1 + 7;
        this.row2 = row1 + 15;
    }

    // 設定圖表的資料來源
    public void setDataSource(int dataFirstRow, int dataLastRow, int xAxisCol, int yAxisCol) {
        this.dataFirstRow = dataFirstRow;
        this.dataLastRow = dataLastRow;
        this.xAxisCol = xAxisCol;
        this.yAxisCol = yAxisCol;
    }

    // 決定是甚麼圖表類型
    protected abstract XDDFChartData createChartData(XSSFChart chart, XDDFCategoryAxis categoryAxis, XDDFValueAxis valueAxis);

    // 各個圖表的特有設定
    protected abstract void setChartItems(XSSFChart chart, XDDFChartData data);

    // 各個圖表共通有的東西
    public void render(ExcelSheet sheet, int startRow, int startCol){
        // 設定sheet中的畫布
        XSSFDrawing drawing = sheet.getXssfSheet().createDrawingPatriarch();

        // 設定圖表位置
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0,0,0,0,
                startCol , startRow ,startCol + getWidth() ,startRow + getHeight()));

        // 設定圖表的軸
        XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
        valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);

        // 選定資料範圍類別的資料來源
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                sheet.getXssfSheet(), new CellRangeAddress(dataFirstRow, dataLastRow, xAxisCol, xAxisCol));

        // 選定資料範圍數值的資料來源
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                sheet.getXssfSheet(), new CellRangeAddress(dataFirstRow, dataLastRow, yAxisCol, yAxisCol));

        // 創建具體的圖表數據並配置
        XDDFChartData data = createChartData(chart, categoryAxis, valueAxis);
        data.addSeries(categories, values);

        // 各圖表特有的設定
        setChartItems(chart, data);

        // 顯示圖表
        chart.plot(data);
    }
}
