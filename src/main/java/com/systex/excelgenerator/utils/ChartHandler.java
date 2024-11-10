package com.systex.excelgenerator.utils;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.PresetColor;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;

public class ChartHandler {

    // 四個常用的圖表

    /**
     * 產生圓餅圖
     * @param sheet
     * @param dataStartRow
     * @param dataLastRow
     * @param xAxisCol
     * @param yAxisCol
     * @param ChartStartRow
     * @param headerRow
     */
    public void genPieChart(Sheet sheet, int headerRow , int dataStartRow, int dataLastRow,
                            int xAxisCol, int yAxisCol, int ChartStartRow){
        System.out.println(headerRow+","+dataStartRow+","+dataLastRow+","+xAxisCol+","+yAxisCol+","+ChartStartRow);

        // 創建圖表
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();

        // 從headerRow中取得標題名稱(數值的名稱來當標題)
        Row row = sheet.getRow(headerRow);
        String valueTitle = row.getCell(yAxisCol).getStringCellValue();

        // 設置圖表的位置,高從indexRow開始往下,寬是col1~col2
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 5, ChartStartRow, 12, ChartStartRow + 15));

        // 選定資料範圍類別的資料來源
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, xAxisCol, xAxisCol)
        );

        // 選定資料範圍數值的資料來源
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, yAxisCol, yAxisCol)
        );

        //設定圖表類型為圓餅圖
        XDDFChartData data = chart.createData(ChartTypes.PIE, null , null);
        data.addSeries(categories, values);

        // 設定table title
        chart.setTitleText(valueTitle);
        chart.setTitleOverlay(false);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT);

        // 顯示圖表圖例
        CTDLbls dLbls = chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewDLbls();
        dLbls.addNewShowCatName().setVal(true);     // 顯示類別名稱
        dLbls.addNewShowVal().setVal(false);        // 不顯示值
        dLbls.addNewShowSerName().setVal(false);    // 不顯示數列名稱
        dLbls.addNewShowPercent().setVal(true);     // 不顯示百分比
        dLbls.addNewShowLeaderLines().setVal(true); // 顯示引導線

        data.setVaryColors(true);

        // 顯示圖表
        chart.plot(data);

        //System.out.println(dataStartRow+","+dataLastRow+","+dataCol+","+valueCol);
    }

    /**
     * 產生雷達圖
     */
    public void genRadarChart(Sheet sheet, int headerRow , int dataStartRow, int dataLastRow,
                              int xAxisCol, int yAxisCol, int ChartStartRow) {
        //System.out.println(dataStartRow + "," + dataLastRow + "," + dataCol + "," + yAxisCol);

        // 創建圖表
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();

        // 設定圖表位置
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 5, ChartStartRow, 12, ChartStartRow + 15));

        // 從headerRow中取得類別軸和數值軸的標題名稱
        Row row = sheet.getRow(headerRow);
        String categoryTitle = row.getCell(xAxisCol).getStringCellValue();
        String valueTitle = row.getCell(yAxisCol).getStringCellValue();

        // 數值軸的名稱來當整張圖表的標題
        chart.setTitleText(valueTitle);
        chart.setTitleOverlay(false);

        // 選定資料範圍類別的資料來源
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, xAxisCol, xAxisCol)
        );

        // 選定資料範圍數值的資料來源
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, yAxisCol, yAxisCol)
        );

        // 設定類別軸和數值軸
        XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        categoryAxis.setTitle(categoryTitle);
        categoryAxis.setVisible(true);

        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
        valueAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
        valueAxis.setVisible(true);

        // 設定圖表類型為雷達圖
        XDDFRadarChartData data = (XDDFRadarChartData) chart.createData(ChartTypes.RADAR, categoryAxis, valueAxis);
        XDDFRadarChartData.Series series = (XDDFRadarChartData.Series) data.addSeries(categories, values);

        // 設定雷達圖樣式為填滿
        data.setVaryColors(false);
        data.setStyle(RadarStyle.FILLED);

        // 設定填充顏色為橘色
        XDDFSolidFillProperties fillProperties = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.ORANGE));
        series.setFillProperties(fillProperties);

        // 顯示主要格線
        chart.getCTChart().getPlotArea().getCatAxArray(0).addNewMajorGridlines();
        chart.getCTChart().getPlotArea().getValAxArray(0).addNewMajorGridlines();

        // 顯示圖表
        chart.plot(data);
    }

    /**
     * 產生直條圖
     */
    public void genBarChart(Sheet sheet, int headerRow , int dataStartRow, int dataLastRow,
                            int xAxisCol, int yAxisCol, int ChartStartRow) {
        //System.out.println(dataStartRow+","+dataLastRow+","+xAxisCol+","+yAxisCol+","+indexRow+","+headerRow);

        // 創建圖表
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();

        // 設定圖表位置
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 5, ChartStartRow, 12, ChartStartRow + 15));

        // 從headerRow中取得類別軸和數值軸的標題名稱
        Row row = sheet.getRow(headerRow);
        String categoryTitle = row.getCell(xAxisCol).getStringCellValue();
        String valueTitle = row.getCell(yAxisCol).getStringCellValue();

        // 數值軸的名稱來當整張圖表的標題
        chart.setTitleText(valueTitle);
        chart.setTitleOverlay(false);

        // 選定資料範圍類別的資料來源
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, xAxisCol, xAxisCol)
        );

        // 選定資料範圍數值的資料來源
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, yAxisCol, yAxisCol)
        );

        // 設定類別軸和數值軸
        XDDFCategoryAxis xAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        xAxis.setTitle(categoryTitle);

        XDDFValueAxis yAxis = chart.createValueAxis(AxisPosition.LEFT);
        yAxis.setTitle(valueTitle);

        // 設為3DBarChart (XDDF'Bar3DChartData'比XDDF'ChartData'有更多細節設定)
        XDDFBar3DChartData barChartData = (XDDFBar3DChartData) chart.createData(ChartTypes.BAR3D, xAxis, yAxis);
        // 可以改成直條圖(改方向)
        barChartData.setBarDirection(BarDirection.COL);

        // 設定資料(圖例)
        XDDFChartData.Series series = barChartData.addSeries(categories, values);
        series.setTitle(valueTitle, null);

        // 顯示圖表
        chart.plot(barChartData);
    }

    /**
     * 產生折線圖
     */
    public void genLineChart(Sheet sheet, int headerRow , int dataStartRow, int dataLastRow,
                             int xAxisCol, int yAxisCol, int ChartStartRow){
        // 創建圖表
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();

        // 設定圖表位置
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 5, ChartStartRow, 12, ChartStartRow + 15));

        // 從headerRow中取得類別軸和數值軸的標題名稱
        Row row = sheet.getRow(headerRow);
        String categoryTitle = row.getCell(xAxisCol).getStringCellValue();
        String valueTitle = row.getCell(yAxisCol).getStringCellValue();

        // 數值軸的名稱來當整張圖表的標題
        chart.setTitleText(valueTitle);
        chart.setTitleOverlay(false);

        // 選定資料範圍類別的資料來源
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, xAxisCol, xAxisCol)
        );

        // 選定資料範圍數值的資料來源
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, yAxisCol, yAxisCol)
        );

        // 設定類別軸和數值軸
        XDDFCategoryAxis xAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        xAxis.setTitle(categoryTitle);

        XDDFValueAxis yAxis = chart.createValueAxis(AxisPosition.LEFT);
        yAxis.setTitle(valueTitle);

        // 設定圖表類型為折線圖
        XDDFChartData data = chart.createData(ChartTypes.LINE, xAxis, yAxis);
        data.setVaryColors(true);

        // 設定資料(圖例)
        data.addSeries(categories, values);

        // 顯示圖表圖例
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT); // 圖表圖例顯示在右邊
        //legend.se

        // 顯示圖表
        chart.plot(data);
    }
}
