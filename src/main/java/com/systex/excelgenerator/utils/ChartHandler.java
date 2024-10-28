package com.systex.excelgenerator.utils;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ChartHandler {

    // every chart
    // pie chart
    public void genPieChart(){
        // input data : 資料範圍

        // 共用workbook , sheet
        Workbook wb = new XSSFWorkbook();
        // XSSF , HSSF...會影響到後續畫圖表的sheet
        XSSFSheet sheet = (XSSFSheet) wb.createSheet();

        // 創建圖表
        XSSFDrawing drawing = sheet.createDrawingPatriarch();

        // 圖表位置設定(要固定設定在資料的最下方嗎..?)
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 3, 0, 10, 15));

        //類別資料來源
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 3, 0, 0));

        //數值資料來源
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 3, 1, 1));

        //設定圖表類型為圓餅圖
        XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);

        // 設定圖表中的每個顏色是否要不同
        data.setVaryColors(true);

        // apply categories and values
        data.addSeries(categories, values);
        chart.plot(data);
    }

    // Rader chart
    public void genRadarChart(){
        // input data : 資料範圍

        // 共用workbook , sheet
        Workbook wb = new XSSFWorkbook();
        // XSSF , HSSF...會影響到後續畫圖表的sheet
        XSSFSheet sheet = (XSSFSheet) wb.createSheet();

        // 創建圖表
        XSSFDrawing drawing = sheet.createDrawingPatriarch();

        // 圖表位置設定
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 3, 0, 10, 15));

        //類別資料來源
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 3, 0, 0));

        //數值資料來源
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 3, 1, 1));

        //設定圖表類型為雷達圖
        XDDFChartData data = chart.createData(ChartTypes.RADAR, null, null);

        // 設定圖表中的每個顏色是否要不同
        data.setVaryColors(true);

        // apply categories and values
        data.addSeries(categories, values);
        chart.plot(data);
    }

    // Bar chart
    public void genBarChart(){
        // input data : 資料範圍

        // 共用workbook , sheet
        Workbook wb = new XSSFWorkbook();
        // XSSF , HSSF...會影響到後續畫圖表的sheet
        XSSFSheet sheet = (XSSFSheet) wb.createSheet();

        // 創建圖表
        XSSFDrawing drawing = sheet.createDrawingPatriarch();

        // 圖表位置設定
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 3, 0, 10, 15));

        //類別資料來源
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 3, 0, 0));

        //數值資料來源
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 3, 1, 1));

        //設定圖表類型為長條圖
        XDDFChartData data = chart.createData(ChartTypes.BAR, null, null);

        // 設定圖表中的每個顏色是否要不同
        data.setVaryColors(true);

        // apply categories and values
        data.addSeries(categories, values);
        chart.plot(data);
    }

    // Line chart
    public void genLineChart(){
        // input data : 資料範圍

        // 共用workbook , sheet
        Workbook wb = new XSSFWorkbook();
        // XSSF , HSSF...會影響到後續畫圖表的sheet
        XSSFSheet sheet = (XSSFSheet) wb.createSheet();

        // 創建圖表
        XSSFDrawing drawing = sheet.createDrawingPatriarch();

        // 圖表位置設定
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 3, 0, 10, 15));

        //類別資料來源
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(1, 3, 0, 0));

        //數值資料來源
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 3, 1, 1));

        //設定圖表類型為折線圖
        XDDFChartData data = chart.createData(ChartTypes.LINE, null, null);

        // 設定圖表中的每個顏色是否要不同
        data.setVaryColors(true);

        // apply categories and values
        data.addSeries(categories, values);
        chart.plot(data);
    }
}
