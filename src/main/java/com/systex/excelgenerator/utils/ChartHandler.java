package com.systex.excelgenerator.utils;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.PresetColor;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFLineProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;

public class ChartHandler {

    // every chart
    // pie chart
    public void genPieChart(Sheet sheet, int dataStartRow, int dataLastRow,
                            int categoryCol, int valueCol, int indexRow, int headerRow){
        //System.out.println(dataStartRow+","+dataLastRow+","+dataCol+","+valueCol);

        // 創建圖表
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();

        // 從 headerRow 中取得 X 軸和 Y 軸的標題名稱
        Row row = sheet.getRow(headerRow);
        String categoryTitle = row.getCell(categoryCol).getStringCellValue(); // 取得 Name 的標題
        String valueTitle = row.getCell(valueCol).getStringCellValue();       // 取得 Level 的標題

        // 設置圖表的位置在 indexRow 開始並向下延伸
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 3, indexRow, 10, indexRow + 15));

        // 類別資料來源
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, categoryCol, categoryCol)
        );

        // 數值資料來源
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, valueCol, valueCol)
        );

        //設定圖表類型為圓餅圖
        XDDFChartData data = chart.createData(ChartTypes.PIE, null , null);
        data.addSeries(categories, values);
    
        // 設定table title
        chart.setTitleText(valueTitle);
        chart.setTitleOverlay(false);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT);

        CTDLbls dLbls = chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewDLbls();
        dLbls.addNewShowCatName().setVal(true);     // 顯示類別名稱
        dLbls.addNewShowVal().setVal(false);        // 不顯示值
        dLbls.addNewShowSerName().setVal(false);    // 不顯示數列名稱
        dLbls.addNewShowPercent().setVal(true);     // 不顯示百分比
        dLbls.addNewShowLeaderLines().setVal(true); // 顯示引導線

        data.setVaryColors(true);
        chart.plot(data);

        //System.out.println(dataStartRow+","+dataLastRow+","+dataCol+","+valueCol);
    }

    // Rader chart
    public void genRadarChart(Sheet sheet, int dataStartRow, int dataLastRow,
                              int categoryCol, int valueCol, int indexRow, int headerRow) {
        //System.out.println(dataStartRow + "," + dataLastRow + "," + dataCol + "," + valueCol);
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();

        // 從 headerRow 中取得 X 軸和 Y 軸的標題名稱
        Row row = sheet.getRow(headerRow);
        String categoryTitle = row.getCell(categoryCol).getStringCellValue(); // 取得 Name 的標題
        String valueTitle = row.getCell(valueCol).getStringCellValue();       // 取得 Level 的標題

        // 設置圖表位置
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 3, indexRow, 10, indexRow + 15));
        chart.setTitleText(valueTitle);
        chart.setTitleOverlay(false);

        // 設定類別和數值資料來源
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, categoryCol, categoryCol)
        );
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, valueCol, valueCol)
        );

        // 創建類別軸和數值軸
        XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        categoryAxis.setTitle(categoryTitle);
        categoryAxis.setVisible(true);

        XDDFValueAxis valueAxis = chart.createValueAxis(AxisPosition.LEFT);
        valueAxis.setTitle(valueTitle);
        valueAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
        valueAxis.setVisible(true);

        // 設定圖表類型為雷達圖
        XDDFRadarChartData data = (XDDFRadarChartData) chart.createData(ChartTypes.RADAR, categoryAxis, valueAxis);
        XDDFRadarChartData.Series series = (XDDFRadarChartData.Series) data.addSeries(categories, values);
        series.setTitle(valueTitle, null);

        // 設定雷達圖樣式為填滿
        data.setVaryColors(false);
        data.setStyle(RadarStyle.FILLED);

        // 設定填充顏色為橘色
        XDDFSolidFillProperties fillProperties = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.ORANGE));
        series.setFillProperties(fillProperties);

        // 繪製圖表
        chart.plot(data);

        // 自訂其他樣式如字型、軸顏色等
        chart.getCTChart().getPlotArea().getCatAxArray(0).addNewMajorGridlines(); // 顯示主要格線
        chart.getCTChart().getPlotArea().getValAxArray(0).addNewMajorGridlines(); // 顯示主要格線
    }


    // 要修改
    // 直條圖
    public void genBarChart(Sheet sheet, int dataStartRow, int dataLastRow,
                            int categoryCol, int valueCol, int indexRow, int headerRow) {
        //System.out.println(dataStartRow+","+dataLastRow+","+categoryCol+","+valueCol+","+indexRow+","+headerRow);

        // 創建圖表
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();

        // 設置圖表位置
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 3, indexRow, 10, indexRow + 15));

        // 從 headerRow 中取得 X 軸和 Y 軸的標題名稱
        Row row = sheet.getRow(headerRow);
        String categoryTitle = row.getCell(categoryCol).getStringCellValue(); // 取得 Name 的標題
        String valueTitle = row.getCell(valueCol).getStringCellValue();       // 取得 Level 的標題

        chart.setTitleText(valueTitle); // 使用 Y 軸的標題作為圖表標題
        chart.setTitleOverlay(false);

        // 類別資料來源
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, categoryCol, categoryCol)
        );

        // 數值資料來源
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, valueCol, valueCol)
        );

        // 設定 X 軸和 Y 軸
        XDDFCategoryAxis xAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        xAxis.setTitle(categoryTitle); // 設定 X 軸標題

        XDDFValueAxis yAxis = chart.createValueAxis(AxisPosition.LEFT);
        yAxis.setTitle(valueTitle); // 設定 Y 軸標題

        // 設為3DBarChart (XDDF'Bar3DChartData'比XDDF'ChartData'有更多細節設定)
        XDDFBar3DChartData barChartData = (XDDFBar3DChartData) chart.createData(ChartTypes.BAR3D, xAxis, yAxis);
        // 可以改成直條圖(改方向)
        barChartData.setBarDirection(BarDirection.COL);

        // 設定資料
        XDDFChartData.Series series = barChartData.addSeries(categories, values);
        series.setTitle(valueTitle, null);

        chart.plot(barChartData);
    }


    // Line chart
    public void genLineChart(Sheet sheet, int dataStartRow, int dataLastRow,
                             int categoryCol, int valueCol, int indexRow, int headerRow){
        // 創建圖表
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
        XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, 3, indexRow, 10, indexRow + 15));

        // 從 headerRow 中取得 X 軸和 Y 軸的標題名稱
        Row row = sheet.getRow(headerRow);
        String categoryTitle = row.getCell(categoryCol).getStringCellValue(); // 取得 Name 的標題
        String valueTitle = row.getCell(valueCol).getStringCellValue();       // 取得 Level 的標題

        chart.setTitleText(valueTitle); // 使用 Y 軸的標題作為圖表標題
        chart.setTitleOverlay(false);

        // 設定X軸和Y軸
        XDDFCategoryAxis xAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        xAxis.setTitle(categoryTitle);

        XDDFValueAxis yAxis = chart.createValueAxis(AxisPosition.LEFT);
        yAxis.setTitle(valueTitle);

        // 類別資料來源
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, categoryCol, categoryCol)
        );

        // 數值資料來源
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                (XSSFSheet) sheet, new CellRangeAddress(dataStartRow, dataLastRow, valueCol, valueCol)
        );

        // 設定圖表類型為折線圖
        XDDFChartData data = chart.createData(ChartTypes.LINE, xAxis, yAxis);
        data.setVaryColors(true);

        // apply categories and values
        data.addSeries(categories, values);
        chart.plot(data);
    }
}
