package com.systex.excelgenerator.style;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;

import java.io.FileOutputStream;

public class Test1 {
    public static void main(String[] args) {
        try (Workbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = (XSSFSheet) workbook.createSheet("Data");

            // 填充示例數據
            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue("Category");
            row.createCell(1).setCellValue("Value");

            row = sheet.createRow(1);
            row.createCell(0).setCellValue("A");
            row.createCell(1).setCellValue(5);

            row = sheet.createRow(2);
            row.createCell(0).setCellValue("B");
            row.createCell(1).setCellValue(10);

            row = sheet.createRow(3);
            row.createCell(0).setCellValue("C");
            row.createCell(1).setCellValue(15);

            // 創建圖表
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 10, 20);
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Sample Bar Chart");
            chart.setTitleOverlay(false);


            // 創建類別軸（X 軸）
            XDDFCategoryAxis categoryAxis = chart.createCategoryAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.BOTTOM);
            categoryAxis.setTitle("Category");

            // 創建數值軸（Y 軸）
            XDDFValueAxis valueAxis = chart.createValueAxis(org.apache.poi.xddf.usermodel.chart.AxisPosition.LEFT);
            valueAxis.setTitle("Value");
            valueAxis.setCrosses(org.apache.poi.xddf.usermodel.chart.AxisCrosses.AUTO_ZERO);

            // 定義圖表的數據源
            XDDFCategoryDataSource categories = XDDFDataSourcesFactory.fromStringCellRange(sheet, new org.apache.poi.ss.util.CellRangeAddress(1, 3, 0, 0));
            XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new org.apache.poi.ss.util.CellRangeAddress(1, 3, 1, 1));

            // 使用數據源創建長條圖的數據集
            XDDFChartData data = chart.createData(org.apache.poi.xddf.usermodel.chart.ChartTypes.BAR, categoryAxis, valueAxis);
            XDDFChartData.Series series = data.addSeries(categories, values);
            series.setTitle("Sample Data", null);
            chart.plot(data);

            // 將工作簿寫入文件
            try (FileOutputStream fileOut = new FileOutputStream("BarChartExample.xlsx")) {
                workbook.write(fileOut);
            }
            System.out.println("Bar chart created successfully.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
