package com.systex.excelgenerator.style;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.annotation.JsonInclude;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import java.io.IOException;

public class ExcelStyle {
    private static final ObjectMapper mapper = new ObjectMapper().setSerializationInclusion(JsonInclude.Include.NON_NULL);

    // JSON 基礎字型資料類別
    static class FontData {
        public String fontName;
        public short fontSize;
        public boolean bold;
        public byte[] color; // 用 byte[] 表示 RGB 顏色
        public boolean italic;

        public FontData() {}

        public FontData(String fontName, short fontSize, boolean bold, byte[] color, boolean italic) {
            this.fontName = fontName;
            this.fontSize = fontSize;
            this.bold = bold;
            this.color = color;
            this.italic = italic;
        }
    }

    // 使用 JSON 複製字型
    public static Font cloneFont(XSSFWorkbook workbook, XSSFFont originalFont) {
        try {
            byte[] colorBytes = originalFont.getXSSFColor() != null ? originalFont.getXSSFColor().getRGB() : null;
            FontData fontData = new FontData(
                    originalFont.getFontName(),
                    originalFont.getFontHeightInPoints(),
                    originalFont.getBold(),
                    colorBytes,
                    originalFont.getItalic()
            );

            // JSON 序列化和反序列化
            FontData clonedFontData = mapper.readValue(mapper.writeValueAsString(fontData), FontData.class);

            XSSFFont newFont = workbook.createFont();
            newFont.setFontName(clonedFontData.fontName);
            newFont.setFontHeightInPoints(clonedFontData.fontSize);
            newFont.setBold(clonedFontData.bold);
            newFont.setItalic(clonedFontData.italic);

            // 設置字型顏色
            if (clonedFontData.color != null) {
                newFont.setColor(new XSSFColor(clonedFontData.color, null));
            }
            return newFont;
        } catch (IOException e) {
            throw new RuntimeException("複製字型時發生錯誤", e);
        }
    }

    // 使用 JSON 複製樣式
    public static CellStyle cloneStyle(XSSFWorkbook workbook, CellStyle originalStyle) {
        try {
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(originalStyle);

            Font originalFont = workbook.getFontAt(originalStyle.getFontIndex());
            Font clonedFont = cloneFont(workbook, (XSSFFont) originalFont);
            newStyle.setFont(clonedFont);

            return newStyle;
        } catch (Exception e) {
            throw new RuntimeException("複製樣式時發生錯誤", e);
        }
    }


}


//package com.systex.excelgenerator.style;
//
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//public class ExcelStyleUtils {
//
//    private CellStyle Style;
//    // 自訂 cloneFont(只有常見的)
//    public static Font cloneFont(XSSFWorkbook workbook, Font originalFont) {
//        Font newFont = workbook.createFont();
//        newFont.setBold(originalFont.getBold());
//        newFont.setFontHeightInPoints(originalFont.getFontHeightInPoints());
//        newFont.setFontName(originalFont.getFontName());
//        newFont.setColor(originalFont.getColor());
//        newFont.setUnderline(originalFont.getUnderline());
//        newFont.setItalic(originalFont.getItalic());
//        return newFont;
//    }
//
//    public static CellStyle cloneStyle(XSSFWorkbook workbook, CellStyle originalStyle) {
//        CellStyle newStyle = workbook.createCellStyle();
//        newStyle.cloneStyleFrom(originalStyle);
//
//        Font originalFont = workbook.getFontAt(originalStyle.getFontIndex());
//        Font clonedFont = cloneFont(workbook, originalFont);
//        newStyle.setFont(clonedFont);
//        return newStyle;
//    }
//    public static CellStyle createSpecialStyle(XSSFWorkbook workbook) {
//        CellStyle specialStyle = workbook.createCellStyle();
//
////        this.style = workbook.createCellStyle();
//        Font font = workbook.createFont();
//        font.setBold(true);
//        specialStyle.setFont(font);
//
//        specialStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
//        specialStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//
//        return specialStyle;
//    }
// }