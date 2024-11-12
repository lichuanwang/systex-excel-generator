//package com.systex.excelgenerator.style;
//
//import com.fasterxml.jackson.databind.ObjectMapper;
//import com.fasterxml.jackson.annotation.JsonInclude;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.*;
//import java.io.IOException;
//
//// 攜帶 Excel 字體樣式的資料
//class FontsData {
//    public String fontName;
//    public short fontSize;
//    public boolean bold;
//    public byte[] color;  // 使用 byte[] 表示 RGB 顏色
//    public boolean italic;
//
//    public FontsData() {}
//
//    public FontsData(String fontName, short fontSize, boolean bold, byte[] color, boolean italic) {
//        this.fontName = fontName;
//        this.fontSize = fontSize;
//        this.bold = bold;
//        this.color = color;
//        this.italic = italic;
//    }
//}
//
//// 儲存格樣式
//class CellStyleData {
//    public HorizontalAlignment alignment;
//    public VerticalAlignment verticalAlignment;
//    public boolean wrapText;
//
//    public CellStyleData() {}
//
//    public CellStyleData(HorizontalAlignment alignment, VerticalAlignment verticalAlignment, boolean wrapText) {
//        this.alignment = alignment;
//        this.verticalAlignment = verticalAlignment;
//        this.wrapText = wrapText;
//    }
//}
//
//// 複製樣式、比較樣式
//public class TestJSONStyleColor {
//    private static final ObjectMapper mapper = new ObjectMapper().setSerializationInclusion(JsonInclude.Include.NON_NULL);
//
//    public static Font cloneFontWithJson(XSSFWorkbook workbook, XSSFFont originalFont) {
//        try {
//            // 獲取 RGB 顏色值
//            byte[] colorBytes = originalFont.getXSSFColor() != null ? originalFont.getXSSFColor().getRGB() : null;
//
//            FontsData fontsData = new FontsData(
//                    originalFont.getFontName(),
//                    originalFont.getFontHeightInPoints(),
//                    originalFont.getBold(),
//                    colorBytes,
//                    originalFont.getItalic()
//            );
//            // 序列化與反序列化
//            FontsData clonedFontsData = mapper.readValue(mapper.writeValueAsString(fontsData), FontsData.class);
//
//            XSSFFont newFont = workbook.createFont();
//            newFont.setFontName(clonedFontsData.fontName);
//            newFont.setFontHeightInPoints(clonedFontsData.fontSize);
//            newFont.setBold(clonedFontsData.bold);
//            newFont.setItalic(clonedFontsData.italic);
//
//            // 設定 RGB 顏色
//            if (clonedFontsData.color != null) {
//                XSSFColor xssfColor = new XSSFColor(clonedFontsData.color, null);
//                newFont.setColor(xssfColor);
//            }
//
//            return newFont;
//        } catch (IOException e) {
//            throw new RuntimeException("Failed to clone font using JSON serialization", e);
//        }
//    }
//
//    public static CellStyle cloneCellStyleWithJson(XSSFWorkbook workbook, CellStyle originalStyle) {
//        try {
//            CellStyleData styleData = new CellStyleData(
//                    originalStyle.getAlignment(),
//                    originalStyle.getVerticalAlignment(),
//                    originalStyle.getWrapText()
//            );
//            CellStyleData clonedStyleData = mapper.readValue(mapper.writeValueAsString(styleData), CellStyleData.class);
//
//            CellStyle newStyle = workbook.createCellStyle();
//            newStyle.setAlignment(clonedStyleData.alignment);
//            newStyle.setVerticalAlignment(clonedStyleData.verticalAlignment);
//            newStyle.setWrapText(clonedStyleData.wrapText);
//
//            Font originalFont = workbook.getFontAt(originalStyle.getFontIndex());
//            newStyle.setFont(cloneFontWithJson(workbook, (XSSFFont) originalFont));
//            return newStyle;
//        } catch (IOException e) {
//            throw new RuntimeException("Failed to clone cell style using JSON serialization", e);
//        }
//    }
//
//    public static boolean areStylesEqual(CellStyle style1, CellStyle style2, Workbook workbook) {
//        return style1.getAlignment() == style2.getAlignment() &&
//                style1.getVerticalAlignment() == style2.getVerticalAlignment() &&
//                style1.getWrapText() == style2.getWrapText() &&
//                areFontsEqual(workbook.getFontAt(style1.getFontIndex()), workbook.getFontAt(style2.getFontIndex()));
//    }
//
//    public static boolean areFontsEqual(Font font1, Font font2) {
//        if (!(font1 instanceof XSSFFont) || !(font2 instanceof XSSFFont)) return true;
//
//        XSSFFont xssfFont1 = (XSSFFont) font1;
//        XSSFFont xssfFont2 = (XSSFFont) font2;
//
//        boolean colorsEqual = false;
//        if (xssfFont1.getXSSFColor() == null && xssfFont2.getXSSFColor() == null) {
//            colorsEqual = true;
//        } else if (xssfFont1.getXSSFColor() != null && xssfFont2.getXSSFColor() != null) {
//            colorsEqual = java.util.Arrays.equals(xssfFont1.getXSSFColor().getRGB(), xssfFont2.getXSSFColor().getRGB());
//        }
//
//        return xssfFont1.getFontName().equals(xssfFont2.getFontName()) &&
//                xssfFont1.getFontHeightInPoints() == xssfFont2.getFontHeightInPoints() &&
//                xssfFont1.getBold() == xssfFont2.getBold() &&
//                xssfFont1.getItalic() == xssfFont2.getItalic() &&
//                colorsEqual;
//    }
//
//    public static void main(String[] args) {
//        XSSFWorkbook workbook = new XSSFWorkbook();
//        CellStyle originalStyle = workbook.createCellStyle();
//        XSSFFont originalFont = (XSSFFont) workbook.createFont();
//        originalFont.setFontName("Arial");
//        originalFont.setBold(true);
//        originalFont.setColor(new XSSFColor(new java.awt.Color(0, 128, 0), null));  // 設定為綠色
//        originalStyle.setFont(originalFont);
//
//        // 使用 JSON 序列化進行深拷貝
//        CellStyle clonedStyle = cloneCellStyleWithJson(workbook, originalStyle);
//
//        // 比較新舊 CellStyle 物件
//        System.out.println("兩者有相同屬性: " + areStylesEqual(originalStyle, clonedStyle, workbook));
//        System.out.println("兩者有相同記憶體位置: " + (originalStyle == clonedStyle));
//    }
//}
