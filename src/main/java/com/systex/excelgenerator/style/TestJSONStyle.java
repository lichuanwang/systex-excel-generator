package com.systex.excelgenerator.style;

import com.fasterxml.jackson.databind.ObjectMapper;  // 來自Jackson的JSON處理器，將物件轉JSON字串或相反
import com.fasterxml.jackson.annotation.JsonInclude; // 指定JSON序列化哪些屬性可以被省略用
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.awt.Color;
import java.io.IOException;


// 攜帶 Excel 字體樣式的資料
class FontData {
    public String fontName;
    public short fontSize;
    public boolean bold;
    public String color;  // 使用 String 表示顏色 (ARGB)
    public boolean italic;

    public FontData() {}

    public FontData(String fontName, short fontSize, boolean bold, String color, boolean italic) {
        this.fontName = fontName;
        this.fontSize = fontSize;
        this.bold = bold;
        this.color = color;
        this.italic = italic;
    }
}

// 儲存格樣式
class CellStyleData {
    public HorizontalAlignment alignment; // 水平對齊
    public VerticalAlignment verticalAlignment; // 垂直對齊
    public boolean wrapText; // 自動換行

    public CellStyleData() {}
    // ObjectMapper 序列化和反序列化工具跳過null
    public CellStyleData(HorizontalAlignment alignment, VerticalAlignment verticalAlignment, boolean wrapText) {
        this.alignment = alignment;
        this.verticalAlignment = verticalAlignment;
        this.wrapText = wrapText;
    }
}

// 複製樣式、比較樣式
public class TestJSONStyle {
    private static final ObjectMapper mapper = new ObjectMapper().setSerializationInclusion(JsonInclude.Include.NON_NULL);

    public static Font cloneFontWithJson(XSSFWorkbook workbook, XSSFFont originalFont) {
        try {
            FontData fontData = new FontData(
                    originalFont.getFontName(),
                    originalFont.getFontHeightInPoints(),
                    originalFont.getBold(),
                    originalFont.getXSSFColor().getARGBHex(), // 使用 ARGBHex 字串表示顏色
                    originalFont.getItalic()
            );
            // 將 fontData 物件序列化為 JSON
            String json = mapper.writeValueAsString(fontData);
            // 將 JSON 字串反序列化為一個新的 FontData 物件 clonedFontData
            FontData clonedFontData = mapper.readValue(json, FontData.class);

            XSSFFont newFont = workbook.createFont();
            newFont.setFontName(clonedFontData.fontName);
            newFont.setFontHeightInPoints(clonedFontData.fontSize);
            newFont.setBold(clonedFontData.bold);
            newFont.setItalic(clonedFontData.italic);

            // 設定顏色
            if (clonedFontData.color != null) {
                String argb = clonedFontData.color;
                int alpha = Integer.parseInt(argb.substring(0, 2), 16);
                int red = Integer.parseInt(argb.substring(2, 4), 16);
                int green = Integer.parseInt(argb.substring(4, 6), 16);
                int blue = Integer.parseInt(argb.substring(6, 8), 16);

                java.awt.Color awtColor = new java.awt.Color(red, green, blue, alpha);
                newFont.setColor(new XSSFColor(awtColor, null));
            }

            return newFont;
        } catch (IOException e) {
            throw new RuntimeException("Failed to clone font using JSON serialization", e);
        }
    }

    public static CellStyle cloneCellStyleWithJson(XSSFWorkbook workbook, CellStyle originalStyle) {
        try {
            CellStyleData styleData = new CellStyleData(
                    originalStyle.getAlignment(),
                    originalStyle.getVerticalAlignment(),
                    originalStyle.getWrapText()
            );

            String json = mapper.writeValueAsString(styleData);
            CellStyleData clonedStyleData = mapper.readValue(json, CellStyleData.class);

            CellStyle newStyle = workbook.createCellStyle();
            newStyle.setAlignment(clonedStyleData.alignment);
            newStyle.setVerticalAlignment(clonedStyleData.verticalAlignment);
            newStyle.setWrapText(clonedStyleData.wrapText);

            Font originalFont = workbook.getFontAt(originalStyle.getFontIndex());
            Font clonedFont = cloneFontWithJson(workbook, (XSSFFont) originalFont);
            newStyle.setFont(clonedFont);

            return newStyle;
        } catch (IOException e) {
            throw new RuntimeException("Failed to clone cell style using JSON serialization", e);
        }
    }

    public static boolean areStylesEqual(CellStyle style1, CellStyle style2, Workbook workbook) {
        if (style1.getAlignment() != style2.getAlignment()) return false;
        if (style1.getVerticalAlignment() != style2.getVerticalAlignment()) return false;
        if (style1.getWrapText() != style2.getWrapText()) return false;

        Font font1 = workbook.getFontAt(style1.getFontIndex());
        Font font2 = workbook.getFontAt(style2.getFontIndex());
        return areFontsEqual(font1, font2);
    }

    public static boolean areFontsEqual(Font font1, Font font2) {
        if (!font1.getFontName().equals(font2.getFontName())) return false;
        if (font1.getFontHeightInPoints() != font2.getFontHeightInPoints()) return false;
        if (font1.getBold() != font2.getBold()) return false;
        if (font1.getItalic() != font2.getItalic()) return false;

        if (font1 instanceof XSSFFont && font2 instanceof XSSFFont) {
            XSSFColor color1 = ((XSSFFont) font1).getXSSFColor();
            XSSFColor color2 = ((XSSFFont) font2).getXSSFColor();
            return color1 != null && color2 != null && color1.getARGBHex().equals(color2.getARGBHex());
        }
        return true;
    }

    public static void main(String[] args) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        CellStyle originalStyle = workbook.createCellStyle();
        XSSFFont originalFont = (XSSFFont) workbook.createFont();
        originalFont.setFontName("Arial");
        originalFont.setBold(true);
        originalFont.setColor(new XSSFColor(new Color(0, 128, 0), null));
        originalStyle.setFont(originalFont);

        // 使用 JSON 序列化進行深拷貝
        CellStyle clonedStyle = cloneCellStyleWithJson(workbook, originalStyle);
        Font clonedFont = workbook.getFontAt(clonedStyle.getFontIndex());

        // 比較新舊 CellStyle 物件
        System.out.println("Original and cloned styles have the same properties: " + areStylesEqual(originalStyle, clonedStyle, workbook));
        System.out.println("Original and cloned fonts have the same properties: " + areFontsEqual(originalFont, clonedFont));
    }
}