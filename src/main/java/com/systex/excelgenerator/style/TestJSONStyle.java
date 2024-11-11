package com.systex.excelgenerator.style;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import java.io.IOException;

/**
 * 字型資料類別，用來儲存字型的基本屬性
 */
class FontData {
    public String fontName;
    public short fontSize;
    public boolean bold;

    public FontData() {}  // 無參數建構子

    public FontData(String fontName, short fontSize, boolean bold) {
        this.fontName = fontName;
        this.fontSize = fontSize;
        this.bold = bold;
    }
}

public class TestJSONStyle {
    private static final ObjectMapper mapper = new ObjectMapper();

    /**
     * 複製字型：將字型資料轉成 JSON 格式，再轉回字型物件
     */
    public static Font cloneFont(XSSFWorkbook workbook, XSSFFont originalFont) {
        try {
            // 將原字型的屬性存入 FontData
            FontData fontData = new FontData(
                    originalFont.getFontName(),
                    originalFont.getFontHeightInPoints(),
                    originalFont.getBold()
            );

            // JSON 序列化後再反序列化，模擬 "複製"
            FontData clonedFontData = mapper.readValue(mapper.writeValueAsString(fontData), FontData.class);

            // 使用複製的字型資料建立新字型
            XSSFFont newFont = workbook.createFont();
            newFont.setFontName(clonedFontData.fontName);
            newFont.setFontHeightInPoints(clonedFontData.fontSize);
            newFont.setBold(clonedFontData.bold);

            return newFont;
        } catch (IOException e) {
            throw new RuntimeException("複製字型時發生錯誤", e);
        }
    }

    public static void main(String[] args) {
        // 創建 Excel 工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 設定字型
        XSSFFont originalFont = workbook.createFont();
        originalFont.setFontName("Arial");
        originalFont.setBold(true);

        // 複製字型
        Font clonedFont = cloneFont(workbook, originalFont);

        // 檢查原字型與複製字型是否相同內容
        System.out.println("兩者有相同屬性: " + originalFont.getFontName().equals(((XSSFFont) clonedFont).getFontName()));
        System.out.println("兩者有相同記憶體位置: " + (originalFont.getBold() == ((XSSFFont) clonedFont).getBold()));
    }
}
