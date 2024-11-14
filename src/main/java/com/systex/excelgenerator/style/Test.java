package com.systex.excelgenerator.style;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.gson.Gson;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.apache.poi.ss.usermodel.Color;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;

import java.io.*;
import java.util.Objects;


public class Test {
        public static void main(String[] args) {
            try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                // 創建一個 XSSFCellStyle
                XSSFCellStyle cellStyle = workbook.createCellStyle();

                // 設置一些樣式屬性
                cellStyle.setWrapText(true);
                cellStyle.setAlignment(HorizontalAlignment.CENTER);

                // 嘗試序列化
                try (ObjectOutputStream oos = new ObjectOutputStream(new FileOutputStream("cellStyle.ser"))) {
                    oos.writeObject(cellStyle);
                }

                System.out.println("序列化成功！");
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

//    public static class SS {
//        private String color;
//
//        public String getColor() {
//            return color;
//        }
//
//        @Override
//        public boolean equals(Object o) {
//            if (this == o) return true;
//            if (o == null || getClass() != o.getClass()) return false;
//            SS ss = (SS) o;
//            return Objects.equals(color, ss.color);
//        }
//
//        @Override
//        public int hashCode() {
//            return Objects.hashCode(color);
//        }
//
//        @Override
//        public String toString() {
//            return "SS{" +
//                    "color='" + color + '\'' +
//                    '}';
//        }
//
//        public void setColor(String color) {
//            this.color = color;
//        }
//    }
//
//
//    public static void main(String[] args) throws JsonProcessingException {
//        SS ss = new SS();
//        ss.setColor("#123456");
//
//        new String();
//
//        ObjectMapper mapper = new ObjectMapper();
//        String s = mapper.writeValueAsString(ss);
//
//        SS ss1 = mapper.readValue(s, SS.class);
//
//        System.out.println(ss1);
//        System.out.println(ss.equals(ss1));
//    }

//    public static void main(String[] args) {
//        try (XSSFWorkbook workbook = new XSSFWorkbook();) {
//            XSSFCellStyle cellStyle = workbook.createCellStyle();
//            byte[] rgb = new byte[3];
//            rgb[0] = (byte) 255;
//            rgb[1] = (byte) 255;
//            rgb[2] = (byte) 255;
//            cellStyle.setBorderColor(XSSFCellBorder.BorderSide.HORIZONTAL, new XSSFColor(rgb));
//
////            ObjectMapper mapper = new ObjectMapper();
////            String serializedCellStyle = mapper.writeValueAsString(cellStyle);
////
////            XSSFCellStyle clonedCellStyle = mapper.readValue(serializedCellStyle, XSSFCellStyle.class);
////            boolean equals = cellStyle.equals(clonedCellStyle);
//
////            NewStyle style = new NewStyle(cellStyle);
////            ByteArrayOutputStream baos = new ByteArrayOutputStream();
////
////            try (ObjectOutputStream oos = new ObjectOutputStream(baos);) {
////                oos.writeObject(style);
////            }
////
////            NewStyle clonedStyle;
////            try (ObjectInputStream ois = new ObjectInputStream(new ByteArrayInputStream(baos.toByteArray()))) {
////                clonedStyle = (NewStyle) ois.readObject();
////            }
////
////            boolean equals = style.equals(clonedStyle);
//
//            Gson mapper = new Gson();
//            String serializedCellStyle = mapper.toJson(cellStyle);
//
//            XSSFCellStyle clonedCellStyle = mapper.fromJson(serializedCellStyle, XSSFCellStyle.class);
//            boolean equals = cellStyle.equals(clonedCellStyle);
//
//            System.out.println(equals);
//
//        } catch (Exception e) {
//            throw new RuntimeException(e);
//        }
//    }

