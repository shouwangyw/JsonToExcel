package com.yw;

public class JsonToExcel {
    public static void main(String[] args) {
        if (args == null || args.length < 1) {
            System.out.println("请指定json文件");
            return;
        }
        String jsonFilePath = args[0];
        String excelFilePath = args.length > 1 ? args[1] : jsonFilePath.replace("json", "xlsx");
        
        JsonToExcelConverter.convertJsonToExcel(jsonFilePath, excelFilePath);
    }
}