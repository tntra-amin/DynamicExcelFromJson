package com.example.DynamicExcel.DynamicExcel.POC;

import org.apache.poi.xssf.usermodel.*;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

@RestController
public class ExcelGeneratorUtil {

    public static void generateMapAndSet(List<Map<String, Object>> values, Object jsonObject) {

        if (jsonObject instanceof JSONArray) {

            ((JSONArray) jsonObject).forEach(jsonObj -> {

                values.add(JsonFlattenUtil.flattenUtil(jsonObj));

            });

        } else if (jsonObject instanceof JSONObject) {

            values.add(JsonFlattenUtil.flattenUtil(jsonObject));

        }

    }

    public static List<Map<String, Object>> indexTheMapUtil(List<Map<String, Object>> values){
        List<Map<String, Object>> modifiedList = new ArrayList<>();

        for (int i = 0; i < values.size(); i++) {
            Map<String, Object> originalMap = values.get(i);
            Map<String, Object> modifiedMap = new HashMap<>();

            int index = 0;
            for (Map.Entry<String, Object> entry : originalMap.entrySet()) {
                String indexedKey = index++ + ":" + entry.getKey();
                modifiedMap.put(indexedKey, entry.getValue());
            }

            modifiedList.add(modifiedMap);
        }

        return modifiedList;
    }

    public static void generateExcel(List<Map<String, Object>> values) {

        try {

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet;
            XSSFRow row;
            XSSFCell cell;

            XSSFFont boldFont = workbook.createFont();
            boldFont.setFontName("Calibri");
            boldFont.setFontHeightInPoints((short) 11);
            boldFont.setBold(true);

            XSSFCellStyle boldCellStyle = workbook.createCellStyle();
            boldCellStyle.setFont(boldFont);

            sheet = workbook.createSheet("TRIAL SHEET");

            XSSFRow headerRow = sheet.createRow(0);

            int rowIndex = 1;

            for (Map<String, Object> valueMap : values) {

                row = sheet.createRow(rowIndex);

                for (Map.Entry<String, Object> currentValue : valueMap.entrySet()) {

                    String key = Arrays.stream(currentValue.getKey().split(":")[1].split("_")).reduce("", (str1, str2) -> str1 + " " + str2).toUpperCase();

                    int columnIndex = getColumnIndexUtil(headerRow,key)==-1 ? Integer.parseInt(currentValue.getKey().split(":")[0]) : getColumnIndexUtil(headerRow,key);

                    if(headerRow.getCell(columnIndex) == null){
                        headerRow.createCell(columnIndex).setCellValue(key);
                    }
                    if(!headerRow.getCell(columnIndex).getStringCellValue().equals(key)){
                        continue;
                    }

                    cell = row.createCell(columnIndex);
                    sheet.autoSizeColumn(columnIndex);

                    if (currentValue.getValue() instanceof Integer)
                        cell.setCellValue((Integer) currentValue.getValue());
                    else if (currentValue.getValue() instanceof String)
                        cell.setCellValue((String) currentValue.getValue());
                    else if (currentValue.getValue() instanceof Long)
                        cell.setCellValue((Long) currentValue.getValue());
                    else if (currentValue.getValue() instanceof Boolean)
                        cell.setCellValue((Boolean) currentValue.getValue());
                    else
                        cell.setCellValue((String) currentValue.getValue());

                }

                rowIndex++;

            }

            FileOutputStream out = new FileOutputStream(new File("SAMPLE_TEST.xlsx"));
            workbook.write(out);
            workbook.close();
            out.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static int getColumnIndexUtil(XSSFRow headerRow,String key){
        for(int i=0;i<headerRow.getPhysicalNumberOfCells();i++){
            if(headerRow.getCell(i) != null && headerRow.getCell(i).getStringCellValue().equals(key))
                return i;
        }
        return -1;
    }
}
