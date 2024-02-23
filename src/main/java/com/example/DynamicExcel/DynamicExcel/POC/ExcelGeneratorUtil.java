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

    public static int maxLength = Integer.MIN_VALUE;
    public static Map<String,Object> headers;
    public static Map<String,Integer> headersIndexMap = new HashMap<>();

    public static void generateMapAndSet(List<Map<String, Object>> values, Object jsonObject) {

        if (jsonObject instanceof JSONArray) {

            ((JSONArray) jsonObject).forEach(jsonObj -> {

                Map<String,Object> current = JsonFlattenUtil.flattenUtil(jsonObj);

                if(current.size() > maxLength){
                    maxLength = current.size();
                    headers = current;
                }

                values.add(current);

            });

        } else if (jsonObject instanceof JSONObject) {

            Map<String,Object> current = JsonFlattenUtil.flattenUtil(jsonObject);

            if(current.size() > maxLength){
                maxLength = current.size();
                headers = current;
            }

            values.add(current);

        }

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

            int colIndex = 0;

            for(String col : headers.keySet()){
                String key = Arrays.stream(col.split("_")).reduce("", (str1, str2) -> str1 + " " + str2).toUpperCase();
                cell = headerRow.createCell(colIndex);
                headersIndexMap.put(key,colIndex);
                cell.setCellValue(key);
                colIndex++;
            }

            int rowIndex = 1;

            for (Map<String, Object> valueMap : values) {

                row = sheet.createRow(rowIndex);

                for (Map.Entry<String, Object> currentValue : valueMap.entrySet()) {

                    String key = Arrays.stream(currentValue.getKey().split("_")).reduce("", (str1, str2) -> str1 + " " + str2).toUpperCase();

                    int columnIndex = headersIndexMap.get(key);

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

}
