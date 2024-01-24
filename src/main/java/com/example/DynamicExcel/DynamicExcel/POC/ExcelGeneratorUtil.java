package com.example.DynamicExcel.DynamicExcel.POC;

import org.apache.poi.xssf.usermodel.*;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Set;

@RestController
public class ExcelGeneratorUtil {

    public static void generateMapAndSet(List<Map<String, Object>> values, Set<String> allValidColumns, Object jsonObject) {

        if (jsonObject instanceof JSONArray) {

            ((JSONArray) jsonObject).forEach(jsonObj -> {

                values.add(JsonFlattenUtil.flattenUtil(jsonObj, allValidColumns));

            });

        } else if (jsonObject instanceof JSONObject) {

            values.add(JsonFlattenUtil.flattenUtil(jsonObject, allValidColumns));

        }

    }

    public static void generateExcel(Map<String, Integer> excelColumns, List<Map<String, Object>> values) {

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

            int rowIndex = 0;
            row = sheet.createRow(rowIndex);

            //Setting values for headers
            for (Map.Entry<String, Integer> header : excelColumns.entrySet()) {

                cell = row.createCell(header.getValue());

                String headerName = Arrays.stream(header.getKey().split("_")).reduce("",(str1, str2) -> str1 + " " + str2).toUpperCase();

                cell.setCellValue(headerName);
                cell.setCellStyle(boldCellStyle);
                sheet.autoSizeColumn(header.getValue());
            }

            rowIndex++;


            for (Map<String, Object> valueMap : values) {

                row = sheet.createRow(rowIndex);

                for (Map.Entry<String, Object> currentValue : valueMap.entrySet()) {

                    if (excelColumns.containsKey(currentValue.getKey())) {

                        cell = row.createCell(excelColumns.get(currentValue.getKey()));
                        if (currentValue.getValue() instanceof Integer)
                            cell.setCellValue((Integer) currentValue.getValue());
                        else if(currentValue.getValue() instanceof String)
                            cell.setCellValue((String) currentValue.getValue());
                        else if(currentValue.getValue() instanceof Long)
                            cell.setCellValue((Long) currentValue.getValue());
                        else if(currentValue.getValue() instanceof Boolean)
                            cell.setCellValue((Boolean) currentValue.getValue());
                        else
                            cell.setCellValue((String) currentValue.getValue());

                    }

                }

                rowIndex++;

            }
            ;

            FileOutputStream out = new FileOutputStream(new File("SAMPLE_TEST.xlsx"));
            workbook.write(out);
            workbook.close();
            out.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
