package com.example.DynamicExcel.DynamicExcel.POC;

import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;


import java.io.*;
import java.util.*;
import java.util.Map;

public class DynamicExcelService {

    public static void main(String[] args) throws IOException, ParseException {

        JSONParser jsonParser = new JSONParser();

        Object jsonObject = jsonParser.parse(new FileReader("src/main/java/com/example/DynamicExcel/DynamicExcel/POC/SampleJson/json.json"));

        List<Map<String, Object>> values = new ArrayList<>();

        ExcelGeneratorUtil.generateMapAndSet(values, jsonObject);

        ExcelGeneratorUtil.generateExcel(values);

    }

}
