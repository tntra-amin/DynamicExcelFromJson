package com.example.DynamicExcel.DynamicExcel.POC;

import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;


import java.io.*;
import java.util.*;
import java.util.Map;

public class DynamicExcelService {
    static List<String> columnsIncludedInExcel = List.of(
            "id",
            "created_at",
            "in_date",
            "out_date",
            "in_time",
            "out_time",
            "in_type",
            "out_type",
            "is_manual",
            "employee_employeejobdetails_ptr_id",
            "employee_first_name",
            "employee_middle_name",
            "employee_last_name",
            "employee_gender",
            "employee_mobile",
            "manualStatus_id",
            "manualStatus_reason",
            "manualStatus_status",
            "manualStatus_role",
            "manualStatus_action_date",
            "manualStatus_hr_action_date",
            "manualStatus_action_by_id",
            "_active");


    public static void main(String[] args) throws IOException, ParseException {

        JSONParser jsonParser = new JSONParser();

        Object jsonObject = jsonParser.parse(new FileReader("src/main/java/com/example/DynamicExcel/DynamicExcel/POC/SampleJson/json.json"));

        List<Map<String, Object>> values = new ArrayList<>();
        Set<String> allValidColumns = new HashSet<>();

        ExcelGeneratorUtil.generateMapAndSet(values, allValidColumns, jsonObject);

        Map<String, Integer> excelColumns = new HashMap<>();

        int[] ind = {0};
        columnsIncludedInExcel.forEach(column -> {
            if (allValidColumns.contains(column))
                excelColumns.put(column, ind[0]++);
        });

        ExcelGeneratorUtil.generateExcel(excelColumns, values);

    }

}
