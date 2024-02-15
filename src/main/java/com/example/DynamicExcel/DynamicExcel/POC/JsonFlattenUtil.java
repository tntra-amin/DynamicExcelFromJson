package com.example.DynamicExcel.DynamicExcel.POC;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

import java.util.LinkedHashMap;
import java.util.Map;

public class JsonFlattenUtil {

    public static Map<String, Object> flattenUtil(Object input) {
        Map<String, Object> flattenedMap = new LinkedHashMap<>();
        String parentKey = "";
        flattenObjectOrArray(input, flattenedMap, parentKey);
        return flattenedMap;
    }

    private static void flattenObjectOrArray(Object input, Map<String, Object> flattenedMap, String parentKey) {
        if (input instanceof JSONObject) {
            flattenObject((JSONObject) input, flattenedMap, parentKey);
        } else if (input instanceof JSONArray) {
            flattenArray((JSONArray) input, flattenedMap, parentKey);
        } else {
            flattenedMap.put(parentKey, input);
        }
    }

    private static void flattenObject(JSONObject inputObject, Map<String, Object> flattenedMap, String parentKey) {
        for (Object key : inputObject.keySet()) {
            String newKey = parentKey.isEmpty() ? (String) key : parentKey + "_" + key;
            Object value = inputObject.get(key);
            flattenObjectOrArray(value, flattenedMap, newKey);
        }
    }

    private static void flattenArray(JSONArray inputArray, Map<String, Object> flattenedMap, String parentKey) {
        for (int i = 0; i < inputArray.size(); i++) {
            Object inputObj = inputArray.get(i);
            String newKey = parentKey.isEmpty() ? "" : parentKey + "_" + (i+1);
            flattenObjectOrArray(inputObj, flattenedMap, newKey);
        }
    }

}
