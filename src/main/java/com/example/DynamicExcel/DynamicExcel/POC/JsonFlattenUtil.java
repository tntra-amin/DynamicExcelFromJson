package com.example.DynamicExcel.DynamicExcel.POC;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;

public class JsonFlattenUtil {

    public static Map<String, Object> flattenUtil(Object input, Set<String> allValidColumns) {
        Map<String, Object> flattenedMap = new HashMap<>();
        String parentKey = "";
        flattenObjectOrArray(input, flattenedMap, parentKey, allValidColumns);
        return flattenedMap;
    }

    private static void flattenObjectOrArray(Object input, Map<String, Object> flattenedMap, String parentKey, Set<String> allValidColumns) {
        if (input instanceof JSONObject) {
            flattenObject((JSONObject) input, flattenedMap, parentKey, allValidColumns);
        } else if (input instanceof JSONArray) {
            flattenArray((JSONArray) input, flattenedMap, parentKey, allValidColumns);
        } else {
            allValidColumns.add(parentKey);
            flattenedMap.put(parentKey, input);
        }
    }

    private static void flattenObject(JSONObject inputObject, Map<String, Object> flattenedMap, String parentKey, Set<String> allValidColumns) {
        for (Object key : inputObject.keySet()) {
            String newKey = parentKey.isEmpty() ? (String) key : parentKey + "_" + key;
            Object value = inputObject.get(key);
            flattenObjectOrArray(value, flattenedMap, newKey, allValidColumns);
        }
    }

    private static void flattenArray(JSONArray inputArray, Map<String, Object> flattenedMap, String parentKey, Set<String> allValidColumns) {
        inputArray.forEach(inputObj -> {
            String newKey = parentKey.isEmpty() ? "" : parentKey;
            flattenObjectOrArray(inputObj, flattenedMap, newKey, allValidColumns);
        });
    }

}
