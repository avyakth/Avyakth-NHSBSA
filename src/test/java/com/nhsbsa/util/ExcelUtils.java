package com.nhsbsa.util;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import com.jayway.jsonpath.JsonPath;

import java.io.File;



public class ExcelUtils {
	    public static JSONArray readExcelFileAsJsonObject_RowWise(String filePath, String SheetName) {
	        DataFormatter dataFormatter = new DataFormatter();
	        JSONArray sheetJson = new JSONArray();
	        try {

	            FileInputStream excelFile = new FileInputStream(new File(filePath));
	            Workbook workbook = new XSSFWorkbook(excelFile);
	            FormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
	            int sheets=workbook.getNumberOfSheets();
	            for(int i=0;i<sheets;i++){
					if(workbook.getSheetName(i).equalsIgnoreCase(SheetName)) {
					JSONObject rowJson = new JSONObject();
		            for (Sheet sheet : workbook) {
		                sheetJson = new JSONArray();
		                int lastRowNum = sheet.getLastRowNum();
		                int lastColumnNum = sheet.getRow(0).getLastCellNum();
		                Row firstRowAsKeys = sheet.getRow(0); // first row as a json keys

		                for (int j = 1; j <= lastRowNum; j++) {
		                    rowJson = new JSONObject();
		                    Row row = sheet.getRow(j);

		                    if (row != null) {
		                        for (int k = 0; k< lastColumnNum; k++) {
		                            formulaEvaluator.evaluate(row.getCell(k));
		                            rowJson.put(firstRowAsKeys.getCell(k).getStringCellValue(),
		                                    dataFormatter.formatCellValue(row.getCell(k), formulaEvaluator));
		                        }
		                        sheetJson.put(rowJson);
		                    }
		                }
		            }
	            }
	            }
	            
	            
	        } catch (Exception e) {
	            e.printStackTrace();
	        }
	        return sheetJson;
	    }

	    public static JSONObject getJSONDataReqVal(JSONArray reqData, String TCName) throws JSONException {
	        int index=-1;
	        for (int i =0;i<reqData.length();i++){
	            JSONObject testObj = reqData.getJSONObject(i);
	            Iterator jsonKeys = testObj.keys();
	            while (jsonKeys.hasNext()){
	                String dummyVal = testObj.get(jsonKeys.next().toString()).toString();
	                if (dummyVal.toString().equals(TCName)){
	                    index = i;
	                    break;
	                }
	            }
	           if (index >= 0){
	               break;
	            }
	        }
	       
	    
	        return reqData.getJSONObject(index);
	        }
	    

}
//import java.io.FileInputStream;
//import java.io.IOException;
//import java.lang.reflect.Field;
//import java.util.ArrayList;
//import java.util.Iterator;
//import java.util.List;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import com.fasterxml.jackson.databind.ObjectMapper;
//
//import Pojo.CreatingUser;
//
//public class ExcelUtils {
//    
//    @SuppressWarnings("deprecation")
//	public static <T> List<T> readDataFromExcel(Class<T> clazz, String filePath, String sheetName) throws IOException {
//        FileInputStream inputStream = new FileInputStream(filePath);
//        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
//        Sheet sheet = workbook.getSheet(sheetName);
//        Iterator<Row> rowIterator = sheet.iterator();
//        List<T> dataList = new ArrayList<>();
//        List<String> fieldNames = new ArrayList<>();
//        while (rowIterator.hasNext()) {
//            Row row = rowIterator.next();
//            if (row.getRowNum() == 0) {
//                Iterator<Cell> cellIterator = row.cellIterator();
//                while (cellIterator.hasNext()) {
//                    Cell cell = cellIterator.next();
//                    fieldNames.add(cell.getStringCellValue());
//                }
//            } else {
//                T data = null;
//                try {
//                    data = clazz.newInstance();
//                } catch (InstantiationException | IllegalAccessException e) {
//                    e.printStackTrace();
//                }
//                Iterator<Cell> cellIterator = row.cellIterator();
//                while (cellIterator.hasNext()) {
//                    Cell cell = cellIterator.next();
//                    int columnIndex = cell.getColumnIndex();
//                    String fieldName = fieldNames.get(columnIndex);
//                    try {
//                        Field field = clazz.getDeclaredField(fieldName);
//                        field.setAccessible(true);
//                        if (field.getType() == String.class) {
//                            field.set(data, cell.getStringCellValue());
//                        } else if (field.getType() == int.class || field.getType() == Integer.class) {
//                            field.set(data, (int) cell.getNumericCellValue());
//                        } else if (field.getType() == double.class || field.getType() == Double.class) {
//                            field.set(data, cell.getNumericCellValue());
//                        } else if (field.getType() == boolean.class || field.getType() == Boolean.class) {
//                            field.set(data, cell.getBooleanCellValue());
//                        }
//                    } catch (NoSuchFieldException | SecurityException | IllegalArgumentException | IllegalAccessException e) {
//                        e.printStackTrace();
//                    }
//                }
//                dataList.add(data);
//            }
//        }
//        workbook.close();
//        return dataList;
//    }
//    
//    public static void main(String[] args) throws IOException {
//        List<CreatingUser> dataList = ExcelUtils.readDataFromExcel(CreatingUser.class, "C:\\Users\\parthiban.vijayan\\Downloads\\NextGenBDDAPI\\NextGenBDD\\sample_RequestData_updated.xlsx", "MandatoryDataReq");
//        ObjectMapper objectMapper = new ObjectMapper();
//        String jsonString = objectMapper.writeValueAsString(dataList);
//        System.out.println(jsonString);
//    }
//}

