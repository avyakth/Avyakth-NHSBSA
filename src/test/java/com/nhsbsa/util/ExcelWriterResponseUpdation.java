package com.nhsbsa.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

//import org.apache.commons.lang.ObjectUtils.Null;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import com.fasterxml.jackson.core.JsonParser;

public class ExcelWriterResponseUpdation {

@SuppressWarnings({ "resource", "unchecked" })
	public static void updateExcel(String filePath, String sheetName, String scenarioId, JSONObject responseValues) throws IOException {
        FileInputStream inputStream = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheet(sheetName);
        Iterator<Row> iterator = sheet.iterator();
        int responseKeyIndex = 2;
        int responseValueIndex = 3;
        int scenarioIdIndex = 0;
        Row headerRow = iterator.next();
        Iterator<Cell> headerIterator = headerRow.cellIterator();
        while (headerIterator.hasNext()) {
            Cell cell = headerIterator.next();
            

            if (cell.getStringCellValue().equals("Responsekey")) {
                responseKeyIndex = cell.getColumnIndex();
            } else if (cell.getStringCellValue().equals("ResponseValue")) {
                responseValueIndex = cell.getColumnIndex();
            } else if (cell.getStringCellValue().equals("ScenarioId")) {
                scenarioIdIndex = cell.getColumnIndex();
            }
        }
        while (iterator.hasNext()) {
            Row row = iterator.next();
            Cell scenarioIdCell = row.getCell(scenarioIdIndex);
            if(row.getCell(scenarioIdIndex)!=null)
            {
            if (scenarioIdCell.getStringCellValue().equals(scenarioId)) {
                Cell responseKeyCell = row.getCell(responseKeyIndex);
                String responseKey = responseKeyCell.getStringCellValue();
                Object aObj = responseValues.get(responseKey);
                if (responseValues.has(responseKey)) {
                    Cell responseValueCell = row.createCell(responseValueIndex);
	        		if(aObj instanceof Integer){
	        			responseValueCell.setCellValue(responseValues.getInt(responseKey));
	        		}
	        		else if (aObj instanceof String) {
	        			responseValueCell.setCellValue(responseValues.getString(responseKey));
					}
	        		else if (aObj instanceof Boolean) {
	        			responseValueCell.setCellValue(responseValues.getBoolean(responseKey));
					}
	        		else if (aObj instanceof Float) {
	        			responseValueCell.setCellValue(responseValues.getFloat(responseKey));
					}
	        		else {
	        			responseValueCell.setCellValue(responseValues.getLong(responseKey));
						
					}{
	        			
                    }
                }
            }
        }
        }
        FileOutputStream outputStream = new FileOutputStream(filePath);
        workbook.write(outputStream);
    }
}
