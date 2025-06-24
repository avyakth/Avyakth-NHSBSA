package com.nhsbsa.util;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.jayway.jsonpath.JsonPath;

public class ResponseToRequest {
    public static void updateExcelWithJsonResponse(File excelFile, File jsonResponseFile, String TCName) throws IOException {
        // Load the Excel file
        FileInputStream inputStream = new FileInputStream(excelFile);
        Workbook workbook = WorkbookFactory.create(inputStream);

        // Iterate over all sheets in the workbook
        //for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate over all rows in the sheet
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Check if the scenario ID matches
                Cell scenarioIdCell = row.getCell(0);
                if (scenarioIdCell != null && scenarioIdCell.getCellType() == CellType.STRING
                        && scenarioIdCell.getStringCellValue().equals(TCName)) {

                    // Iterate over all cells in the row
                   // Iterator<Cell> cellIterator = row.cellIterator();
                   // while (cellIterator.hasNext()) {
                     //   Cell cell = cellIterator.next();

                        // Check if the cell is a field name
                      //  if (cell.getColumnIndex() == 1 && cell.getCellType() == CellType.STRING) {
                         //  
                			Cell field = row.getCell(1);
                			 String fieldName = field.getStringCellValue();
                            Cell jsonPathCell = row.getCell(2);
                            if (jsonPathCell != null && jsonPathCell.getCellType() == CellType.STRING) {
                                String jsonPath = jsonPathCell.getStringCellValue();
                                String jsonResponse = getJsonResponse(jsonResponseFile);
                                Object value = JsonPath.read(jsonResponse, jsonPath);

                                // Update the value in the Excel file
                                Cell valueCell = row.getCell(3);
                                if (valueCell == null) {
                                    valueCell = row.createCell(3);
                                }
                                if (value instanceof String) {
                                    valueCell.setCellValue((String) value);
                                } else if (value instanceof Boolean) {
                                    valueCell.setCellValue((Boolean) value);
                                } else if (value instanceof Double) {
                                    valueCell.setCellValue((Double) value);
                                } else if (value instanceof Integer) {
                                    valueCell.setCellValue((Integer) value);
                                }
                            }
                        }
                    }
              //  }
          //  }
       // }

        FileOutputStream outputStream = new FileOutputStream(excelFile);
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }

    private static String getJsonResponse(File file) throws IOException {
        BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(file), "UTF-8"));
        StringBuilder builder = new StringBuilder();
        String line;
        while ((line = reader.readLine()) != null) {
            builder.append(line);
        }
        reader.close();
        return builder.toString();
    }
        
    
    public static Map<String, String> readExcelResponse(String file, String sheetName, String TCName) throws IOException {
        Map<String, String> map = new HashMap<>();
        Workbook workbook = WorkbookFactory.create(new FileInputStream(file));
        Sheet sheet = workbook.getSheet(sheetName);
        Map<String, Integer> headerMap = getHeaderMap(sheet.getRow(0));
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                String rowScenarioId = getCellValue(row.getCell(headerMap.get("scenarioid")));
                if (rowScenarioId != null && rowScenarioId.equals(TCName)) {
                    String fieldName = getCellValue(row.getCell(headerMap.get("fieldname")));
                    String value = getCellValue(row.getCell(headerMap.get("value")));
                    if (fieldName != null && value != null) {
                        map.put(fieldName, value);
                    }
                }
            }
        }
        workbook.close();
        return map;
    }

    private static Map<String, Integer> getHeaderMap(Row headerRow) {
        Map<String, Integer> headerMap = new HashMap<>();
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            if (cell != null) {
                headerMap.put(cell.getStringCellValue().toLowerCase(), i);
            }
        }
        return headerMap;
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        switch (cell.getCellType()) {
        case STRING:
            return cell.getStringCellValue();
        case NUMERIC:
            return String.valueOf(cell.getNumericCellValue());
        case BOOLEAN:
            return String.valueOf(cell.getBooleanCellValue());
        case FORMULA:
            return String.valueOf(cell.getCellFormula());
        default:
            return null;
        }
}

}