package com.nhsbsa.util;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import io.restassured.response.Response;
import io.restassured.specification.ResponseSpecification;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.io.Closeable;
import java.io.File;
import java.io.FileWriter;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.NumberToTextConverter;

public class ExcelUtilResponse {

	public static FileInputStream excelFile;
	public static Workbook workbook;
	public static Sheet sheet;
	Row row;

	public static final String STATUS_COLUMN_NAME = "status";
	@SuppressWarnings("unlikely-arg-type")
	public void writeDataToExcel(String filePath, String sheetName, String TCName, String Value,
			int rowID , int Column) throws IOException {
		try {
			int scenario_id_col = 0;
			//int status_col = 5;
			Workbook workbook = null;
			Sheet sheet = null;
			Row row = null;
			File file = new File(filePath);
			FileInputStream inputStream = new FileInputStream(file);
			workbook = new XSSFWorkbook(inputStream);
			sheet = workbook.getSheet(sheetName);
			row = sheet.getRow(rowID);

			Cell scenario_id_cell = row.getCell(scenario_id_col);
			if (scenario_id_cell != null && (TCName.toLowerCase().contains(scenario_id_cell.getStringCellValue().toLowerCase()))) {
				Cell cell = row.getCell(Column, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

				cell.setCellValue(Value); // or "fail"
				System.out.println( Value + " : updated in sheet");
			}

			// Save the changes to the Excel workbook

			workbook.write(new FileOutputStream(new File(filePath)));
			excelFile.close();
			workbook.close();



		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public static List<List<String>> readExcel(String filePath, String sheetName) throws IOException {
		excelFile = new FileInputStream(new File(filePath));
		workbook = WorkbookFactory.create(excelFile);
		sheet = workbook.getSheet(sheetName);
		List<List<String>> data = new ArrayList<>();
		for (Row row : sheet) {
			List<String> rowData = new ArrayList<>();
			for (Cell cell : row) {
				switch (cell.getCellType()) {
				case STRING:
					rowData.add(cell.getStringCellValue());
					break;
				case NUMERIC:
					rowData.add((NumberToTextConverter.toText(cell.getNumericCellValue())));
					break;
				case BOOLEAN:
					rowData.add(String.valueOf(cell.getBooleanCellValue()));
					break;
				default:
				}

			}
			data.add(rowData);
		}
		return data;
	}

	public static FileWriter fileWriter(File file ,String fileName, Response RESP) throws IOException {
		FileWriter fileWriter = new FileWriter(file);
		fileWriter.write(RESP.getBody().asString());
		fileWriter.close();
		return fileWriter;
	}

	public static void writeJsonData(String filePath, String jsonData) throws IOException {
		try (FileWriter file = new FileWriter(filePath)) {
			file.write(jsonData);
			file.flush();
		}
	}
}

