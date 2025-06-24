package com.nhsbsa.util;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Properties;
import org.apache.commons.io.FileUtils;
import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;
import com.nhsbsa.steps.CommonFunctions;

import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Reportgenerator extends CommonFunctions {
	// static String selectedApp = "";
	static String launchURL = "";
	public static String appName = ITestListenerImpl.defaultConfigProperty.get().getProperty("AppName");
	public static String scenarioName = "";
	// static String featureName = "";
	public static String elementName = "";
	public static String ErrType = "";
	public static String comments = "";
	public static String filepath = "";
	public static String currentDirExcel = "";
	public static ThreadLocal<Integer> stepFailCnt = new ThreadLocal<Integer>();
	public static ThreadLocal<Integer> stepPassCnt = new ThreadLocal<Integer>();
	// public static ThreadLocal<String> comments = new ThreadLocal<String>();
	static int Critical_Error_Counter = 0;
	static String User_Name = ITestListenerImpl.defaultConfigProperty.get().getProperty("MailFromTeam");
	static String teamEmailID = ITestListenerImpl.defaultConfigProperty.get().getProperty("MailFromTeamID");
	static int Critical_Counter = 0;
	static int Failed_Case_Counter = 0;
	static int Passed_Case_Counter = 0;

	// public static synchronized Reportgenerator getReport() {
	// try {
	// Reportgenerator report = new Reportgenerator();
	// return report;
	// } catch(Exception e) {
	// e.printStackTrace();
	// return null;
	// }
	// }

	public static synchronized void report1(Boolean status, String Cmnt, String scenarioName)
			throws IOException, RowsExceededException, WriteException, BiffException {
		appName = "MP";
		// scenarioName = annotation.scenarioName;
		System.out.println("After Report - Status - " + status + " scenarioName - " + scenarioName);
		// featureName = annotation.featureName;
		// comments for element
		if (status == false) {
			// comments = Cmnt;
			comments = "";
			System.out.println("comments " + comments);
		} else {
			comments = "";
		}
		// else { comments = "ERROR MESSAGE:" + element.getErrorMessage(); comments +=
		// "\nERROR:The " + event + " action was not performed."; }
		// crtical ErrType if ((techtablet.annotation.criticalCount > 0) &&
		// !(status)) { ErrType = "CRITICAL"; } else { ErrType = "NON-CRTICAL"; }
		Calendar cal = Calendar.getInstance();
		DateFormat dateFormat = new SimpleDateFormat("MM_dd_yyyy");
		String cal1 = dateFormat.format(cal.getTime());
		String currentDir = System.getProperty("user.dir");
		String filepath = currentDir + File.separator + "Results" + File.separator + appName + " Final Report " + cal1
				+ ".xls";
		File ifilepath = new File(
				currentDir + File.separator + "Results" + File.separator + appName + " Final Report " + cal1 + ".xls");
		String ofilepath = currentDir + File.separator + "Results" + File.separator + appName + " Final Report " + cal1
				+ "_temp.xls";
		File logfile = new File(filepath);// Created object of java File
		// String filepath ="c:\\bdd\\" + appName + " Final Report " + cal1 + ".xls";
		// File ifilepath = new File("c:\\bdd\\" + appName + " Final Report" + cal1 +
		// ".xls"); String ofilepath = "c:\\bdd\\" + appName + " Final Report" + cal1 +
		// "_temp.xls"; File logfile = new File(filepath);
		// Created object of java File
		// String filepath = currentDir + File.separator + "Results\\" + appName +
		// " Final Report " + cal1 + "//Results//reports.xls"; File ifilepath = new
		// File(currentDir + File.separator + "Results\\" + appName + " Final Report " +
		// cal1 + "//Results//reports.xls"); String ofilepath = currentDir +
		// File.separator + "Results\\" + appName + " Final Report " + cal1 +
		// "//Results//temp.xls"; File logfile =new File(filepath);// Created object of
		// java File class.

		if (!logfile.exists()) { // if1
			WritableWorkbook workbook = Workbook.createWorkbook(new File(filepath));
			WritableSheet sheet = workbook.createSheet("Report", 0);
			sheet.setName("Report");
			WritableFont arialfont = new WritableFont(WritableFont.ARIAL, 11, WritableFont.BOLD);
			WritableCellFormat cellFormat = new WritableCellFormat(arialfont);
			cellFormat.setBackground(Colour.ICE_BLUE);
			cellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
			sheet.addCell(new Label(0, 0, "Date", cellFormat));
			/* sheet.addCell(new Label(1, 0, "Feature", cellFormat)); */
			sheet.addCell(new Label(1, 0, "Scenario", cellFormat));
			sheet.addCell(new Label(2, 0, "TestStatus", cellFormat));
			sheet.addCell(new Label(3, 0, "Error Type", cellFormat));
			sheet.addCell(new Label(4, 0, "Comments", cellFormat));
			workbook.write();
			workbook.close();
		} // if1 ends
		Workbook wb1 = Workbook.getWorkbook(ifilepath);
		WritableWorkbook wbcopy = Workbook.createWorkbook(new File(ofilepath), wb1);
		WritableSheet sheet1 = wbcopy.getSheet(0);
		Sheet sheet = wb1.getSheet(0);
		int newrow = sheet.getRows();
		sheet1.setName("Report");
		WritableFont arialfont1 = new WritableFont(WritableFont.ARIAL, 10);
		WritableCellFormat cellFormat1 = new WritableCellFormat(arialfont1);
		cellFormat1.setBorder(Border.ALL, BorderLineStyle.THIN);
		WritableCellFormat passcellFormat = new WritableCellFormat(arialfont1);
		passcellFormat.setBackground(Colour.LIGHT_GREEN);
		passcellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
		WritableCellFormat failcellFormat = new WritableCellFormat(arialfont1);
		failcellFormat.setBackground(Colour.RED);
		failcellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
		int col = 0;
		int widthInChars = 12;
		sheet1.setColumnView(col, widthInChars);
		sheet1.addCell(new Label(col, newrow, cal1, cellFormat1));
		// col = 1; widthInChars = 16; sheet1.setColumnView(col, widthInChars);
		// sheet1.addCell(new Label(col, newrow, featureName, cellFormat1));
		col = 1;
		widthInChars = 70;
		sheet1.setColumnView(col, widthInChars);
		sheet1.addCell(new Label(col, newrow, scenarioName, cellFormat1));
		col = 2;
		widthInChars = 15;
		sheet1.setColumnView(col, widthInChars);
		if (status) { // if2
			sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
			col = 3;
			widthInChars = 20;
			sheet1.setColumnView(col, widthInChars);
			sheet1.addCell(new Label(col, newrow, "", cellFormat1));
			comments = "";
		} else {
			sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
			col = 3;
			widthInChars = 20;
			sheet1.setColumnView(col, widthInChars);
			sheet1.addCell(new Label(col, newrow, ErrType, cellFormat1));
			col = 4;
			widthInChars = 20;
			sheet1.setColumnView(col, widthInChars);
			sheet1.addCell(new Label(col, newrow, comments, cellFormat1));
		} // if2 ends
		wb1.close();
		wbcopy.write();
		wbcopy.close();
		logfile.delete();
		Workbook wb2 = Workbook.getWorkbook(new File(ofilepath));
		WritableWorkbook wbmain = Workbook.createWorkbook(new File(filepath), wb2);
		WritableSheet sheet2 = wbcopy.getSheet(0);
		sheet2.setName("Report");
		wbmain.write();
		wbmain.close();
		new File(ofilepath).delete();
	} // report ends

	public static synchronized void report2(Boolean status, String Cmnt, String scenarioName)
			throws IOException, RowsExceededException, WriteException, BiffException {
		appName = "MP";
		System.out.println("After Report - Status - " + status + " scenarioName - " + scenarioName);
		// featureName = annotation.featureName;
		// comments for element
		if (status == false) {
			// comments = Cmnt;
			comments = "";
			System.out.println("comments " + comments);
		} else {
			comments = "";
		}
		// else { comments = "ERROR MESSAGE:" + element.getErrorMessage(); comments +=
		// "\nERROR:The " + event + " action was not performed."; }
		// crtical ErrType if ((techtablet.annotation.criticalCount > 0) &&
		// !(status)) { ErrType = "CRITICAL"; } else { ErrType = "NON-CRTICAL"; }
		Calendar cal = Calendar.getInstance();
		DateFormat dateFormat = new SimpleDateFormat("MM_dd_yyyy");
		String cal1 = dateFormat.format(cal.getTime());
		String currentDir = System.getProperty("user.dir");
		String filepath = currentDir + File.separator + "Results" + File.separator + appName + " Final Report " + cal1
				+ ".xls";
		File ifilepath = new File(
				currentDir + File.separator + "Results" + File.separator + appName + " Final Report " + cal1 + ".xls");
		String ofilepath = currentDir + File.separator + "Results" + File.separator + appName + " Final Report " + cal1
				+ "_temp.xls";
		File logfile = new File(filepath);// Created object of java File
		if (!logfile.exists()) { // if1
			WritableWorkbook writeWB = Workbook.createWorkbook(new File(filepath));
			WritableSheet sheet = writeWB.createSheet("Report", 0);
			sheet.setName("Report");
			WritableFont arialfont = new WritableFont(WritableFont.ARIAL, 11, WritableFont.BOLD);
			WritableCellFormat blueFormat = new WritableCellFormat(arialfont);
			blueFormat.setBackground(Colour.ICE_BLUE);
			blueFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
			sheet.addCell(new Label(0, 0, "S.NO", blueFormat));
			sheet.addCell(new Label(1, 0, "TESTCASE", blueFormat));
			sheet.addCell(new Label(2, 0, "OVERALL STATUS", blueFormat));
			sheet.addCell(new Label(3, 0, "CHROME", blueFormat));
			sheet.addCell(new Label(4, 0, "IE", blueFormat));
			sheet.addCell(new Label(5, 0, "EDGE", blueFormat));
			sheet.addCell(new Label(6, 0, "ERRTYPE", blueFormat));
			sheet.addCell(new Label(7, 0, "COMMENTS", blueFormat));
			writeWB.write();
			writeWB.close();
		} // if1 ends

		Workbook readWB = Workbook.getWorkbook(ifilepath);
		WritableWorkbook writeWBCopy = Workbook.createWorkbook(new File(ofilepath), readWB);
		WritableSheet writeSheet = writeWBCopy.getSheet(0);
		Sheet readSheet = readWB.getSheet(0);
		int newrow = readSheet.getRows();
		int totalNoOfCols = readSheet.getColumns();
		writeSheet.setName("Report");
		WritableFont arialfont1 = new WritableFont(WritableFont.ARIAL, 10);
		WritableCellFormat whiteFormat = new WritableCellFormat(arialfont1);
		whiteFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
		WritableCellFormat passcellFormat = new WritableCellFormat(arialfont1);
		passcellFormat.setBackground(Colour.LIGHT_GREEN);
		passcellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
		WritableCellFormat failcellFormat = new WritableCellFormat(arialfont1);
		failcellFormat.setBackground(Colour.RED);
		failcellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
		// =================COLUMN 0 S.NO=========================
		int col = 0;
		int widthInChars = 12;
		writeSheet.setColumnView(col, widthInChars);
		writeSheet.addCell(new Label(col, newrow, Integer.toString(newrow), whiteFormat));
		// =================COLUMN 1 SCENARIO=========================
		System.out.println("In Scenario1");
		col = 1;
		widthInChars = 70;
		writeSheet.setColumnView(col, widthInChars);
		for (int i = 1; i < newrow; i++) {
			if (!readSheet.getCell(1, i).getContents().equalsIgnoreCase(scenarioName)) {
				writeSheet.addCell(new Label(col, newrow, scenarioName, whiteFormat));
			}
			// =================COLUMN 3,4,5 SCENARIO=========================
			System.out.println("browserName.get() " + browserName.get());
			if (browserName.get().equalsIgnoreCase("Chrome"))
				if (status) {
					writeSheet.addCell(new Label(3, newrow, "PASS", passcellFormat));
				} else if (!status) {
					writeSheet.addCell(new Label(3, newrow, "FAIL", failcellFormat));
				} else {
					writeSheet.addCell(new Label(3, newrow, "-", whiteFormat));
				}
			else if (browserName.get().equalsIgnoreCase("IE")) {
				if (status) {
					writeSheet.addCell(new Label(4, newrow, "PASS", passcellFormat));
				} else if (!status) {
					writeSheet.addCell(new Label(4, newrow, "FAIL", failcellFormat));
				} else {
					writeSheet.addCell(new Label(4, newrow, "-", whiteFormat));
				}
			} else if (browserName.get().equalsIgnoreCase("EDGE")) {
				if (status) {
					writeSheet.addCell(new Label(5, newrow, "PASS", passcellFormat));
				} else if (!status) {
					writeSheet.addCell(new Label(5, newrow, "FAIL", failcellFormat));
				} else {
					writeSheet.addCell(new Label(5, newrow, "-", whiteFormat));
				}
			}
		}
		readWB.close();
		writeWBCopy.write();
		writeWBCopy.close();
		logfile.delete();
		Workbook wb23 = Workbook.getWorkbook(new File(ofilepath));
		WritableWorkbook wbmain23 = Workbook.createWorkbook(new File(filepath), wb23);
		WritableSheet sheet23 = writeWBCopy.getSheet(0);
		sheet23.setName("Report");
		wbmain23.write();
		wbmain23.close();
		new File(ofilepath).delete();
		// =================COLUMN 2 OVERALL STATUS=========================
		Workbook readWB1 = Workbook.getWorkbook(ifilepath);
		WritableWorkbook writeWB1 = Workbook.createWorkbook(new File(ofilepath), readWB1);
		WritableSheet writeSheet1 = writeWBCopy.getSheet(0);
		Sheet readSheet1 = readWB1.getSheet(0);
		col = 2;
		widthInChars = 15;
		/* writeSheet1.setColumnView(col, widthInChars); */
		// if (readSheet1.getCell(3,
		// newrow).getContents().equalsIgnoreCase("FAIL")&&readSheet1.getCell(4,
		// newrow).getContents().equalsIgnoreCase("FAIL")||readSheet1.getCell(5,
		// newrow).getContents().equalsIgnoreCase("FAIL")) {
		if (readSheet1.getCell(3, newrow).getContents().equalsIgnoreCase("FAIL")) {
			writeSheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
			col = 6;
			widthInChars = 20;
			writeSheet1.setColumnView(col, widthInChars);
			writeSheet1.addCell(new Label(col, newrow, ErrType, whiteFormat));
			col = 7;
			widthInChars = 20;
			writeSheet1.setColumnView(col, widthInChars);
			writeSheet1.addCell(new Label(col, newrow, comments, whiteFormat));
		} else {
			col = 2;
			writeSheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
			col = 7;
			widthInChars = 20;
			writeSheet1.setColumnView(col, widthInChars);
			writeSheet1.addCell(new Label(col, newrow, "", whiteFormat));
			comments = "";
		}
		// if2 ends
		readWB1.close();
		writeWB1.write();
		writeWB1.close();
		logfile.delete();
		Workbook wb2 = Workbook.getWorkbook(new File(ofilepath));
		WritableWorkbook wbmain = Workbook.createWorkbook(new File(filepath), wb2);
		WritableSheet sheet2 = writeWB1.getSheet(0);
		sheet2.setName("Report");
		wbmain.write();
		wbmain.close();
		new File(ofilepath).delete();
	} // report ends

	public static synchronized void report3(Boolean status, String Cmnt, String scenarioName)
			throws IOException, RowsExceededException, WriteException, BiffException {
		// appName = ITestListenerImpl.property.get().getProperty("AppName");
		// scenarioName = annotation.scenarioName;
		System.out.println("After Report - Status - " + " Thread: " + Thread.currentThread().getId() + " Browser: "
				+ browserName.get() + status + " scenarioName - " + scenarioName);
		// featureName = annotation.featureName;
		// comments for element

		if (status == false) {
			// comments = Cmnt;
			comments = "";
			System.out.println("comments " + comments);
		} else {
			comments = "";
		}
		/*
		 * else { comments = "ERROR MESSAGE:" + element.getErrorMessage(); comments +=
		 * "\nERROR:The " + event + " action was not performed."; }
		 */

		/*
		 * // crtical ErrType if ((techtablet.annotation.criticalCount > 0) &&
		 * !(status)) { ErrType = "CRITICAL"; } else { ErrType = "NON-CRTICAL"; }
		 */

		Calendar cal = Calendar.getInstance();
		DateFormat dateFormat = new SimpleDateFormat("MM_dd_yyyy");
		String cal1 = dateFormat.format(cal.getTime());
		String currentDir = System.getProperty("user.dir");
		String filepath = currentDir + File.separator + "Results" + File.separator + appName + " Final Report " + cal1
				+ ".xls";
		File ifilepath = new File(
				currentDir + File.separator + "Results" + File.separator + appName + " Final Report " + cal1 + ".xls");
		String ofilepath = currentDir + File.separator + "Results" + File.separator + appName + " Final Report " + cal1
				+ "_temp.xls";
		File logfile = new File(filepath);// Created object of java File
		/*
		 * String filepath ="c:\\bdd\\" + appName + " Final Report " + cal1 + ".xls";
		 * File ifilepath = new File("c:\\bdd\\" + appName + " Final Report
		 * " + cal1 + ".xls"); String ofilepath = "c:\\bdd\\" + appName + " Final Report
		 * " + cal1 + "_temp.xls"; File logfile = new File(filepath);
		 */// Created object of java File

		/*
		 * String filepath = currentDir + File.separator + "Results\\" + appName +
		 * " Final Report " + cal1 + "//Results//reports.xls"; File ifilepath = new
		 * File(currentDir + File.separator + "Results\\" + appName + " Final Report " +
		 * cal1 + "//Results//reports.xls"); String ofilepath = currentDir +
		 * File.separator + "Results\\" + appName + " Final Report " + cal1 +
		 * "//Results//temp.xls"; File logfile = new File(filepath);// Created object of
		 * java File class.
		 */

		if (!logfile.exists()) { // if1
			WritableWorkbook workbook = Workbook.createWorkbook(new File(filepath));
			WritableSheet sheet = workbook.createSheet("Report", 0);
			sheet.setName("Report");
			WritableFont arialfont = new WritableFont(WritableFont.ARIAL, 11, WritableFont.BOLD);
			WritableCellFormat cellFormat = new WritableCellFormat(arialfont);
			cellFormat.setBackground(Colour.ICE_BLUE);
			cellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
			sheet.addCell(new Label(0, 0, "S.NO", cellFormat));
			sheet.addCell(new Label(1, 0, "TESTCASE", cellFormat));
			sheet.addCell(new Label(2, 0, "STATUS", cellFormat));
			sheet.addCell(new Label(3, 0, "CHROME", cellFormat));
			sheet.addCell(new Label(4, 0, "IE", cellFormat));
			sheet.addCell(new Label(5, 0, "EDGE", cellFormat));
			sheet.addCell(new Label(6, 0, "FAIL REASON", cellFormat));
			sheet.addCell(new Label(7, 0, "COMMENTS", cellFormat));
			workbook.write();
			workbook.close();
		} // if1 ends
		Workbook wb1 = Workbook.getWorkbook(ifilepath);
		WritableWorkbook wbcopy = Workbook.createWorkbook(new File(ofilepath), wb1);
		WritableSheet sheet1 = wbcopy.getSheet(0);
		Sheet sheet = wb1.getSheet(0);
		int newrow = sheet.getRows();
		System.out.println("newrow: " + newrow);
		sheet1.setName("Report");
		WritableFont arialfont1 = new WritableFont(WritableFont.ARIAL, 10);
		WritableCellFormat cellFormat1 = new WritableCellFormat(arialfont1);
		cellFormat1.setBorder(Border.ALL, BorderLineStyle.THIN);
		WritableCellFormat passcellFormat = new WritableCellFormat(arialfont1);
		passcellFormat.setBackground(Colour.LIGHT_GREEN);
		passcellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
		WritableCellFormat failcellFormat = new WritableCellFormat(arialfont1);
		failcellFormat.setBackground(Colour.RED);
		failcellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
		int col = 0;
		int widthInChars = 12;

		/*
		 * // =================COLUMN 0 S.NO========================
		 * sheet1.setColumnView(col, widthInChars); sheet1.addCell(new Label(col,
		 * newrow, Integer.toString(newrow), cellFormat1)); // =================COLUMN 1
		 * SCENARIO========================= System.out.println("In Column1 Scenario");
		 * col = 1; widthInChars = 70; sheet1.setColumnView(col, widthInChars);
		 */
		// System.out.println("get content: "+ sheet.getCell(1,
		// 1).getContents().equalsIgnoreCase(scenarioName));
		if (newrow == 1) {
			// =================COLUMN 0 S.NO========================
			col = 0;
			widthInChars = 12;
			sheet1.setColumnView(col, widthInChars);
			sheet1.addCell(new Label(col, newrow, Integer.toString(newrow), cellFormat1));
			// =================COLUMN 1 SCENARIO=========================
			System.out.println("In Column1 Scenario");
			col = 1;
			widthInChars = 70;
			sheet1.setColumnView(col, widthInChars);
			sheet1.addCell(new Label(col, newrow, scenarioName, cellFormat1));
			if (status) { // if2
				if (browserName.get().equalsIgnoreCase("Chrome")) {
					col = 3;
					widthInChars = 15;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
				} else if (browserName.get().equalsIgnoreCase("internet explorer")) {
					col = 4;
					widthInChars = 15;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
				} else if (browserName.get().equalsIgnoreCase("MicrosoftEdge")) {
					col = 5;
					widthInChars = 15;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
				} else {
					System.out.println("Invalid Browser while creating Excel Report");
				}
				col = 7;
				widthInChars = 20;
				sheet1.setColumnView(col, widthInChars);
				sheet1.addCell(new Label(col, newrow, "", cellFormat1));
				comments = "";
			} else if (!status) {
				if (browserName.get().equalsIgnoreCase("Chrome")) {
					col = 3;
					widthInChars = 15;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
				} else if (browserName.get().equalsIgnoreCase("internet explorer")) {
					col = 4;
					widthInChars = 15;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
				} else if (browserName.get().equalsIgnoreCase("MicrosoftEdge")) {
					col = 5;
					widthInChars = 15;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
				}
				col = 6;
				widthInChars = 20;
				sheet1.setColumnView(col, widthInChars);
				sheet1.addCell(new Label(col, newrow, ErrType, cellFormat1));
				col = 7;
				widthInChars = 20;
				sheet1.setColumnView(col, widthInChars);
				sheet1.addCell(new Label(col, newrow, comments, cellFormat1));
			}
		} else {
			int notDuplicateScenario = 0;
			for (int i = 1; i < newrow; i++) {
				if (!sheet.getCell(1, i).getContents().equalsIgnoreCase(scenarioName)) {
					notDuplicateScenario++;
				}
				System.out.println("notDuplicateScenarios: " + notDuplicateScenario);
				if (notDuplicateScenario == newrow - 1) {
					// =================COLUMN 0 S.NO========================
					col = 0;
					widthInChars = 12;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, Integer.toString(newrow), cellFormat1));
					// =================COLUMN 1 SCENARIO=========================
					col = 1;
					widthInChars = 70;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, scenarioName, cellFormat1));
					if (status) { // if2
						if (browserName.get().equalsIgnoreCase("Chrome")) {
							col = 3;
							widthInChars = 15;
							sheet1.setColumnView(col, widthInChars);
							sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
						} else if (browserName.get().equalsIgnoreCase("internet explorer")) {
							col = 4;
							widthInChars = 15;
							sheet1.setColumnView(col, widthInChars);
							sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
						} else if (browserName.get().equalsIgnoreCase("MicrosoftEdge")) {
							col = 5;
							widthInChars = 15;
							sheet1.setColumnView(col, widthInChars);
							sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
						}
						col = 7;
						widthInChars = 20;
						sheet1.setColumnView(col, widthInChars);
						sheet1.addCell(new Label(col, newrow, "", cellFormat1));
						comments = "";
					} else if (!status) {
						if (browserName.get().equalsIgnoreCase("Chrome")) {
							col = 3;
							widthInChars = 15;
							sheet1.setColumnView(col, widthInChars);
							sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
						} else if (browserName.get().equalsIgnoreCase("internet explorer")) {
							col = 4;
							widthInChars = 15;
							sheet1.setColumnView(col, widthInChars);
							sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
						} else if (browserName.get().equalsIgnoreCase("MicrosoftEdge")) {
							col = 5;
							widthInChars = 15;
							sheet1.setColumnView(col, widthInChars);
							sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
						}
						col = 6;
						widthInChars = 20;
						sheet1.setColumnView(col, widthInChars);
						sheet1.addCell(new Label(col, newrow, ErrType, cellFormat1));
						col = 7;
						widthInChars = 20;
						sheet1.setColumnView(col, widthInChars);
						sheet1.addCell(new Label(col, newrow, comments, cellFormat1));
					}
				} else {
					if (status) { // if2
						if (browserName.get().equalsIgnoreCase("Chrome")) {
							col = 3;
							widthInChars = 15;
							sheet1.setColumnView(col, widthInChars);
							sheet1.addCell(new Label(col, i, "PASS", passcellFormat));
						} else if (browserName.get().equalsIgnoreCase("internet explorer")) {
							col = 4;
							widthInChars = 15;
							sheet1.setColumnView(col, widthInChars);
							sheet1.addCell(new Label(col, i, "PASS", passcellFormat));
						} else if (browserName.get().equalsIgnoreCase("MicrosoftEdge")) {
							col = 5;
							widthInChars = 15;
							sheet1.setColumnView(col, widthInChars);
							sheet1.addCell(new Label(col, i, "PASS", passcellFormat));
						}
						col = 7;
						widthInChars = 20;
						sheet1.setColumnView(col, widthInChars);
						sheet1.addCell(new Label(col, i, "", cellFormat1));
						comments = "";
					} else if (!status) {
						if (browserName.get().equalsIgnoreCase("Chrome")) {
							col = 3;
							widthInChars = 15;
							sheet1.setColumnView(col, widthInChars);
							sheet1.addCell(new Label(col, i, "FAIL", failcellFormat));
						} else if (browserName.get().equalsIgnoreCase("internet explorer")) {
							col = 4;
							widthInChars = 15;
							sheet1.setColumnView(col, widthInChars);
							sheet1.addCell(new Label(col, i, "FAIL", failcellFormat));
						} else if (browserName.get().equalsIgnoreCase("MicrosoftEdge")) {
							col = 5;
							widthInChars = 15;
							sheet1.setColumnView(col, widthInChars);
							sheet1.addCell(new Label(col, i, "FAIL", failcellFormat));
						}
						col = 6;
						widthInChars = 20;
						sheet1.setColumnView(col, widthInChars);
						sheet1.addCell(new Label(col, i, ErrType, cellFormat1));
						col = 7;
						widthInChars = 20;
						sheet1.setColumnView(col, widthInChars);
						sheet1.addCell(new Label(col, i, comments, cellFormat1));
					}
				}
			}
		}
		// =================COLUMN 3,4,5,6,7 SCENARIO=========================
		System.out.println("In Column 3,4,5,6,7 Scenario");
		/*
		 * if (status) { // if2 if (browserName.get().equalsIgnoreCase("Chrome")) { col
		 * = 3; widthInChars = 15; sheet1.setColumnView(col, widthInChars);
		 * sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat)); } else if
		 * (browserName.get().equalsIgnoreCase("internet explorer")) { col = 4;
		 * widthInChars = 15; sheet1.setColumnView(col, widthInChars);
		 * sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat)); } else if
		 * (browserName.get().equalsIgnoreCase("MicrosoftEdge")) { col = 5; widthInChars
		 * = 15; sheet1.setColumnView(col, widthInChars); sheet1.addCell(new Label(col,
		 * newrow, "PASS", passcellFormat)); } col = 7; widthInChars = 20;
		 * sheet1.setColumnView(col, widthInChars); sheet1.addCell(new Label(col,
		 * newrow, "", cellFormat1)); comments = ""; } else if (!status){ if
		 * (browserName.get().equalsIgnoreCase("Chrome")) { col = 3; widthInChars = 15;
		 * sheet1.setColumnView(col, widthInChars); sheet1.addCell(new Label(col,
		 * newrow, "FAIL", failcellFormat)); } else if
		 * (browserName.get().equalsIgnoreCase("internet explorer")) { col = 4;
		 * widthInChars = 15; sheet1.setColumnView(col, widthInChars);
		 * sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat)); } else if
		 * (browserName.get().equalsIgnoreCase("MicrosoftEdge")) { col = 5; widthInChars
		 * = 15; sheet1.setColumnView(col, widthInChars); sheet1.addCell(new Label(col,
		 * newrow, "FAIL", failcellFormat)); } col = 6; widthInChars = 20;
		 * sheet1.setColumnView(col, widthInChars); sheet1.addCell(new Label(col,
		 * newrow, ErrType, cellFormat1)); col = 7; widthInChars = 20;
		 * sheet1.setColumnView(col, widthInChars); sheet1.addCell(new Label(col,
		 * newrow, comments, cellFormat1)); }
		 */ // if2 ends
		wb1.close();
		wbcopy.write();
		wbcopy.close();
		logfile.delete();
		Workbook wb2 = Workbook.getWorkbook(new File(ofilepath));
		WritableWorkbook wbmain = Workbook.createWorkbook(new File(filepath), wb2);
		WritableSheet sheet2 = wbcopy.getSheet(0);
		sheet2.setName("Report");
		wbmain.write();
		wbmain.close();
		new File(ofilepath).delete();
	}

	public static synchronized void report(Boolean status, String Cmnt, String scenarioName)
			throws IOException, RowsExceededException, WriteException, BiffException {
		// appName = ITestListenerImpl.property.get().getProperty("AppName");
		// scenarioName = annotation.scenarioName;
		System.out.println("After Report - Status - " + " Thread: " + Thread.currentThread().getId() + " Browser: "
				+ browserName.get() + status + " scenarioName - " + scenarioName);
		// featureName = annotation.featureName;
		// comments for element
		if (status == false) {
			// comments = Cmnt;
			comments = "";
			System.out.println("comments " + comments);
		} else {
			comments = "";
		}
		/*
		 * else { comments = "ERROR MESSAGE:" + element.getErrorMessage(); comments +=
		 * "\nERROR:The " + event + " action was not performed."; }
		 */

		/*
		 * // crtical ErrType if ((techtablet.annotation.criticalCount > 0) &&
		 * !(status)) { ErrType = "CRITICAL"; } else { ErrType = "NON-CRTICAL"; }
		 */
		Calendar cal = Calendar.getInstance();
		DateFormat dateFormat = new SimpleDateFormat("MM_dd_yyyy");
		String cal1 = dateFormat.format(cal.getTime());
		String currentDir = System.getProperty("user.dir");
		String filepath = currentDir + File.separator + "Results" + File.separator + appName + " Final Report " + cal1
				+ ".xls";
		File ifilepath = new File(
				currentDir + File.separator + "Results" + File.separator + appName + " Final Report " + cal1 + ".xls");
		String ofilepath = currentDir + File.separator + "Results" + File.separator + appName + " Final Report " + cal1
				+ "_temp.xls";
		File logfile = new File(filepath);// Created object of java File

		/*
		 * String filepath ="c:\\bdd\\" + appName + " Final Report " + cal1 + ".xls";
		 * File ifilepath = new File("c:\\bdd\\" + appName + " Final Report
		 * " + cal1 + ".xls"); String ofilepath = "c:\\bdd\\" + appName + " Final Report
		 * " + cal1 + "_temp.xls"; File logfile = new File(filepath);
		 */// Created object of java File
		/*
		 * String filepath = currentDir + File.separator + "Results\\" + appName +
		 * " Final Report " + cal1 + "//Results//reports.xls"; File ifilepath = new
		 * File(currentDir + File.separator + "Results\\" + appName + " Final Report " +
		 * cal1 + "//Results//reports.xls"); String ofilepath = currentDir +
		 * File.separator + "Results\\" + appName + " Final Report " + cal1 +
		 * "//Results//temp.xls"; File logfile = new File(filepath);// Created object of
		 * java File class.
		 */
		if (!logfile.exists()) { // if1
			WritableWorkbook workbook = Workbook.createWorkbook(new File(filepath));
			WritableSheet sheet = workbook.createSheet("Report", 0);
			sheet.setName("Report");
			WritableFont arialfont = new WritableFont(WritableFont.ARIAL, 11, WritableFont.BOLD);
			WritableCellFormat cellFormat = new WritableCellFormat(arialfont);
			cellFormat.setBackground(Colour.ICE_BLUE);
			cellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
			sheet.addCell(new Label(0, 0, "S.NO", cellFormat));
			sheet.addCell(new Label(1, 0, "TESTCASE", cellFormat));
			sheet.addCell(new Label(2, 0, "STATUS", cellFormat));
			sheet.addCell(new Label(3, 0, "CHROME", cellFormat));
			sheet.addCell(new Label(4, 0, "IE", cellFormat));
			sheet.addCell(new Label(5, 0, "EDGE", cellFormat));
			sheet.addCell(new Label(6, 0, "FAIL REASON", cellFormat));
			sheet.addCell(new Label(7, 0, "COMMENTS", cellFormat));
			workbook.write();
			workbook.close();
		} // if1 ends
		Workbook wb1 = Workbook.getWorkbook(ifilepath);
		WritableWorkbook wbcopy = Workbook.createWorkbook(new File(ofilepath), wb1);
		WritableSheet sheet1 = wbcopy.getSheet(0);
		Sheet sheet = wb1.getSheet(0);
		int newrow = sheet.getRows();
		System.out.println("newrow: " + newrow);
		sheet1.setName("Report");
		WritableFont arialfont1 = new WritableFont(WritableFont.ARIAL, 10);
		WritableCellFormat cellFormat1 = new WritableCellFormat(arialfont1);
		cellFormat1.setBorder(Border.ALL, BorderLineStyle.THIN);
		WritableCellFormat passcellFormat = new WritableCellFormat(arialfont1);
		passcellFormat.setBackground(Colour.LIGHT_GREEN);
		passcellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
		WritableCellFormat failcellFormat = new WritableCellFormat(arialfont1);
		failcellFormat.setBackground(Colour.RED);
		failcellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
		int col = 0;
		int widthInChars = 12;
		/*
		 * // =================COLUMN 0 S.NO========================
		 * sheet1.setColumnView(col, widthInChars); sheet1.addCell(new Label(col,
		 * newrow, Integer.toString(newrow), cellFormat1)); // =================COLUMN 1
		 * SCENARIO========================= System.out.println("In Column1 Scenario");
		 * col = 1; widthInChars = 70; sheet1.setColumnView(col, widthInChars);
		 */
		// System.out.println("get content: "+ sheet.getCell(1,
		// 1).getContents().equalsIgnoreCase(scenarioName));
		if (newrow == 1) {
			// =================COLUMN 0 S.NO========================
			col = 0;
			widthInChars = 12;
			sheet1.setColumnView(col, widthInChars);
			sheet1.addCell(new Label(col, newrow, Integer.toString(newrow), cellFormat1));
			// =================COLUMN 1 SCENARIO=========================
			System.out.println("In Column1 Scenario");
			col = 1;
			widthInChars = 70;
			sheet1.setColumnView(col, widthInChars);
			sheet1.addCell(new Label(col, newrow, scenarioName, cellFormat1));
			if (status) { // if2
				if (browserName.get().equalsIgnoreCase("Chrome")) {
					col = 3;
					widthInChars = 15;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
				} else if (browserName.get().equalsIgnoreCase("internet explorer")) {
					col = 4;
					widthInChars = 15;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
				} else if (browserName.get().equalsIgnoreCase("MicrosoftEdge")) {
					col = 5;
					widthInChars = 15;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
				} else {
					System.out.println("Invalid Browser while creating Excel Report");
				}
				col = 7;
				widthInChars = 20;
				sheet1.setColumnView(col, widthInChars);
				sheet1.addCell(new Label(col, newrow, "", cellFormat1));
				comments = "";
			} else if (!status) {
				if (browserName.get().equalsIgnoreCase("Chrome")) {
					col = 3;
					widthInChars = 15;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
				} else if (browserName.get().equalsIgnoreCase("internet explorer")) {
					col = 4;
					widthInChars = 15;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
				} else if (browserName.get().equalsIgnoreCase("MicrosoftEdge")) {
					col = 5;
					widthInChars = 15;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
				}
				col = 6;
				widthInChars = 20;
				sheet1.setColumnView(col, widthInChars);
				sheet1.addCell(new Label(col, newrow, ErrType, cellFormat1));
				col = 7;
				widthInChars = 20;
				sheet1.setColumnView(col, widthInChars);
				sheet1.addCell(new Label(col, newrow, comments, cellFormat1));
			}
		} else {
			int duplicateScenarioRow = 0;
			for (int i = 1; i < newrow; i++) {
				if (sheet.getCell(1, i).getContents().equalsIgnoreCase(scenarioName)) {
					duplicateScenarioRow = i;
				}
			}
			System.out.println("duplicateScenarioRow: " + duplicateScenarioRow);
			if (duplicateScenarioRow == 0) {
				// =================COLUMN 0 S.NO========================
				col = 0;
				widthInChars = 12;
				sheet1.setColumnView(col, widthInChars);
				sheet1.addCell(new Label(col, newrow, Integer.toString(newrow), cellFormat1));
				// =================COLUMN 1 SCENARIO=========================
				col = 1;
				widthInChars = 70;
				sheet1.setColumnView(col, widthInChars);
				sheet1.addCell(new Label(col, newrow, scenarioName, cellFormat1));
				if (status) { // if2
					if (browserName.get().equalsIgnoreCase("Chrome")) {
						col = 3;
						widthInChars = 15;
						sheet1.setColumnView(col, widthInChars);
						sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
					} else if (browserName.get().equalsIgnoreCase("internet explorer")) {
						col = 4;
						widthInChars = 15;
						sheet1.setColumnView(col, widthInChars);
						sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
					} else if (browserName.get().equalsIgnoreCase("MicrosoftEdge")) {
						col = 5;
						widthInChars = 15;
						sheet1.setColumnView(col, widthInChars);
						sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
					}
					col = 7;
					widthInChars = 20;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, "", cellFormat1));
					comments = "";
				} else if (!status) {
					if (browserName.get().equalsIgnoreCase("Chrome")) {
						col = 3;
						widthInChars = 15;
						sheet1.setColumnView(col, widthInChars);
						sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
					} else if (browserName.get().equalsIgnoreCase("internet explorer")) {
						col = 4;
						widthInChars = 15;
						sheet1.setColumnView(col, widthInChars);
						sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
					} else if (browserName.get().equalsIgnoreCase("MicrosoftEdge")) {
						col = 5;
						widthInChars = 15;
						sheet1.setColumnView(col, widthInChars);
						sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
					}
					col = 6;
					widthInChars = 20;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, ErrType, cellFormat1));
					col = 7;
					widthInChars = 20;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, newrow, comments, cellFormat1));
				}
			} else {
				if (status) { // if2
					if (browserName.get().equalsIgnoreCase("Chrome")) {
						col = 3;
						widthInChars = 15;
						sheet1.setColumnView(col, widthInChars);
						sheet1.addCell(new Label(col, duplicateScenarioRow, "PASS", passcellFormat));
					} else if (browserName.get().equalsIgnoreCase("internet explorer")) {
						col = 4;
						widthInChars = 15;
						sheet1.setColumnView(col, widthInChars);
						sheet1.addCell(new Label(col, duplicateScenarioRow, "PASS", passcellFormat));
					} else if (browserName.get().equalsIgnoreCase("MicrosoftEdge")) {
						col = 5;
						widthInChars = 15;
						sheet1.setColumnView(col, widthInChars);
						sheet1.addCell(new Label(col, duplicateScenarioRow, "PASS", passcellFormat));
					}
					col = 7;
					widthInChars = 20;
					sheet1.setColumnView(col, widthInChars);
					sheet1.addCell(new Label(col, duplicateScenarioRow, "", cellFormat1));
					comments = "";
				} else if (!status) {
					if (browserName.get().equalsIgnoreCase("Chrome")) {
						col = 3;
						widthInChars = 15;

						sheet1.setColumnView(col, widthInChars);

						sheet1.addCell(new Label(col, duplicateScenarioRow, "FAIL", failcellFormat));

					} else if (browserName.get().equalsIgnoreCase("internet explorer")) {

						col = 4;

						widthInChars = 15;

						sheet1.setColumnView(col, widthInChars);

						sheet1.addCell(new Label(col, duplicateScenarioRow, "FAIL", failcellFormat));

					} else if (browserName.get().equalsIgnoreCase("MicrosoftEdge")) {

						col = 5;

						widthInChars = 15;

						sheet1.setColumnView(col, widthInChars);

						sheet1.addCell(new Label(col, duplicateScenarioRow, "FAIL", failcellFormat));

					}

					col = 6;

					widthInChars = 20;

					sheet1.setColumnView(col, widthInChars);

					sheet1.addCell(new Label(col, duplicateScenarioRow, ErrType, cellFormat1));

					col = 7;

					widthInChars = 20;

					sheet1.setColumnView(col, widthInChars);

					sheet1.addCell(new Label(col, duplicateScenarioRow, comments, cellFormat1));

				}

			}

		}

		// =================COLUMN 3,4,5,6,7 SCENARIO=========================

		System.out.println("In Column 3,4,5,6,7 Scenario");

		/*
		 * 
		 * if (status) { // if2 if (browserName.get().equalsIgnoreCase("Chrome")) { col
		 * 
		 * = 3; widthInChars = 15; sheet1.setColumnView(col, widthInChars);
		 * 
		 * sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat)); } else if
		 * 
		 * (browserName.get().equalsIgnoreCase("internet explorer")) { col = 4;
		 * 
		 * widthInChars = 15; sheet1.setColumnView(col, widthInChars);
		 * 
		 * sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat)); } else if
		 * 
		 * (browserName.get().equalsIgnoreCase("MicrosoftEdge")) { col = 5; widthInChars
		 * 
		 * = 15; sheet1.setColumnView(col, widthInChars); sheet1.addCell(new Label(col,
		 * 
		 * newrow, "PASS", passcellFormat)); } col = 7; widthInChars = 20;
		 * 
		 * sheet1.setColumnView(col, widthInChars); sheet1.addCell(new Label(col,
		 * 
		 * newrow, "", cellFormat1)); comments = ""; } else if (!status){ if
		 * 
		 * (browserName.get().equalsIgnoreCase("Chrome")) { col = 3; widthInChars = 15;
		 * 
		 * sheet1.setColumnView(col, widthInChars); sheet1.addCell(new Label(col,
		 * 
		 * newrow, "FAIL", failcellFormat)); } else if
		 * 
		 * (browserName.get().equalsIgnoreCase("internet explorer")) { col = 4;
		 * 
		 * widthInChars = 15; sheet1.setColumnView(col, widthInChars);
		 * 
		 * sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat)); } else if
		 * 
		 * (browserName.get().equalsIgnoreCase("MicrosoftEdge")) { col = 5; widthInChars
		 * 
		 * = 15; sheet1.setColumnView(col, widthInChars); sheet1.addCell(new Label(col,
		 * 
		 * newrow, "FAIL", failcellFormat)); } col = 6; widthInChars = 20;
		 * 
		 * sheet1.setColumnView(col, widthInChars); sheet1.addCell(new Label(col,
		 * 
		 * newrow, ErrType, cellFormat1)); col = 7; widthInChars = 20;
		 * 
		 * sheet1.setColumnView(col, widthInChars); sheet1.addCell(new Label(col,
		 * 
		 * newrow, comments, cellFormat1));
		 *
		 * 
		 * 
		 * }
		 * 
		 */ // if2 ends

		wb1.close();

		wbcopy.write();

		wbcopy.close();

		logfile.delete();

		Workbook wb2 = Workbook.getWorkbook(new File(ofilepath));

		WritableWorkbook wbmain = Workbook.createWorkbook(new File(filepath), wb2);

		WritableSheet sheet2 = wbcopy.getSheet(0);

		sheet2.setName("Report");

		wbmain.write();

		wbmain.close();

		new File(ofilepath).delete();

	}

	/*
	 * public static synchronized void reportExtent(Boolean status, String Cmnt,
	 * String scenarioName, String browserName,
	 * 
	 * String testCaseException) throws IOException, RowsExceededException,
	 * WriteException, BiffException {
	 * 
	 * 
	 * 
	 * System.out.println("After Report - Status - " + status + " Browser: " +
	 * 
	 * browserName + " scenarioName - " + scenarioName + " Thread: " +
	 * 
	 * Thread.currentThread().getId());
	 * 
	 * 
	 * 
	 * if (status == false) {
	 * 
	 * // comments = Cmnt;
	 * 
	 * comments = "";
	 * 
	 * // System.out.println("comments " + comments);
	 * 
	 * } else {
	 * 
	 * comments = "";
	 * 
	 * }
	 * 
	 * Calendar cal = Calendar.getInstance();
	 * 
	 * DateFormat dateFormat = new SimpleDateFormat("MM_dd_yyyy");
	 * 
	 * String cal1 = dateFormat.format(cal.getTime());
	 * 
	 * // String currentDir = System.getProperty("user.dir");
	 * 
	 * // String filepath = currentDir + File.separator + "Results\\" + appName + "
	 * Final Report // " + cal1 + ".xls";
	 * 
	 * // File ifilepath = new File(currentDir + File.separator +
	 * "Results\\" + appName + " Final // Report " + cal1 + ".xls");
	 * 
	 * // String ofilepath = currentDir + File.separator + "Results\\" + appName + "
	 * Final Report // " + cal1 + "_temp.xls";
	 * 
	 * // File logfile = new File(filepath);// Created object of java File
	 * 
	 * if (ITestListenerImpl.defaultConfigProperty.get().getProperty("CICD").
	 * equalsIgnoreCase("N")) {
	 * 
	 * currentDirExcel = ITestListenerImpl.defaultConfigProperty.get().getProperty(
	 * "latestReportLocation") + File.separator
	 * 
	 * + File.separator + "Reports_" + cal1;
	 * 
	 * } else {
	 * 
	 * currentDirExcel = ITestListenerImpl.defaultConfigProperty.get().getProperty(
	 * "CICDlatestReportLocation")
	 * 
	 * + File.separator + File.separator + "Reports_" + cal1;
	 * 
	 * }
	 * 
	 * String filepath = currentDirExcel + File.separator + "Final Report " + cal1 +
	 * ".xls";
	 * 
	 * File ifilepath = new File(currentDirExcel + File.separator + "Final Report "
	 * + cal1 + ".xls");
	 * 
	 * String ofilepath = currentDirExcel + File.separator + "Final Report " + cal1
	 * + "_temp.xls";
	 * 
	 * File logfile = new File(filepath);// Created object of java File
	 * 
	 * if (!logfile.exists()) { // if1
	 * 
	 * WritableWorkbook workbook = Workbook.createWorkbook(new File(filepath));
	 * 
	 * WritableSheet sheet = workbook.createSheet("Report", 0);
	 * 
	 * sheet.setName("Report");
	 * 
	 * WritableFont arialfont = new WritableFont(WritableFont.ARIAL, 11,
	 * WritableFont.BOLD);
	 * 
	 * WritableCellFormat cellFormat = new WritableCellFormat(arialfont);
	 * 
	 * cellFormat.setBackground(Colour.ICE_BLUE);
	 * 
	 * cellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
	 * 
	 * sheet.addCell(new Label(0, 0, "S.No", cellFormat));
	 * 
	 * sheet.addCell(new Label(1, 0, "Test Case", cellFormat));
	 * 
	 * sheet.addCell(new Label(2, 0, "Overall Status", cellFormat));
	 * 
	 * sheet.addCell(new Label(3, 0, "Chrome", cellFormat));
	 * 
	 * sheet.addCell(new Label(4, 0, " Edge ", cellFormat));
	 * 
	 * // sheet.addCell(new Label(6, 0, "Fail Reason", cellFormat));
	 * 
	 * sheet.addCell(new Label(5, 0, "Comments", cellFormat));
	 * 
	 * workbook.write();
	 * 
	 * workbook.close();
	 * 
	 * } // if1 ends
	 * 
	 * Workbook wb1 = Workbook.getWorkbook(ifilepath);
	 * 
	 * WritableWorkbook wbcopy = Workbook.createWorkbook(new File(ofilepath), wb1);
	 * 
	 * WritableSheet sheet1 = wbcopy.getSheet(0);
	 * 
	 * Sheet sheet = wb1.getSheet(0);
	 * 
	 * int newrow = sheet.getRows();
	 * 
	 * // System.out.println("newrow: " + newrow);
	 * 
	 * sheet1.setName("Report");
	 * 
	 * WritableFont arialfont1 = new WritableFont(WritableFont.ARIAL, 10);
	 * 
	 * WritableCellFormat cellFormat1 = new WritableCellFormat(arialfont1);
	 * 
	 * cellFormat1.setBorder(Border.ALL, BorderLineStyle.THIN);
	 * 
	 * WritableCellFormat passcellFormat = new WritableCellFormat(arialfont1);
	 * 
	 * passcellFormat.setBackground(Colour.LIGHT_GREEN);
	 * 
	 * passcellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
	 * 
	 * WritableCellFormat failcellFormat = new WritableCellFormat(arialfont1);
	 * 
	 * failcellFormat.setBackground(Colour.RED);
	 * 
	 * failcellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
	 * 
	 * int col = 0;
	 * 
	 * int widthInChars = 12;
	 * 
	 * 
	 * 
	 * // =================COLUMN 0 S.NO========================
	 * 
	 * sheet1.setColumnView(col, widthInChars); sheet1.addCell(new Label(col,
	 * 
	 * newrow, Integer.toString(newrow), cellFormat1)); // =================COLUMN 1
	 * 
	 * SCENARIO========================= System.out.println("In Column1 Scenario");
	 * 
	 * col = 1; widthInChars = 70; sheet1.setColumnView(col, widthInChars);
	 * 
	 * 
	 * 
	 * // System.out.println("get content: "+ sheet.getCell(1,
	 * 
	 * // 1).getContents().equalsIgnoreCase(scenarioName));
	 * 
	 * if (newrow == 1) {
	 * 
	 * // =================COLUMN 0 S.NO========================
	 * 
	 * col = 0;
	 * 
	 * widthInChars = 12;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, Integer.toString(newrow),
	 * cellFormat1));
	 * 
	 * // =================COLUMN 1 SCENARIO=========================
	 * 
	 * // System.out.println("In Column1 Scenario");
	 * 
	 * col = 1;
	 * 
	 * widthInChars = 70;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, scenarioName, cellFormat1));
	 * 
	 * if (status) { // if2
	 * 
	 * if (browserName.equalsIgnoreCase("Chrome")) {
	 * 
	 * col = 3;
	 * 
	 * widthInChars = 15;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
	 * 
	 * } else if (browserName.equalsIgnoreCase("MicrosoftEdge")) {
	 * 
	 * col = 4;
	 * 
	 * widthInChars = 15;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
	 * 
	 * } else {
	 * 
	 * System.out.println("Invalid Browser while creating Excel Report");
	 * 
	 * }
	 * 
	 * col = 5;
	 * 
	 * widthInChars = 20;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, "", cellFormat1));
	 * 
	 * comments = "";
	 * 
	 * } else if (!status) {
	 * 
	 * if (browserName.equalsIgnoreCase("Chrome")) {
	 * 
	 * col = 3;
	 * 
	 * widthInChars = 15;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
	 * 
	 * } else if (browserName.equalsIgnoreCase("MicrosoftEdge")) {
	 * 
	 * col = 4;
	 * 
	 * widthInChars = 15;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
	 * 
	 * }
	 * 
	 * col = 5;
	 * 
	 * widthInChars = 20;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, comments, cellFormat1));
	 * 
	 * 
	 * 
	 * col = 7; widthInChars = 20; sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, testCaseException, cellFormat1));
	 * 
	 * 
	 * 
	 * }
	 * 
	 * } else {
	 * 
	 * int duplicateScenarioRow = 0;
	 * 
	 * for (int i = 1; i < newrow; i++) {
	 * 
	 * if (sheet.getCell(1, i).getContents().equalsIgnoreCase(scenarioName)) {
	 * 
	 * duplicateScenarioRow = i;
	 * 
	 * }
	 * 
	 * }
	 * 
	 * // System.out.println("duplicateScenarioRow: " + duplicateScenarioRow);
	 * 
	 * if (duplicateScenarioRow == 0) {
	 * 
	 * // =================COLUMN 0 S.NO========================
	 * 
	 * col = 0;
	 * 
	 * widthInChars = 12;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, Integer.toString(newrow),
	 * cellFormat1));
	 * 
	 * // =================COLUMN 1 SCENARIO=========================
	 * 
	 * col = 1;
	 * 
	 * widthInChars = 70;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, scenarioName, cellFormat1));
	 * 
	 * if (status) { // if2
	 * 
	 * if (browserName.equalsIgnoreCase("Chrome")) {
	 * 
	 * col = 3;
	 * 
	 * widthInChars = 15;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
	 * 
	 * } else if (browserName.equalsIgnoreCase("MicrosoftEdge")) {
	 * 
	 * col = 4;
	 * 
	 * widthInChars = 15;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
	 * 
	 * }
	 * 
	 * col = 5;
	 * 
	 * widthInChars = 20;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, "", cellFormat1));
	 * 
	 * comments = "";
	 * 
	 * } else if (!status) {
	 * 
	 * if (browserName.equalsIgnoreCase("Chrome")) {
	 * 
	 * col = 3;
	 * 
	 * widthInChars = 15;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
	 * 
	 * } else if (browserName.equalsIgnoreCase("MicrosoftEdge")) {
	 * 
	 * col = 4;
	 * 
	 * widthInChars = 15;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
	 * 
	 * }
	 * 
	 * col = 5;
	 * 
	 * widthInChars = 20;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, comments, cellFormat1));
	 * 
	 * 
	 * 
	 * col = 7; widthInChars = 20; sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, newrow, testCaseException, cellFormat1));
	 * 
	 * 
	 * 
	 * }
	 * 
	 * } else {
	 * 
	 * if (status) { // if2
	 * 
	 * if (browserName.equalsIgnoreCase("Chrome")) {
	 * 
	 * col = 3;
	 * 
	 * widthInChars = 15;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, duplicateScenarioRow, "PASS", passcellFormat));
	 * 
	 * } else if (browserName.equalsIgnoreCase("MicrosoftEdge")) {
	 * 
	 * col = 4;
	 * 
	 * widthInChars = 15;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, duplicateScenarioRow, "PASS", passcellFormat));
	 * 
	 * }
	 * 
	 * col = 5;
	 * 
	 * widthInChars = 20;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, duplicateScenarioRow, "", cellFormat1));
	 * 
	 * comments = "";
	 * 
	 * } else if (!status) {
	 * 
	 * if (browserName.equalsIgnoreCase("CHROME")) {
	 * 
	 * col = 3;
	 * 
	 * widthInChars = 15;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, duplicateScenarioRow, "FAIL", failcellFormat));
	 * 
	 * } else if (browserName.equalsIgnoreCase("MICROSOFTEDGE")) {
	 * 
	 * col = 4;
	 * 
	 * widthInChars = 15;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, duplicateScenarioRow, "FAIL", failcellFormat));
	 * 
	 * }
	 * 
	 * col = 5;
	 * 
	 * widthInChars = 20;
	 * 
	 * sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, duplicateScenarioRow, comments, cellFormat1));
	 * 
	 * col = 6;
	 * 
	 * 
	 * 
	 * widthInChars = 20; sheet1.setColumnView(col, widthInChars);
	 * 
	 * sheet1.addCell(new Label(col, duplicateScenarioRow, testCaseException,
	 * 
	 * cellFormat1));
	 * 
	 * 
	 * 
	 * }
	 * 
	 * }
	 * 
	 * }
	 * 
	 * // =================COLUMN 3,4,5,6,7 SCENARIO=========================
	 * 
	 * // System.out.println("In Column 3,4,5,6,7 Scenario");
	 * 
	 * wb1.close();
	 * 
	 * wbcopy.write();
	 * 
	 * wbcopy.close();
	 * 
	 * logfile.delete();
	 * 
	 * Workbook wb2 = Workbook.getWorkbook(new File(ofilepath));
	 * 
	 * WritableWorkbook wbmain = Workbook.createWorkbook(new File(filepath), wb2);
	 * 
	 * WritableSheet sheet2 = wbcopy.getSheet(0);
	 * 
	 * sheet2.setName("Report");
	 * 
	 * wbmain.write();
	 * 
	 * wbmain.close();
	 * 
	 * new File(ofilepath).delete();
	 * 
	 * }
	 */
	
	public static synchronized void reportExtent(Boolean status, String Cmnt, String scenarioName, String browserName,
	        String testCaseException) throws IOException, RowsExceededException, WriteException, BiffException {

	    if (!status) {
	        comments = "";
	    } else {
	        comments = "";
	    }

	    Calendar cal = Calendar.getInstance();
	    DateFormat dateFormat = new SimpleDateFormat("MM_dd_yyyy");
	    String cal1 = dateFormat.format(cal.getTime());

	    if (ITestListenerImpl.defaultConfigProperty.get().getProperty("CICD").equalsIgnoreCase("N")) {
	        currentDirExcel = ITestListenerImpl.defaultConfigProperty.get().getProperty("latestReportLocation")
	                + File.separator + "Reports_" + cal1;
	    } else {
	        currentDirExcel = ITestListenerImpl.defaultConfigProperty.get().getProperty("CICDlatestReportLocation")
	                + File.separator + "Reports_" + cal1;
	    }

	    String filepath = currentDirExcel + File.separator + "Final Report " + cal1 + ".xls";
	    File ifilepath = new File(filepath);
	    String ofilepath = currentDirExcel + File.separator + "Final Report " + cal1 + "_temp.xls";
	    File logfile = new File(filepath);

	    if (!logfile.exists()) {
	        WritableWorkbook workbook = Workbook.createWorkbook(new File(filepath));
	        WritableSheet sheet = workbook.createSheet("Report", 0);
	        sheet.setName("Report");

	        WritableFont arialfont = new WritableFont(WritableFont.ARIAL, 11, WritableFont.BOLD);
	        WritableCellFormat cellFormat = new WritableCellFormat(arialfont);
	        cellFormat.setBackground(Colour.ICE_BLUE);
	        cellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

	        sheet.addCell(new Label(0, 0, "S.No", cellFormat));
	        sheet.addCell(new Label(1, 0, "Test Case", cellFormat));
	        sheet.addCell(new Label(2, 0, "Overall Status", cellFormat));
	        sheet.addCell(new Label(3, 0, "Chrome", cellFormat));
	        sheet.addCell(new Label(4, 0, "Firefox", cellFormat));
	        sheet.addCell(new Label(5, 0, "Comments", cellFormat));

	        workbook.write();
	        workbook.close();
	    }

	    Workbook wb1 = Workbook.getWorkbook(ifilepath);
	    WritableWorkbook wbcopy = Workbook.createWorkbook(new File(ofilepath), wb1);
	    WritableSheet sheet1 = wbcopy.getSheet(0);
	    Sheet sheet = wb1.getSheet(0);
	    int newrow = sheet.getRows();
	    sheet1.setName("Report");

	    WritableFont arialfont1 = new WritableFont(WritableFont.ARIAL, 10);
	    WritableCellFormat cellFormat1 = new WritableCellFormat(arialfont1);
	    cellFormat1.setBorder(Border.ALL, BorderLineStyle.THIN);

	    WritableCellFormat passcellFormat = new WritableCellFormat(arialfont1);
	    passcellFormat.setBackground(Colour.LIGHT_GREEN);
	    passcellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

	    WritableCellFormat failcellFormat = new WritableCellFormat(arialfont1);
	    failcellFormat.setBackground(Colour.RED);
	    failcellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

	    int col;
	    int widthInChars;

	    if (newrow == 1) {
	        col = 0;
	        widthInChars = 12;
	        sheet1.setColumnView(col, widthInChars);
	        sheet1.addCell(new Label(col, newrow, Integer.toString(newrow), cellFormat1));

	        col = 1;
	        widthInChars = 70;
	        sheet1.setColumnView(col, widthInChars);
	        sheet1.addCell(new Label(col, newrow, scenarioName, cellFormat1));

	        if (status) {
	            if (browserName.equalsIgnoreCase("Chrome")) {
	                col = 3;
	                widthInChars = 15;
	                sheet1.setColumnView(col, widthInChars);
	                sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
	            } else if (browserName.equalsIgnoreCase("Firefox")) {
	                col = 4;
	                widthInChars = 15;
	                sheet1.setColumnView(col, widthInChars);
	                sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
	            } else {
	                System.out.println("Invalid Browser while creating Excel Report");
	            }
	            col = 5;
	            widthInChars = 20;
	            sheet1.setColumnView(col, widthInChars);
	            sheet1.addCell(new Label(col, newrow, "", cellFormat1));
	            comments = "";
	        } else {
	            if (browserName.equalsIgnoreCase("Chrome")) {
	                col = 3;
	                widthInChars = 15;
	                sheet1.setColumnView(col, widthInChars);
	                sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
	            } else if (browserName.equalsIgnoreCase("Firefox")) {
	                col = 4;
	                widthInChars = 15;
	                sheet1.setColumnView(col, widthInChars);
	                sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
	            }
	            col = 5;
	            widthInChars = 20;
	            sheet1.setColumnView(col, widthInChars);
	            sheet1.addCell(new Label(col, newrow, comments, cellFormat1));
	        }
	    } else {
	        int duplicateScenarioRow = 0;
	        for (int i = 1; i < newrow; i++) {
	            if (sheet.getCell(1, i).getContents().equalsIgnoreCase(scenarioName)) {
	                duplicateScenarioRow = i;
	                break;
	            }
	        }

	        if (duplicateScenarioRow == 0) {
	            col = 0;
	            widthInChars = 12;
	            sheet1.setColumnView(col, widthInChars);
	            sheet1.addCell(new Label(col, newrow, Integer.toString(newrow), cellFormat1));

	            col = 1;
	            widthInChars = 70;
	            sheet1.setColumnView(col, widthInChars);
	            sheet1.addCell(new Label(col, newrow, scenarioName, cellFormat1));

	            if (status) {
	                if (browserName.equalsIgnoreCase("Chrome")) {
	                    col = 3;
	                    widthInChars = 15;
	                    sheet1.setColumnView(col, widthInChars);
	                    sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
	                } else if (browserName.equalsIgnoreCase("Firefox")) {
	                    col = 4;
	                    widthInChars = 15;
	                    sheet1.setColumnView(col, widthInChars);
	                    sheet1.addCell(new Label(col, newrow, "PASS", passcellFormat));
	                }
	                col = 5;
	                widthInChars = 20;
	                sheet1.setColumnView(col, widthInChars);
	                sheet1.addCell(new Label(col, newrow, "", cellFormat1));
	                comments = "";
	            } else {
	                if (browserName.equalsIgnoreCase("Chrome")) {
	                    col = 3;
	                    widthInChars = 15;
	                    sheet1.setColumnView(col, widthInChars);
	                    sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
	                } else if (browserName.equalsIgnoreCase("Firefox")) {
	                    col = 4;
	                    widthInChars = 15;
	                    sheet1.setColumnView(col, widthInChars);
	                    sheet1.addCell(new Label(col, newrow, "FAIL", failcellFormat));
	                }
	                col = 5;
	                widthInChars = 20;
	                sheet1.setColumnView(col, widthInChars);
	                sheet1.addCell(new Label(col, newrow, comments, cellFormat1));
	            }
	        } else {
	            if (status) {
	                if (browserName.equalsIgnoreCase("Chrome")) {
	                    col = 3;
	                    widthInChars = 15;
	                    sheet1.setColumnView(col, widthInChars);
	                    sheet1.addCell(new Label(col, duplicateScenarioRow, "PASS", passcellFormat));
	                } else if (browserName.equalsIgnoreCase("Firefox")) {
	                    col = 4;
	                    widthInChars = 15;
	                    sheet1.setColumnView(col, widthInChars);
	                    sheet1.addCell(new Label(col, duplicateScenarioRow, "PASS", passcellFormat));
	                }
	                col = 5;
	                widthInChars = 20;
	                sheet1.setColumnView(col, widthInChars);
	                sheet1.addCell(new Label(col, duplicateScenarioRow, "", cellFormat1));
	                comments = "";
	            } else {
	                if (browserName.equalsIgnoreCase("Chrome")) {
	                    col = 3;
	                    widthInChars = 15;
	                    sheet1.setColumnView(col, widthInChars);
	                    sheet1.addCell(new Label(col, duplicateScenarioRow, "FAIL", failcellFormat));
	                } else if (browserName.equalsIgnoreCase("Firefox")) {
	                    col = 4;
	                    widthInChars = 15;
	                    sheet1.setColumnView(col, widthInChars);
	                    sheet1.addCell(new Label(col, duplicateScenarioRow, "FAIL", failcellFormat));
	                }
	                col = 5;
	                widthInChars = 20;
	                sheet1.setColumnView(col, widthInChars);
	                sheet1.addCell(new Label(col, duplicateScenarioRow, comments, cellFormat1));
	            }
	        }
	    }

	    wb1.close();
	    wbcopy.write();
	    wbcopy.close();

	    logfile.delete();

	    Workbook wb2 = Workbook.getWorkbook(new File(ofilepath));
	    WritableWorkbook wbmain = Workbook.createWorkbook(new File(filepath), wb2);
	    WritableSheet sheet2 = wbcopy.getSheet(0);
	    sheet2.setName("Report");
	    wbmain.write();
	    wbmain.close();
	    new File(ofilepath).delete();
	}

	public static synchronized boolean overallStatus()

			throws IOException, RowsExceededException, WriteException, BiffException {

		System.out.println("==========>Updating Overall Status in Excel Report<==========");

		Calendar cal = Calendar.getInstance();

		DateFormat dateFormat = new SimpleDateFormat("MM_dd_yyyy");

		String cal1 = dateFormat.format(cal.getTime());

		// String currentDir = System.getProperty("user.dir");

		String currentDir;

		if (ITestListenerImpl.defaultConfigProperty.get().getProperty("CICD").equalsIgnoreCase("N")) {

			currentDir = ITestListenerImpl.defaultConfigProperty.get().getProperty("latestReportLocation")
					+ File.separator

					+ File.separator + "Reports_" + cal1;

		} else {

			currentDir = ITestListenerImpl.defaultConfigProperty.get().getProperty("CICDlatestReportLocation")
					+ File.separator

					+ File.separator + "Reports_" + cal1;

		}

		/*
		 * 
		 * String currentDir ="C:\\Users\\" + System.getProperty("user.name") +
		 * 
		 * [file://OneDrive%20-%20St.%20James's%20Place/Desktop/NextGenBDD_Results/]\\
		 * OneDrive - St. James's Place\\Desktop\\NextGenBDD_Results\\ + appName + "
		 * 
		 * \\" + "Detailed_Report " + cal1 + File.separator;
		 * 
		 */

		// String filepath = currentDir + File.separator + "Results\\" + appName + "
		// Final Report
		// " + cal1 + ".xls";

		// File ifilepath = new File(currentDir + File.separator + "Results\\" + appName
		// + " Final
		// Report " + cal1 + ".xls");

		// String ofilepath = currentDir + File.separator + "Results\\" + appName + "
		// Final Report
		// " + cal1 + "_temp.xls";

		// File logfile = new File(filepath);// Created object of java File

		String filepath = currentDir + File.separator + "Final Report " + cal1 + ".xls";

		File ifilepath = new File(currentDir + File.separator + "Final Report " + cal1 + ".xls");

		String ofilepath = currentDir + File.separator + "Final Report " + cal1 + "_temp.xls";

		File logfile = new File(filepath);// Created object of java File

		if (!logfile.exists()) {

			System.out.println("Excel File not Present : " + filepath);

			return false;

		} // if1 ends

		else {

			Workbook wb1 = Workbook.getWorkbook(ifilepath);

			WritableWorkbook wbcopy = Workbook.createWorkbook(new File(ofilepath), wb1);

			WritableSheet sheet1 = wbcopy.getSheet(0);

			Sheet sheet = wb1.getSheet(0);

			int newrow = sheet.getRows();

			// System.out.println("TotalRows: " + newrow);

			WritableFont arialfont1 = new WritableFont(WritableFont.ARIAL, 10);

			WritableCellFormat cellFormat1 = new WritableCellFormat(arialfont1);

			cellFormat1.setBorder(Border.ALL, BorderLineStyle.THIN);

			WritableCellFormat passcellFormat = new WritableCellFormat(arialfont1);

			passcellFormat.setBackground(Colour.LIGHT_GREEN);

			passcellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

			WritableCellFormat failcellFormat = new WritableCellFormat(arialfont1);

			failcellFormat.setBackground(Colour.RED);

			failcellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);

			int col = 2;

			int widthInChars = 20;

			sheet1.setColumnView(col, widthInChars);

			for (int i = 1; i < newrow; i++) {

				String chromeStatus = sheet.getCell(3, i).getContents();

				String edgeStatus = sheet.getCell(4, i).getContents();

				if (!(chromeStatus.equalsIgnoreCase("FAIL") || chromeStatus.equalsIgnoreCase("PASS"))) {

					// =================COLUMN 3 CHROME NA=======================

					sheet1.addCell(new Label(3, i, "NA", cellFormat1));

				}

				if (!(edgeStatus.equalsIgnoreCase("FAIL") || edgeStatus.equalsIgnoreCase("PASS"))) {

					// =================COLUMN 4 EDGE NA=======================

					sheet1.addCell(new Label(4, i, "NA", cellFormat1));

				}

			}

			for (int i = 1; i < newrow; i++) {

				String chromeStatus = sheet.getCell(3, i).getContents();

				String edgeStatus = sheet.getCell(4, i).getContents();

				if (chromeStatus.equalsIgnoreCase("FAIL") || edgeStatus.equalsIgnoreCase("FAIL")) {

					// =================COLUMN 2 STATUS=======================

					sheet1.addCell(new Label(col, i, "FAILED", failcellFormat));

				} else {

					sheet1.addCell(new Label(col, i, "PASSED", passcellFormat));

				}

			}

			wb1.close();

			wbcopy.write();

			wbcopy.close();

			logfile.delete();

			Workbook wb2 = Workbook.getWorkbook(new File(ofilepath));

			WritableWorkbook wbmain = Workbook.createWorkbook(new File(filepath), wb2);

			WritableSheet sheet2 = wbcopy.getSheet(0);

			sheet2.setName("Report");

			wbmain.write();

			wbmain.close();

			new File(ofilepath).delete();

			return true;

		}

	}

	public static void reportDel() throws IOException {

		// String currentDir = System.getProperty("user.dir");

		String currentDir;

		if (ITestListenerImpl.defaultConfigProperty.get().getProperty("CICD").equalsIgnoreCase("N")) {

			currentDir = ITestListenerImpl.defaultConfigProperty.get().getProperty("latestReportLocation");

		} else {

			currentDir = ITestListenerImpl.defaultConfigProperty.get().getProperty("CICDlatestReportLocation");

		}

		File file = new File(currentDir);

		/*
		 * 
		 * Properties prop = new Properties(); FileInputStream fis = new
		 * 
		 * FileInputStream( System.getProperty("user.dir") + \\OR.properties);
		 * 
		 * prop.load(fis);
		 * 
		 */

		// report delete

		try {

			if (ITestListenerImpl.defaultConfigProperty.get().getProperty("reportAppend").equalsIgnoreCase("N")) {

				System.out.println("Existing reports will get deleted");

				if (file.exists()) {

					FileUtils.cleanDirectory(file);

				} else {

					FileUtils.forceMkdir(file);

				}

				// System.out.println("Existing reports will get deleted");

				// String filepath = currentDir + File.separator + "Results\\";

				// File file = new File(filepath);

				// String[] myFiles;

				// if (file.exists()) {

				// if (file.isDirectory()) {

				// myFiles = file.list();

				// for (int i = 0; i < myFiles.length; i++) {

				//

				// File myFile = new File(file, myFiles[i]);

				// /*

				// * System.out.println("Absolute File Path " + myFile.getAbsolutePath());

				// * System.out.println("File[" + i + "] : " + myFiles[i]);

				// */

				// // if (myFiles[i].contains(appName)

				// if (myFiles[i].contains(appName) || myFiles[i].contains(".html")) {

				// // && myFiles[i].contains(cal1)) {

				// Runtime.getRuntime().exec("cmd /c taskkill /f /im excel.exe");

				// System.out.println(" Deleting File[" + i + "] : " + myFiles[i]);

				// myFile.delete();

				// }

				// }

				// }

				// }

			} else {

				System.out.println("Existing report will get appended");

			}

		} catch (Exception e) {

			// TODO Auto-generated catch block

			e.printStackTrace();

		}

	}

	public static String excelMailReport0() throws BiffException {

		String html;

		html = null;

		String UsrNme;

		String CertifyingAs;

		String color;

		String fontColor = "";

		int Total_Cases;

		try {

			Properties prop = new Properties();

			FileInputStream fis = new FileInputStream(
					System.getProperty("user.dir") + File.separator + "OR.properties");

			prop.load(fis);

			launchURL = prop.getProperty("url");

			BufferedWriter out;

			// Date

			Calendar cal = Calendar.getInstance();

			DateFormat dateFormat = new SimpleDateFormat("MM_dd_yyyy");

			String cal1 = dateFormat.format(cal.getTime());

			DateFormat dateFormat1 = new SimpleDateFormat("MM/dd/yyyy");

			String cal2 = dateFormat1.format(cal.getTime());

			String indexHtmlPath = System.getProperty("user.dir") + "Results" + File.separator + appName
					+ " Final Report " + cal1

					+ ".html";

			out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(indexHtmlPath), "UTF-8"));

			System.out.println("App name is : " + appName);

			String currentDir = System.getProperty("user.dir");

			String file_path = currentDir + File.separator + "Results" + File.separator + appName + " Final Report "
					+ cal1 + ".xls";

			// String filepath = "Results\\" + selectedApp + " Final Report " +

			// cal1 + ".xls";

			File actualFile = new File(file_path);

			System.out.println("Does actual file exists :" + actualFile.exists());

			System.out.println("Creating WorkBook");

			Workbook wb = Workbook.getWorkbook(new File(file_path));

			Sheet sh = wb.getSheet(0);

			int totalNoOfRows = sh.getRows();

			int totalNoOfCols = sh.getColumns();

			Total_Cases = totalNoOfRows - 1;

			System.out.println("Total Rows and cols " + totalNoOfRows + " " + totalNoOfCols);

			// Failed, Passed and Critical_Error_Counter and Critical_Counter

			System.out.println("10");

			for (int i = 1; i < totalNoOfRows; i++) {

				for (int j = 0; j < totalNoOfCols; j++) {

					if (sh.getCell(j, i).getContents().equalsIgnoreCase("PASSED")) {

						Passed_Case_Counter = Passed_Case_Counter + 1;

					} else if (sh.getCell(j, i).getContents().equalsIgnoreCase("FAILED")) {

						Failed_Case_Counter = Failed_Case_Counter + 1;

					} else if (sh.getCell(j, i).getContents().equalsIgnoreCase("CRITICAL")) {

						Critical_Error_Counter = Critical_Error_Counter + 1;

					}

					if (j == 3) {

						if (sh.getCell(j, i).getContents().equalsIgnoreCase("FAIL")

								&& sh.getCell(j + 1, i).getContents().equalsIgnoreCase("CRITICAL")) {

							Critical_Counter = Critical_Counter + 1;

						}

					}

				}

			}

			System.out.println("Failed, Passed and Critical_Error_Counter and Critical_Counter " + Failed_Case_Counter

					+ " " + Passed_Case_Counter + " " + Critical_Error_Counter + " " + Critical_Counter);

			// CertifyingAs

			if (Critical_Counter > 0) {

				CertifyingAs = "<font face='Calibri' size='12px' color='Red'><b>No-Go</b></font>";

			} else if (Failed_Case_Counter >= Passed_Case_Counter) {

				CertifyingAs = "<font face='Calibri' size='12px' color='Red'><b>No-Go</b></font>";

			} else {

				CertifyingAs = "<font face='Calibri' size='12px' color='Green'><b>Go</b></font>";

			}

			// Body Content

			/*
			 * 
			 * html = "<HTML><BODY> <font face='Calibri' size='12px'>Hello Team, " +
			 * 
			 * "<br><br>" + "Please find the " + appName + " Automated Test Report below" +
			 * 
			 * " <br> <br>" + "Certifying as " + CertifyingAs + "<br><br> URL: " + launchURL
			 * 
			 * + "<br> Date: " + cal2 + " </font><br><br>";
			 * 
			 */

			html = "<HTML><BODY> <font face='Calibri' size='10px'>Hello Team, " + "<br>" + "Please find the "

					+ appName + " Automated Test Report below" + "<br><br><b> URL: </b><b><a href>" + launchURL

					+ "</a></b><br><b> Date: </b>" + cal2 + " </font><br><br>";

			// Summary Table

			html += "<font face='Calibri' size='4.5'><b>SUMMARY</b></font><br><TABLE border='1' style='border-collapse:collapse' cellpadding '5' cellspacing='0'>";

			html += "<tr><TH bgcolor='#1A5276' align='left'>&nbsp<font face='calibri' size='3' color='White'>&nbsp<b>Total Cases Executed</b>&nbsp</font></th>";

			html += "<TH bgcolor='#1A5276' align='left'>&nbsp<font face='calibri' size='3' color='White'>&nbsp<b>Total Cases Passed</b>&nbsp</font></th>";

			html += "<TH bgcolor='#1A5276' align='left'>&nbsp<font face='calibri' size='3' color='White'>&nbsp<b>Total Cases Failed</b>&nbsp</font></th>";

			// html += "<TH bgcolor='#1A5276' align='left'>&nbsp<font face='calibri'

			// size='3' color='White'>&nbsp<b>Total Critical Errors</b>&nbsp</font></th>";

			html += "</tr>";

			html += "<tr align='center'><Td bgcolor='Whitesmoke'>&nbsp<font face='calibri' size='2.5'>&nbsp "

					+ Total_Cases + " &nbsp</font></td>";

			html += "<Td bgcolor='Whitesmoke'>&nbsp<font face='calibri' size='2.5'>&nbsp " + Passed_Case_Counter

					+ " &nbsp</font></td>";

			html += "<Td bgcolor='Whitesmoke'>&nbsp<font face='calibri' size='2.5'>&nbsp " + Failed_Case_Counter

					+ " &nbsp</font></td>";

			// html += "<Td bgcolor='Whitesmoke'>&nbsp<font face='calibri' size='2.5'>&nbsp

			// " + Critical_Error_Counter + " &nbsp</font></td>";

			html += "</tr></Table><br><br>";

			// Split-Up Table

			html += "<font face='Calibri' size='4.5'><b>SPLIT-UP</b></font><br><TABLE border='1' style='border-collapse:collapse' cellpadding'5' cellspacing='0'>";

			html += "<TR>";

			for (int i = 0; i < totalNoOfRows; i++) {

				for (int j = 0; j < totalNoOfCols; j++) {

					if (i == 0) {

						color = "#1A5276";

						html += "<TH bgcolor=" + color

								+ " align='center'>&nbsp<font face='calibri' size='3' color='White'><b>"

								+ sh.getCell(j, i).getContents().toString() + "&nbsp</b></font></th>";

					} else {

						if (sh.getCell(j, i).getContents().equalsIgnoreCase("PASS")) {

							color = "#00ed00";

						} else if (sh.getCell(j, i).getContents().equalsIgnoreCase("PASSED")) {

							color = "#00ed00";

							fontColor = "#000000";

						} else if (sh.getCell(j, i).getContents().equalsIgnoreCase("FAIL")) {

							color = "#ed0000";

						} else if (sh.getCell(j, i).getContents().equalsIgnoreCase("FAILED")) {

							color = "#ed0000";

							fontColor = "#000000";

						} else {

							color = "Whitesmoke";

						}

						if (j == 2) {

							html += "<Td align='center' bgcolor=" + color + ">&nbsp<font face='calibri' size='3' color="

									+ fontColor + ">&nbsp<b> " + sh.getCell(j, i).getContents().toString()

									+ "</b> &nbsp</font></td>";

						} else {

							html += "<Td align='left' bgcolor=" + color + ">&nbsp<font face='calibri' size='2.5'>&nbsp "

									+ sh.getCell(j, i).getContents().toString() + " &nbsp</font></td>";

						}

					}

				}

				html += "</TR>";

			}

			// Tail

			html += "</TABLE> <br><br><font face='Calibri' size='10px'><b>Note:</b><i>  Don't reply to all. This is a auto-generated report.</i></b><br> <br> Thanks & Regards <br>"

					+ User_Name;

			html += "<br>Any further queries please contact <b><a href='mailto:" + teamEmailID + "'>" + User_Name

					+ "</a></b><br>";

			html += "<br></Font></BODY></HTML>";

			wb.close();

			out.write(html.toString());

			out.close();

		} catch (Exception e) {

			e.printStackTrace();

			System.out.println("Some problem has occured" + e.getMessage());

		}

		System.out.println("returned html");

		return html;

	}

	public static String excelMailReport1() throws BiffException {

		String html;

		html = null;

		String UsrNme;

		String CertifyingAs;

		String color;

		int Total_Cases;

		try {

			Properties prop = new Properties();

			FileInputStream fis = new FileInputStream(
					System.getProperty("user.dir") + File.separator + "OR.properties");

			prop.load(fis);

			launchURL = prop.getProperty("url");

			BufferedWriter out;

			// Date

			System.out.println("8");

			Calendar cal = Calendar.getInstance();

			DateFormat dateFormat = new SimpleDateFormat("MM_dd_yyyy");

			String cal1 = dateFormat.format(cal.getTime());

			DateFormat dateFormat1 = new SimpleDateFormat("MM/dd/yyyy");

			String cal2 = dateFormat1.format(cal.getTime());

			String indexHtmlPath = System.getProperty("user.dir") + File.separator + "Results" + File.separator
					+ "MP_REPORT.html";

			out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(indexHtmlPath), "UTF-8"));

			System.out.println("9");

			System.out.println("App name is : " + appName);

			/*
			 * 
			 * String file_path = currentDir + "Results\\" + appName + " Final Report " +
			 * 
			 * cal1 + "//Results//reports.xls";
			 * 
			 */

			String currentDir = System.getProperty("user.dir");

			String file_path = currentDir + File.separator + "Results" + File.separator + appName + " Final Report "
					+ cal1 + ".xls";

			// String filepath = "Results\\" + selectedApp + " Final Report " +

			// cal1 + ".xls";

			File actualFile = new File(file_path);

			System.out.println("Does actual file exists :" + actualFile.exists());

			System.out.println("10 - File present");

			System.out.println("Creating WorkBook");

			Workbook wb = Workbook.getWorkbook(new File(file_path));

			Sheet sh = wb.getSheet(0);

			int totalNoOfRows = sh.getRows();

			int totalNoOfCols = sh.getColumns();

			Total_Cases = totalNoOfRows - 1;

			System.out.println("Total Rows and cols " + totalNoOfRows + " " + totalNoOfCols);

			// Failed, Passed and Critical_Error_Counter and Critical_Counter

			System.out.println("10");

			for (int i = 1; i < totalNoOfRows; i++) {

				for (int j = 0; j < totalNoOfCols; j++) {

					if (sh.getCell(j, i).getContents().equalsIgnoreCase("PASS")) {

						Passed_Case_Counter = Passed_Case_Counter + 1;

					} else if (sh.getCell(j, i).getContents().equalsIgnoreCase("FAIL")) {

						Failed_Case_Counter = Failed_Case_Counter + 1;

					} else if (sh.getCell(j, i).getContents().equalsIgnoreCase("CRITICAL")) {

						Critical_Error_Counter = Critical_Error_Counter + 1;

					}

					if (j == 3) {

						if (sh.getCell(j, i).getContents().equalsIgnoreCase("FAIL")

								&& sh.getCell(j + 1, i).getContents().equalsIgnoreCase("CRITICAL")) {

							Critical_Counter = Critical_Counter + 1;

						}

					}

				}

			}

			System.out.println("Failed, Passed and Critical_Error_Counter and Critical_Counter " + Failed_Case_Counter

					+ " " + Passed_Case_Counter + " " + Critical_Error_Counter + " " + Critical_Counter);

			System.out.println("11");

			/*
			 * 
			 * // CertifyingAs if (Critical_Counter > 0) { CertifyingAs =
			 * 
			 * "<font face='Calibri' size='12px' color='Red'><b>No-Go</b></font>"; } else if
			 * 
			 * (Failed_Case_Counter >= Passed_Case_Counter) { CertifyingAs =
			 * 
			 * "<font face='Calibri' size='12px' color='Red'><b>No-Go</b></font>"; } else {
			 * 
			 * CertifyingAs =
			 * 
			 * "<font face='Calibri' size='12px' color='Green'><b>Go</b></font>"; }
			 * 
			 */

			System.out.println("12");

			// Body Content

			/*
			 * 
			 * html = "<HTML><BODY> <font face='Calibri' size='12px'>Hello Team, " +
			 * 
			 * "<br><br>" + "Please find the " + appName + " Automated Test Report below" +
			 * 
			 * " <br> <br>" + "Certifying as " + CertifyingAs + "<br><br> URL: " + launchURL
			 * 
			 * + "<br> Date: " + cal2 + " </font><br><br>";
			 * 
			 */

			html = "<HTML><BODY> <font face='Calibri' size='12px'>Hello Team, " + "<br><br>" + "Please find the "

					+ appName + " Automated Test Report below" + "<br><br> URL: " + launchURL + "<br> Date: " + cal2

					+ " </font><br><br>";

			System.out.println("13");

			// Summary Table

			html += "<font face='Calibri' size='10px'><b>Summary</b></font><br><TABLE border='1' style='border-collapse:collapse' cellpadding '5' cellspacing='0'>";

			html += "<tr><TH bgcolor='#1A5276' align='left'>&nbsp<font face='calibri' size='3' color='White'>&nbsp<b>Total Cases Executed</b>&nbsp</font></th>";

			html += "<TH bgcolor='#1A5276' align='left'>&nbsp<font face='calibri' size='3' color='White'>&nbsp<b>Total Cases Passed</b>&nbsp</font></th>";

			html += "<TH bgcolor='#1A5276' align='left'>&nbsp<font face='calibri' size='3' color='White'>&nbsp<b>Total Cases Failed</b>&nbsp</font></th>";

			html += "<TH bgcolor='#1A5276' align='left'>&nbsp<font face='calibri' size='3' color='White'>&nbsp<b>Total Critical Errors</b>&nbsp</font></th></tr>";

			html += "<tr align='center'><Td bgcolor='Whitesmoke'>&nbsp<font face='calibri' size='2.5'>&nbsp "

					+ Total_Cases + " &nbsp</font></td>";

			html += "<Td bgcolor='Whitesmoke'>&nbsp<font face='calibri' size='2.5'>&nbsp " + Passed_Case_Counter

					+ " &nbsp</font></td>";

			html += "<Td bgcolor='Whitesmoke'>&nbsp<font face='calibri' size='2.5'>&nbsp " + Failed_Case_Counter

					+ " &nbsp</font></td>";

			html += "<Td bgcolor='Whitesmoke'>&nbsp<font face='calibri' size='2.5'>&nbsp " + Critical_Error_Counter

					+ " &nbsp</font></td>";

			html += "</tr></Table><br><br>";

			System.out.println("14");

			// Split-Up Table

			html += "<font face='Calibri' size='10px'><b>Split-up</b></font><br><TABLE border='1' style='border-collapse:collapse' cellpadding'5' cellspacing='0'>";

			html += "<TR>";

			for (int i = 0; i < totalNoOfRows; i++) {

				for (int j = 0; j < totalNoOfCols; j++) {

					if (i == 0) {

						color = "#1A5276";

						html += "<TH bgcolor=" + color

								+ " align='left'>&nbsp<font face='calibri' size='3' color='White'>&nbsp<b>"

								+ sh.getCell(j, i).getContents().toString() + "</b>&nbsp</font></th>";

					} else {

						if (sh.getCell(j, i).getContents().equalsIgnoreCase("PASS")) {

							color = "#00ed00";

						} else if (sh.getCell(j, i).getContents().equalsIgnoreCase("FAIL")) {

							color = "#ed0000";

						} else {

							color = "Whitesmoke";

						}

						if (j == 3) {

							html += "<Td bgcolor=" + color + ">&nbsp<font face='calibri' size='2.5'>&nbsp<b> "

									+ sh.getCell(j, i).getContents().toString() + "</b> &nbsp</font></td>";

						} else {

							html += "<Td bgcolor=" + color + ">&nbsp<font face='calibri' size='2.5'>&nbsp "

									+ sh.getCell(j, i).getContents().toString() + " &nbsp</font></td>";

						}

					}

				}

				html += "</TR>";

			}

			System.out.println("15");

			// Tail

			html += "</TABLE> <br><br><font face='Calibri' size='10px'><b>Note:</b><i>  Don't reply to all. This is a auto-generated report.</i></b><br> <br> Thanks & Regards <br>"

					+ User_Name + "<br></Font></BODY></HTML>";

			System.out.println("16");

			wb.close();

			out.write(html.toString());

			out.close();

		} catch (Exception e) {

			e.printStackTrace();

			System.out.println("Some problem has occured" + e.getMessage());

		}

		System.out.println("returned html");

		return html;

	}

	public static String excelMailReport() throws BiffException {

		String html;

		html = null;

		String UsrNme;

		String CertifyingAs;

		String color;

		String fontColor = "";

		int Total_Cases;

		try {

			Properties prop = new Properties();

			FileInputStream fis = new FileInputStream(
					System.getProperty("user.dir") + File.separator + "Properties" + File.separator + "OR.properties");

			prop.load(fis);

			 launchURL = prop.getProperty("url");

			//launchURL = ITestListenerImpl.envURL;

			BufferedWriter out;

			// Date

			Calendar cal = Calendar.getInstance();

			DateFormat dateFormat = new SimpleDateFormat("MM_dd_yyyy");

			String cal1 = dateFormat.format(cal.getTime());

			DateFormat dateFormat1 = new SimpleDateFormat("MM/dd/yyyy");

			String cal2 = dateFormat1.format(cal.getTime());

			String indexHtmlPath;

			if (ITestListenerImpl.defaultConfigProperty.get().getProperty("CICD").equalsIgnoreCase("N")) {

				indexHtmlPath = ITestListenerImpl.defaultConfigProperty.get().getProperty("latestReportLocation")
						+ File.separator

						+ File.separator + "Reports_" + cal1 + File.separator + "Final Report " + cal1 + ".html";

			} else {

				indexHtmlPath = ITestListenerImpl.defaultConfigProperty.get().getProperty("CICDlatestReportLocation")

						+ File.separator + File.separator + "Reports_" + cal1 + File.separator + "Final Report " + cal1

						+ ".html";

			}

			// String indexHtmlPath = System.getProperty("user.dir") + "//Results//" +
			// appName + " Final Report " + cal1

			// + ".html";

			/*
			 * 
			 * String indexHtmlPath = "C:\\Users\\" + System.getProperty("user.name") +
			 * 
			 * [file://OneDrive%20-%20St.%20James's%20Place/Desktop/NextGenBDD_Results/]\\
			 * OneDrive - St. James's Place\\Desktop\\NextGenBDD_Results\\ + appName + "
			 * 
			 * \\" + "_Report " + cal1 + File.separator+ appName + " Final Report
			 * " + cal1+ ".html";
			 * 
			 */

			out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(indexHtmlPath), "UTF-8"));

			// System.out.println("App name is : " + appName);

			// String currentDir = System.getProperty("user.dir");

			// String file_path = currentDir + "//" + "Results//" + appName + " Final Report
			// " + cal1 + ".xls";

			/*
			 * 
			 * String currentDir ="C:\\Users\\" + System.getProperty("user.name") +
			 * 
			 * [file://OneDrive%20-%20St.%20James's%20Place/Desktop/NextGenBDD_Results/]\\
			 * OneDrive - St. James's Place\\Desktop\\NextGenBDD_Results\\ + appName + "
			 * 
			 * \\" + "_Report " + cal1 + File.separator;
			 * 
			 */

			String currentDir;

			if (ITestListenerImpl.defaultConfigProperty.get().getProperty("CICD").equalsIgnoreCase("N")) {

				currentDir = ITestListenerImpl.defaultConfigProperty.get().getProperty("latestReportLocation")
						+ File.separator

						+ File.separator + "Reports_" + cal1;

			} else {

				currentDir = ITestListenerImpl.defaultConfigProperty.get().getProperty("CICDlatestReportLocation")

						+ File.separator + File.separator + "Reports_" + cal1;

			}

			String file_path = currentDir + File.separator + "Final Report " + cal1 + ".xls";

			File actualFile = new File(file_path);

			Workbook wb = Workbook.getWorkbook(new File(file_path));

			Sheet sh = wb.getSheet(0);

			int totalNoOfRows = sh.getRows();

			int totalNoOfCols = sh.getColumns();

			// Total_Cases = totalNoOfRows - 1;

			// System.out.println("Total Rows and cols " + totalNoOfRows + " " +

			// totalNoOfCols);

			// Failed, Passed and Critical_Error_Counter and Critical_Counter

			for (int i = 1; i < totalNoOfRows; i++) {

				for (int j = 0; j < totalNoOfCols; j++) {

					if (sh.getCell(j, i).getContents().equalsIgnoreCase("PASSED")) {
						String tcName = sh.getCell(1, i).getContents().toLowerCase().replaceAll(" ", "");
						if (!tcName.contains("Pre-requisite")) {
							Passed_Case_Counter = Passed_Case_Counter + 1;
						}
					} else if (sh.getCell(j, i).getContents().equalsIgnoreCase("FAILED")) {
						String tcName = sh.getCell(1, i).getContents().toLowerCase().replaceAll(" ", "");
						if (!tcName.contains("Pre-requisite")) {
							Failed_Case_Counter = Failed_Case_Counter + 1;
						}
					} else if (sh.getCell(j, i).getContents().equalsIgnoreCase("CRITICAL")) {

						Critical_Error_Counter = Critical_Error_Counter + 1;

					}

					if (j == 3) {

						if (sh.getCell(j, i).getContents().equalsIgnoreCase("FAIL")

								&& sh.getCell(j + 1, i).getContents().equalsIgnoreCase("CRITICAL")) {

							Critical_Counter = Critical_Counter + 1;

						}

					}

				}

			}

			// System.out.println("Failed, Passed Counter: " + Failed_Case_Counter + " " +

			// Passed_Case_Counter);

			// CertifyingAs

			if (Critical_Counter > 0) {

				CertifyingAs = "<font face='Calibri' size='12px' color='Red'><b>No-Go</b></font>";

			} else if (Failed_Case_Counter >= Passed_Case_Counter) {

				CertifyingAs = "<font face='Calibri' size='12px' color='Red'><b>No-Go</b></font>";

			} else {

				CertifyingAs = "<font face='Calibri' size='12px' color='Green'><b>Go</b></font>";

			}
			Total_Cases = Passed_Case_Counter + Failed_Case_Counter;

			// Body Content
			// 1) Body intro
			html = "<HTML><BODY>"
				     + "<font face='Calibri' size='5'><b>Hello Team,</b></font><br>"

				// Exercise disclaimer
				     + "<font face='Calibri' size='3' color='#555;'>"
				     + "🤖 Rest assured—this email is simply the result of my NHSBSA - Automation Test Analyst role technical exercise, not any phishing attempt!😜"
				     + "</font><br><br>"

				// … then the rest of your content …
				     + "<font face='Calibri' size='4'>Please find the "
				     + appName + " Automated Test Report below.</font><br><br>"

			// if
			// (ITestListenerImpl.defaultConfigProperty.get().getProperty("CICD").equalsIgnoreCase("N"))
			// {
			//
			// html += "<br>Please find the Detailed Report under <b><a href="
			//
			// + ITestListenerImpl.defaultConfigProperty.get().getProperty("MailReportPath")
			// + ">"
			//
			// + File_Transfer.backUpReportName + "</a></b>" + ".</font><br><br>";
			//
			// // html += "<br><font face='Calibri' size='12px'><i>Detailed Automation
			// report
			//
			// // is available at <b><a
			//
			// //
			// href="+ITestListenerImpl.defaultConfigProperty.get().getProperty("DetailedReportPath")+">Detailed
			//
			// // Automation Report</a></b></i></font><br><br>";
			//
			// }

 // 2) URL / Date on one block, no extra blank lines
 +"<div style=\"font-size:16px; margin-bottom:20px;\">" +
   "<strong>URL:</strong> <a href=\"" + launchURL + "\" style=\"color:#1A5276; text-decoration:none;\">" + launchURL + "</a><br/>" +
   "<strong>Date:</strong> " + cal2 +
 "</div>" ;

			// Summary Table

			/*
			 * html +=
			 * "<font face='Calibri' size='4.5'><b>SUMMARY:</b></font><br><TABLE border='1' style='border-collapse:collapse' cellpadding '5' cellspacing='0'>"
			 * ;
			 * 
			 * html +=
			 * "<tr height='30'><TH bgcolor='#1A5276' align='left'>&nbsp<font face='calibri' size='3' color='White'>&nbsp<b>Total Scenarios Executed</b>&nbsp</font></th>"
			 * ;
			 * 
			 * html +=
			 * "<TH bgcolor='#1A5276' align='left'>&nbsp<font face='calibri' size='3' color='White'>&nbsp<b>Total Scenarios Passed</b>&nbsp</font></th>"
			 * ;
			 * 
			 * html +=
			 * "<TH bgcolor='#1A5276' align='left'>&nbsp<font face='calibri' size='3' color='White'>&nbsp<b>Total Scenarios Failed</b>&nbsp</font></th>"
			 * ;
			 * 
			 * html += "</tr>";
			 * 
			 * html +=
			 * "<tr align='center'><Td bgcolor='Whitesmoke'>&nbsp<font face='calibri' size='2.5'>&nbsp "
			 * 
			 * + Total_Cases + " &nbsp</font></td>";
			 * 
			 * html +=
			 * "<Td bgcolor='Whitesmoke'>&nbsp<font face='calibri' size='2.5'>&nbsp " +
			 * Passed_Case_Counter
			 * 
			 * + " &nbsp</font></td>";
			 * 
			 * html +=
			 * "<Td bgcolor='Whitesmoke'>&nbsp<font face='calibri' size='2.5'>&nbsp " +
			 * Failed_Case_Counter
			 * 
			 * + " &nbsp</font></td>";
			 * 
			 * html += "</tr></Table><br><br>";
			 * 
			 * // Split-Up Table
			 * 
			 * html +=
			 * "<font face='Calibri' size='4.5'><b>SPLIT-UP:</b></font><br><TABLE border='1' style='border-collapse:collapse' cellpadding'5' cellspacing='0'>"
			 * ;
			 * 
			 * html += "<TR height='30'>";
			 * 
			 * for (int i = 0; i < totalNoOfRows; i++) {
			 * 
			 * for (int j = 0; j < totalNoOfCols; j++) {
			 * 
			 * if (i == 0) {
			 * 
			 * if (j == 4 || j == 5) {
			 * 
			 * color = "#1A5276";
			 * 
			 * html += "<TH width='55' bgcolor=" + color
			 * 
			 * + " align='center'>&nbsp<font face='calibri' size='3' color='White'><b>"
			 * 
			 * + sh.getCell(j, i).getContents().toString() + "&nbsp</b></font></th>";
			 * 
			 * } else if (j == 3) {
			 * 
			 * color = "#1A5276";
			 * 
			 * html += "<TH width='45' bgcolor=" + color
			 * 
			 * + " align='center'>&nbsp<font face='calibri' size='3' color='White'><b>"
			 * 
			 * + sh.getCell(j, i).getContents().toString() + "&nbsp</b></font></th>";
			 * 
			 * } else {
			 * 
			 * color = "#1A5276";
			 * 
			 * html += "<TH bgcolor=" + color
			 * 
			 * + " align='center'>&nbsp<font face='calibri' size='3' color='White'><b>"
			 * 
			 * + sh.getCell(j, i).getContents().toString() + "&nbsp</b></font></th>";
			 * 
			 * }
			 * 
			 * } else {
			 * 
			 * if (sh.getCell(j, i).getContents().equalsIgnoreCase("PASS")) {
			 * 
			 * color = "#00ed00";
			 * 
			 * } else if (sh.getCell(j, i).getContents().equalsIgnoreCase("PASSED")) {
			 * 
			 * color = "#0d6b2e";
			 * 
			 * fontColor = "#FFFFFF";
			 * 
			 * } else if (sh.getCell(j, i).getContents().equalsIgnoreCase("FAIL")) {
			 * 
			 * color = "#ed0000";
			 * 
			 * } else if (sh.getCell(j, i).getContents().equalsIgnoreCase("FAILED")) {
			 * 
			 * color = "#930000";
			 * 
			 * fontColor = "#FFFFFF";
			 * 
			 * } else {
			 * 
			 * color = "Whitesmoke";
			 * 
			 * }
			 * 
			 * if (j == 2) {
			 * 
			 * html += "<Td align='center' bgcolor=" + color +
			 * ">&nbsp<font face='calibri' size='4' color="
			 * 
			 * + fontColor + ">&nbsp<b> " + sh.getCell(j, i).getContents().toString()
			 * 
			 * + "</b> &nbsp</font></td>";
			 * 
			 * } else if (j == 0 || j == 3 || j == 4 || j == 5) {
			 * 
			 * html += "<Td align='center' bgcolor=" + color +
			 * ">&nbsp<font face='calibri' size='2'>&nbsp "
			 * 
			 * + sh.getCell(j, i).getContents().toString() + " &nbsp</font></td>";
			 * 
			 * } else {
			 * 
			 * html += "<Td align='left' bgcolor=" + color +
			 * ">&nbsp<font face='calibri' size='2'>&nbsp "
			 * 
			 * + sh.getCell(j, i).getContents().toString() + " &nbsp</font></td>";
			 * 
			 * }
			 * 
			 * }
			 * 
			 * }
			 * 
			 * html += "</TR>";
			 * 
			 * }
			 */
			// … before SUMMARY …
			html += "<font face='Calibri' size='4.5'><b>SUMMARY:</b></font><br>"
			     + "<table"
			     +    " cellpadding='5' cellspacing='0'"
			     +    " style='"
			     +      "width:80%;"
			     +      "table-layout:fixed;"
			     +      "border-collapse:collapse;"
			     +      "font-family:Calibri;"
			     +    "'>"
			     +   "<tr style='background-color:#1A5276;color:#ffffff;text-align:center;'>"
			     +     "<th style='width:33%;border:1px solid #dddddd;'>Total Scenarios Executed</th>"
			     +     "<th style='width:33%;border:1px solid #dddddd;'>Total Scenarios Passed</th>"
			     +     "<th style='width:34%;border:1px solid #dddddd;'>Total Scenarios Failed</th>"
			     +   "</tr>"
			     +   "<tr style='background-color:#ffffff;color:#000000;text-align:center;'>"
			     +     "<td style='border:1px solid #dddddd;'>" + Total_Cases + "</td>"
			     +     "<td style='border:1px solid #dddddd;'>" + Passed_Case_Counter + "</td>"
			     +     "<td style='border:1px solid #dddddd;'>" + Failed_Case_Counter + "</td>"
			     +   "</tr>"
			     + "</table><br><br>";


			/* ----------------------------------------------------------------------
			   SPLIT-UP TABLE
			------------------------------------------------------------------------*/
			// just before your loop:
			html += "<font face='Calibri' size='4.5'><b>SPLIT-UP:</b></font><br>"
				     + "<table"
				     +    " cellpadding='5' cellspacing='0'"
				     +    " style='"
				     +      "width:80%;"
				     +      "table-layout:fixed;"
				     +      "border-collapse:collapse;"
				     +      "font-family:Calibri;"
				     +    "'>"
				     +   "<tr style='background-color:#1A5276;color:#ffffff;text-align:center;'>"
				     +     "<th style='width:10%;border:1px solid #dddddd;'>S.No</th>"
				     +     "<th style='width:30%;border:1px solid #dddddd;'>Test Case</th>"
				     +     "<th style='width:20%;border:1px solid #dddddd;'>Overall Status</th>"
				     +     "<th style='width:10%;border:1px solid #dddddd;'>Chrome</th>"
				     +     "<th style='width:10%;border:1px solid #dddddd;'>Firefox</th>"
				     +     "<th style='width:20%;border:1px solid #dddddd;'>Comments</th>"
				     +   "</tr>";
			
			for (int i = 1; i < totalNoOfRows; i++) {
			    html += "<tr style='background-color:#ffffff;color:#000000;text-align:center;'>";
			    // S.No
			    html += "<td style='border:1px solid #dddddd;'>" + sh.getCell(0, i).getContents() + "</td>";
			    // Test Case
			    html += "<td style='border:1px solid #dddddd;'>" + sh.getCell(1, i).getContents() + "</td>";
			    // Overall Status
			    String status = sh.getCell(2, i).getContents();
			    String bgColor, fontColor1;

			 // green for PASSED / PASS
			 if (status.equalsIgnoreCase("PASSED")) {
			     bgColor   = "#0d6b2e";  // dark green
			     fontColor1 = "#FFFFFF";
			 } else if (status.equalsIgnoreCase("PASS")) {
			     bgColor   = "#00ed00";  // bright green
			     fontColor1 = "#000000";
			 }

			 // **swapped reds**:
			 // bright red for FAILED, dark red for FAIL
			 else if (status.equalsIgnoreCase("FAILED")) {
			     bgColor   = "#ed0000";  // bright red
			     fontColor1 = "#FFFFFF";
			 } else if (status.equalsIgnoreCase("FAIL")) {
			     bgColor   = "#930000";  // dark red
			     fontColor1 = "#FFFFFF";
			 }

			 // fallback
			 else {
			     bgColor   = "Whitesmoke";
			     fontColor1 = "#000000";
			 }

			 // now render the cell:
			 html += "<td style='border:1px solid #dddddd;"
			       + "background-color:" + bgColor + ";"
			       + "color:" + fontColor1 + ";"
			       + "text-align:center;'>"
			       + (status.startsWith("PAS") ? "<b>" + status + "</b>" : status)
			       + "</td>";
			    // Chrome
			    html += "<td style='border:1px solid #dddddd;'>" + sh.getCell(3, i).getContents() + "</td>";
			    // Firefox
			    html += "<td style='border:1px solid #dddddd;'>" + sh.getCell(4, i).getContents() + "</td>";
			    // Comments
			    html += "<td style='border:1px solid #dddddd;'>" + sh.getCell(5, i).getContents() + "</td>";
			    html += "</tr>";
			}

			html += "</table><br><br>";
			// Tail

			html += "</TABLE> <br><font face='Calibri' size='7px'>";

			// html += "<br><br><b>Note:</b>";

			// html += "<i> Please find attached Detailed Automation Report.</i>";

/*			html += "</TABLE>" +
			        "<div style=\"font-family:Calibri; font-size:14px; margin:20px 0 0 0;\">" +
			          "Thanks & Regards,<br/>" +
			          User_Name +
			        "</div>" +
			      "</BODY></HTML>";

			html += "<br><font face='Calibri' size='5px'><i>Don't reply to all as this is a auto-generated report. Any further queries please contact <b><a href='mailto:"

					+ teamEmailID + "'>" + User_Name + "</a></b></i></font>";

			html += "<br></Font></BODY></HTML>";*/
			
			html +=  // close your tables here…
				       "</table>" +

				       // Sign-off
				       "<div style=\"font-size:16px; margin-top:30px;\">" +
				         "Thanks &amp; Regards,<br/>" +
				         User_Name +
				       "</div>" +

				       // Uniform smaller footer
				       "<div style=\"font-size:16px; font-style:italic; color:#555; margin-top:16px;\">" +
				         "Don't reply to all as this is an auto-generated report. Any further queries please contact " +
				         "<strong>NHSBSA Test Automation</strong>" +
				       "</div>" +

				       "</body></html>";

			wb.close();

			out.write(html.toString());

			out.close();

			System.out.println("==========>Created HTML Report from Excel Report<==========");

		} catch (Exception e) {

			// e.printStackTrace();

			System.out.println("Some problem has occured while sending mail" + e.getMessage());

		}

		return html;

	}

} // reportgenerator