package com.majesco.itaf.util;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.MainController;
import com.majesco.itaf.main.ObjectFactory;
import com.majesco.itaf.vo.Reporter;

public class DB_ValidationBilling {

	public static MainController controller = ObjectFactory.getMainController();

	public static Connection conn = null;
	public static Statement stmt = null;
	public static ResultSet rs = null;
	public static String dateForFileName = (controller.cycleDateCellValue).replaceAll("/", "_");
	public static String filename2 = Config.resultFilePath.toString() + "DBVerification//DBData_Result" + ".csv";
	public static File f2 = new File(filename2);
	public static char myChar = 34;
	public static PrintStream print = null;
	protected static List<String> status = new ArrayList<String>();
	protected static List<String> rowStatus = new ArrayList<String>();
	protected static List<String> actualValue = new ArrayList<String>();
	public static List<Integer> PassCount = new ArrayList<Integer>();
	public static List<Integer> FailCount = new ArrayList<Integer>();
	public static int firstRow = 1;
	public static int TotalpassCount;
	public static int TotalfailCount;
	protected static List<List<String>> actualRows = new ArrayList<List<String>>();
	public static int DResult = 1;
	public static Date frmDate = new Date();
	public static DateFormat dtFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
	public static String interfaceServerLocation;

	public static void DBValidationWrtitetoExcel(Map<List<String>, List<List<String>>> Databasemap) throws IOException // Added
	{
		String CurrentDate = controller.cycleDateCellValue;
		String dateForFile = CurrentDate.replaceAll("/", "_");
		String filename = Config.inputDataFilePath + "DBVerification//DatabaseActualOutput" + "\\"
				+ controller.controllerTestCaseID.toString() + "_" + dateForFile + "_Response.xlsx";

		File f = new File(filename);
		String filename1 = Config.inputDataFilePath + "DBVerification//DatabaseExpectedOutput" + "\\"
				+ controller.controllerTestCaseID.toString() + "_" + dateForFile + "_Expected" + ".xlsx";

		File f1 = new File(filename1);
		XSSFWorkbook wb = new XSSFWorkbook();;
		XSSFSheet sheet = wb.createSheet("Actual");
		for (Map.Entry<List<String>, List<List<String>>> entry : Databasemap.entrySet()) {
			List<String> Columnkey = entry.getKey();
			List<List<String>> values = entry.getValue();

			int rownum = 0;
			int cellnum = 0;
			Row row = sheet.createRow(rownum);
			for (String ColValue : Columnkey) {
				Cell cell = row.createCell(cellnum++);
				cell.setCellValue(ColValue);
			}

			int size = values.size();
			for (List<String> SubList : values) {
				for (int rc = 1; rc <= size; rc++) {
					Row row1 = sheet.createRow(rc);
					cellnum = 0;
					Cell cell1;
					for (String Rowvalue : SubList) {
						if (cellnum < Columnkey.size()) {
							cell1 = row1.createCell(cellnum++);
							cell1.setCellValue(Rowvalue);
						} else {
							row1 = sheet.createRow(++rc);
							cellnum = 0;
							cell1 = row1.createCell(cellnum++);
							cell1.setCellValue(Rowvalue);
						}
					}
				}
			}
		}
		if (f.exists() && !f.isDirectory()) {
			FileOutputStream fileOut = new FileOutputStream(f);
			wb.write(fileOut);
			fileOut.close();
			System.out.println("Your excel file_1 has been generated!");
		} else {
			FileOutputStream fileOut = new FileOutputStream(f);
			wb.write(fileOut);
			fileOut.close();
			System.out.println("Your excel file_0 has been generated!");
		}
		wb.close();

		if (f.exists() && f1.exists()) {

			@SuppressWarnings("unused")
			DataFormatter df = new DataFormatter();
			try {
				FileInputStream file1 = new FileInputStream(f1);
				FileInputStream file2 = new FileInputStream(f);
				Workbook wb1 = WorkbookFactory.create(file1);// expected file

				Sheet sh1 = wb1.getSheet("Actual");
				if (sh1 != null) {
					Workbook wb2 = WorkbookFactory.create(file2);// actual file
					Sheet sh2 = wb2.getSheet("Actual");
					@SuppressWarnings({ "unused" })
					XSSFSheet expectedSheet = (XSSFSheet) sh1;
					@SuppressWarnings("unused")
					XSSFSheet actualSheet = (XSSFSheet) sh2;
					@SuppressWarnings("unused")
					Reporter report = new Reporter();
				} else {
					Exception exception = new Exception();
					System.out.println("Expected Sheet is not present " + exception);
				}

			}

			catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

	}

	public static Reporter CompareExcel(XSSFSheet actualSheet, XSSFSheet expectedSheet, List<String> columns,
			List<String> columnsData) throws IOException {

		@SuppressWarnings("unused")
		boolean isrowFound = false;
		int expSheetRowCount = expectedSheet.getPhysicalNumberOfRows(); // getRowCount(expectedSheet);
		Reporter report = new Reporter();
		report.setReport(report);
		int passCount = 0;
		int failCount = 0;
		int colCount = 0;

		for (int rowIndex = firstRow; rowIndex < expSheetRowCount; rowIndex++) {
			passCount = 0;
			failCount = 0;

			XSSFRow actualRow = actualSheet.getRow(rowIndex);
			XSSFRow expectedRow = expectedSheet.getRow(rowIndex);
			System.out.println(actualRow.getCell(0).toString());
			System.out.println(expectedRow.getCell(0).toString());

			actualValue = new ArrayList<String>();
			if (actualRow == null || expectedRow == null) {
				break;
			}

			colCount = expectedRow.getLastCellNum();
			for (int columnIndex = 0; columnIndex <= colCount; columnIndex++) {
				XSSFCell actualCell = actualRow.getCell(columnIndex);
				DataFormatter fmt = new DataFormatter();
				XSSFCell expectedCell = expectedRow.getCell(columnIndex);

				if (actualCell != null || expectedCell != null) {
					String expectedValue = fmt.formatCellValue(expectedCell);

					if (!actualCell.toString().equalsIgnoreCase(expectedValue)) {
						if (actualCell.getColumnIndex() == 0 || actualCell.getColumnIndex() == 3
								|| actualCell.getColumnIndex() == 39) {
							passCount += 1;
							TotalpassCount += 1;
							report.setStatus("PASS");
							// report.setStrStatus(report.strStatus);
							report.setActualValue(actualCell.toString());
							System.out.println(actualCell.toString());
						} else {
							report.setStatus("FAIL");
							// report.setStrStatus(report.strStatus);
							failCount += 1;
							TotalfailCount += 1;
							System.out.println(TotalfailCount);
							report.setActualValue(DB_ValidationBilling.myChar + "," + DB_ValidationBilling.myChar
									+ "Expected : " + expectedValue + DB_ValidationBilling.myChar // changes by bhushan
									// for in one cell
									// expected and
									// actual data
									+ DB_ValidationBilling.myChar + "Actual : " + actualCell.toString()
									+ DB_ValidationBilling.myChar + "," + DB_ValidationBilling.myChar);
						}
					} else {
						passCount += 1;
						TotalpassCount += 1;
						report.setStatus("PASS");
						// report.setStrStatus(report.strStatus);
						report.setActualValue(actualCell.toString());
						System.out.println(actualCell.toString());

					}
					status.add(report.getStatus());
					actualValue.add(report.getActualValue());
				}

			}
			if (status.contains("FAIL")) {
				report.setStatus("FAIL");
			} else {
				report.setStatus("PASS");
			}

			status.clear();
			System.out.println("---" + rowIndex + "----");
			rowStatus.add(report.getStatus());
			PassCount.add(passCount);
			FailCount.add(failCount);
			actualRows.add(actualValue);
			report.setReport(report);
		}

		if (rowStatus.contains("FAIL")) {
			report.setStatus("FAIL");
		}
		WriteToDetailResults(columns, actualRows, passCount, failCount, expSheetRowCount, colCount, report, rowStatus);
		PassCount.clear();
		FailCount.clear();
		return report;
	}

	public static void WriteToDetailResults(List<String> columns, List<List<String>> columnsData, int passCount,
			int failCount, int rowCount, int colCount, Reporter report, List<String> status) throws IOException {
		try {

			report = report.getReport();
			report.getFromDate();
			report.setTestcaseId(controller.controllerTestCaseID.toString());
			report.setTrasactionType(controller.controllerTransactionType.toString());
			// report.strCycleDate = controller.cycleDateCellValue;
			report.getCycleDate();
			report.getStatus();
			if (DResult == 1) { // Bhushan: for deleting old result file
				f2.delete();
				DResult++;
			}
			if (f2.exists() == false) {
				System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$");
				print = new PrintStream(f2);
			}

			print = new PrintStream(new FileOutputStream(f2, true));
			int usedRows = count(f2);
			if (usedRows == 0) {

				print.println(
						"Iteration,TestCaseID,TransactionType,CycleDate,CurrentDate,RowType,Status,PassCount,FailCount,Pass/FailData");
			}
			usedRows = count(f2);

			print.print(DB_ValidationBilling.myChar + "" + DB_ValidationBilling.myChar + ","
					+ DB_ValidationBilling.myChar + report.getTestcaseId() + DB_ValidationBilling.myChar + ","
					+ DB_ValidationBilling.myChar + report.getTrasactionType() + DB_ValidationBilling.myChar + ","
					+ DB_ValidationBilling.myChar + report.getCycleDate() + DB_ValidationBilling.myChar + ","
					+ DB_ValidationBilling.myChar + report.getFromDate() + DB_ValidationBilling.myChar + ","
					+ DB_ValidationBilling.myChar + "Header" + DB_ValidationBilling.myChar + ","
					+ DB_ValidationBilling.myChar + report.getStatus() + DB_ValidationBilling.myChar + ","
					+ DB_ValidationBilling.myChar + "" + DB_ValidationBilling.myChar + "," + DB_ValidationBilling.myChar
					+ "" + DB_ValidationBilling.myChar);
			String r = report.getStatus();
			System.out.println(r);
			int counter = 0;
			while (columns.isEmpty() == false) {
				if (counter != columns.size()) {
					print.print("," + DB_ValidationBilling.myChar + columns.get(counter) + DB_ValidationBilling.myChar);
					counter++;
				} else {
					break;
				}
			}
			print.println();
			rowCount = actualRows.size();
			for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
				print.print(DB_ValidationBilling.myChar + "" + DB_ValidationBilling.myChar + ","
						+ DB_ValidationBilling.myChar + report.getTestcaseId() + DB_ValidationBilling.myChar + ","
						+ DB_ValidationBilling.myChar + report.getTrasactionType() + DB_ValidationBilling.myChar + ","
						+ DB_ValidationBilling.myChar + report.getCycleDate() + DB_ValidationBilling.myChar + ","
						+ DB_ValidationBilling.myChar + report.getFromDate() + DB_ValidationBilling.myChar + ","
						+ DB_ValidationBilling.myChar + "Data" + DB_ValidationBilling.myChar + ","
						+ DB_ValidationBilling.myChar + rowStatus.get(rowIndex).toString() + DB_ValidationBilling.myChar
						+ "," + DB_ValidationBilling.myChar + PassCount.get(rowIndex) + DB_ValidationBilling.myChar
						+ "," + DB_ValidationBilling.myChar + FailCount.get(rowIndex) + DB_ValidationBilling.myChar
						+ "," + DB_ValidationBilling.myChar + "" + DB_ValidationBilling.myChar);

				counter = 0;
				while (actualRows.isEmpty() == false) {
					if (counter != actualRows.get(rowIndex).size()) {
						System.out.print(actualRows.get(rowIndex).get(counter));
						print.print(DB_ValidationBilling.myChar + actualRows.get(rowIndex).get(counter)
								+ DB_ValidationBilling.myChar);
						counter++;
					} else {
						break;
					}
				}
				print.println();

			}
		} catch (Exception e) {
			System.out.println(e.getMessage());

		} finally {
			actualRows.clear();
			rowStatus.clear();
			columns.clear();
			columnsData.clear();
		}

	}

	public static int count(File filename) throws IOException {
		InputStream is = new BufferedInputStream(new FileInputStream(filename));
		try {
			byte[] c = new byte[1024];
			int count = 0;
			int readChars = 0;
			boolean empty = true;
			while ((readChars = is.read(c)) != -1) {
				empty = false;
				for (int i = 0; i < readChars; ++i) {
					if (c[i] == '\n')
						++count;
				}
			}
			return (count == 0 && !empty) ? 1 : count;
		} finally {
			is.close();
		}
	}
}
