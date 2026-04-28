package com.majesco.itaf.report.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.commons.lang.math.NumberUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.ITAFWebDriver;
import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;

public class Functions {
	static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	private final static Logger log = LogManager.getLogger(Functions.class.getName());
	static String[] ignoreTransArray = { "SC_END", "Freeze" };
	static List<String> ignoreTransList = Arrays.asList(ignoreTransArray);

	/* new change */
	static String[] ignoreTestCaseArray = { "Common", "TC_Freeze", "RunBatch", "RunBatch_JBEAMFF", "RunBatch_JBEAM_DAY",
			"RunBatch_JBEAM_NIGHT", "RunBatch_JBEAM", "CopyFlatFile" };
	static List<String> ignoreTestCaseList = Arrays.asList(ignoreTestCaseArray);

	/* Converts csv file to .xlsx file */
	public static String convertCsvToXlsx(XSSFWorkbook workBook, String xlsFileLocation, String csvFilePath) {
		String FILE_EXTN = ".xlsx";
		String FILE_NAME = "VisualReport";
		XSSFSheet sheet = null;
		CSVReader reader = null;
		String generatedXlsFilePath = "";
		@SuppressWarnings("unused")
		int startIndex = csvFilePath.lastIndexOf("\\");
		@SuppressWarnings("unused")
		int lastIndex = csvFilePath.lastIndexOf(".csv");
		// FILE_NAME= csvFilePath.substring(startIndex+1, lastIndex);

		try {
			/**** Get the CSVReader Instance & Specify The Delimiter To Be Used ****/
			String[] nextLine;
			CSVParser csvParser = new CSVParserBuilder().withSeparator(',').build();

			reader = new CSVReaderBuilder(new FileReader(csvFilePath)).withCSVParser(csvParser).build();
			// reader = new CSVReader(new FileReader(csvFilePath), ',');

			sheet = (XSSFSheet) workBook.createSheet("Summary");

			int rowNum = 0;
			while ((nextLine = reader.readNext()) != null) {
				Row currentRow = sheet.createRow(rowNum++);
				for (int i = 0; i < nextLine.length; i++) {
					if (NumberUtils.isDigits(nextLine[i])) {
						currentRow.createCell(i).setCellValue(Integer.parseInt(nextLine[i]));
					} else if (NumberUtils.isNumber(nextLine[i])) {
						currentRow.createCell(i).setCellValue(Double.parseDouble(nextLine[i]));
					} else {
						currentRow.createCell(i).setCellValue(nextLine[i]);
					}
				}
			}
			generatedXlsFilePath = xlsFileLocation + "\\" + FILE_NAME + FILE_EXTN;
		} catch (Exception exObj) {
			System.out.println("Error occured while converting file" + exObj.getMessage());
		}

		finally {
			try {
				reader.close();
			} catch (IOException ioExObj) {
			}
		}

		return generatedXlsFilePath;
	}

	/* Import image in passed XSSFsheet in XSSFworkbook */
	public static void importImage(String imgFolderLocation, XSSFWorkbook workbook, String sheet, String filename,
			int rownum1, int colnum1, int rownum2, int colnum2) throws IOException {
		InputStream inputStream = new FileInputStream(imgFolderLocation + "\\" + filename);
		byte[] bytes = IOUtils.toByteArray(inputStream);
		int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
		inputStream.close();
		XSSFSheet VisualSheet = workbook.getSheet(sheet);
		CreationHelper helper = workbook.getCreationHelper();
		@SuppressWarnings("rawtypes")
		Drawing drawing = VisualSheet.createDrawingPatriarch();
		ClientAnchor anchor = helper.createClientAnchor();
		anchor.setCol1(colnum1);
		anchor.setRow1(rownum1);
		anchor.setCol2(colnum2);
		anchor.setRow2(rownum2);
		Picture pict = drawing.createPicture(anchor, pictureIdx);
		pict.resize();
	}

	/* converts time to milliseconds in long format from string format */
	public static long converttomillis(String time) {
		String[] arr = time.split(":"); // splits string in hh:mm:ss or hh:mm:ss:millis format
		long milliseconds = 0;
		if (arr.length == 4) // splits hh:mm:ss:millis
		{
			milliseconds = (Integer.parseInt(arr[0]) * 60 * 60 + Integer.parseInt(arr[1]) * 60
					+ Integer.parseInt(arr[2])) * 1000 + Integer.parseInt(arr[3]);

		} else if (arr.length == 3) // splits hh:mm:ss
		{
			milliseconds = (Integer.parseInt(arr[0]) * 60 * 60 + Integer.parseInt(arr[1]) * 60
					+ Integer.parseInt(arr[2])) * 1000;

		} else // for time in seconds
		{
			milliseconds = (long) (Double.parseDouble(time) * 1000);
		}
		return milliseconds;
	}

	/* converts seconds to hh:mm:ss string */
	public static String convertSecondsToHMmSs(long seconds) {
		long s = seconds % 60;
		long m = (seconds / 60) % 60;
		long h = (seconds / (60 * 60)) % 24;
		return String.format("%02d:%02d:%02d", h, m, s);
	}

	/* calculates total time for all transactions and returns hh:m:ss string */
	public static String totalTime() {
		long etime = 0;
		SummaryObject obj = null;
		for (Map.Entry<String, SummaryObject> m : MapGenerator.SummaryMap.entrySet()) {
			obj = (SummaryObject) m.getValue();
			etime = etime + obj.getEtime();
		}
		return Functions.convertSecondsToHMmSs(etime / 1000);
	}

	/* get total count for all scenarios from maincontroller */
	public static HashMap<String, Integer> getTotalCount() throws IOException {
		String mainContPath = Config.controllerFilePath;
		Sheet mainContSheet = getSheet(mainContPath, "MainControlSheet");

		HashMap<String, Integer> sheetHeaderMap = getValueFromHashMap(mainContSheet);
		int testCaseIdColNo = 0;
		if (Config.executionApproach.equalsIgnoreCase("Linear") && VisualReport.sortByGroup) {
			testCaseIdColNo = sheetHeaderMap.get("GroupName");
		} else {
			testCaseIdColNo = sheetHeaderMap.get("TestCaseID");
		}

		HashMap<String, Integer> totalTranCountMap = new HashMap<String, Integer>();

		boolean advMainCont = false;
		// as only billing & claims teams use pointer concept in maincontroller//Sel4
		if (ITAFWebDriver.isBillingApplication() || ITAFWebDriver.isClaimsApplication())
			advMainCont = true;

		Row row = null;
		String tempScName = null;
		for (String scenarioName : MapGenerator.SummaryMap.keySet()) {
			int tranCount = 0;
			for (int i = 1; i < mainContSheet.getPhysicalNumberOfRows(); i++) {
				row = mainContSheet.getRow(i);
				tempScName = row.getCell(testCaseIdColNo).getStringCellValue();
				if ((tempScName.equals(scenarioName) && Config.executionApproach.equalsIgnoreCase("Linear"))
						| ((tempScName.startsWith(scenarioName + "_") || tempScName.equals(scenarioName))
								&& Config.executionApproach.equalsIgnoreCase("TimeTravel")))
					tranCount = tranCount + calTransInRow(row, advMainCont);
			}
			totalTranCountMap.put(scenarioName, tranCount);
		}
		return totalTranCountMap;
	}

	/* get particular sheet from workbook at passed filepath */
	public static Sheet getSheet(String FilePath, String SheetName) throws IOException {
		Workbook workBook = null;

		Sheet workSheet = null;
		try {
			InputStream myXls = new FileInputStream(FilePath);
			if (FilePath.endsWith(".xls")) {
				workBook = new HSSFWorkbook(myXls);
			} else if (FilePath.endsWith(".xlsx") | FilePath.endsWith(".xlsm")) {
				workBook = new XSSFWorkbook(myXls);
			}
			workSheet = workBook.getSheet(SheetName);
		} catch (IOException e) {
			System.out.println("Exception occured while accessing file: " + e.getMessage());
		} finally {
			workBook.close();
		}
		return workSheet;
	}

	/* returns hashmap with all columnnumbers */
	public static HashMap<String, Integer> getValueFromHashMap(Sheet reqSheet) {
		HashMap<String, Integer> inputHashTable = new HashMap<String, Integer>();
		Row rowHeader = reqSheet.getRow(0);
		int columnCount = rowHeader.getPhysicalNumberOfCells();
		for (int colIndex = 0; colIndex < columnCount; colIndex++) {
			inputHashTable.put(rowHeader.getCell(colIndex).toString(), colIndex);
		}
		return inputHashTable;
	}

	/* calculates transactions of scenario in given row from main controller */
	public static int calTransInRow(Row row, boolean advMainCont) {
		int startColNo;
		int transCount = 0;
		if (advMainCont)
			startColNo = 10;
		else
			startColNo = 6;

		for (int i = startColNo; i < row.getLastCellNum(); i++) {
			Cell c = row.getCell(i);
			if (c != null && !c.getStringCellValue().equals("") && !ignoreTransList.contains(c.getStringCellValue()))
				transCount++;
		}
		return transCount;
	}

	/*
	 * set background as white for mentioned rows and columns of sheet in workbook
	 */
	public static void setBackgroundWhite(XSSFWorkbook workbook, String sheetname, int startRowIndex, int startColIndex,
			int lastRowIndex, int lastColIndex) {
		XSSFSheet sheet = workbook.getSheet(sheetname);
		CellStyle style = StyleGenerator.getCellStyleWhite(workbook);
		for (int rownum = startRowIndex; rownum <= lastRowIndex; rownum++) {
			Row row = sheet.createRow(rownum);

			for (int colnum = startColIndex; colnum <= lastColIndex; colnum++) {
				Cell cell = row.createCell(colnum);
				cell.setCellStyle(style);
			}
		}
	}

	public static void clearDirectory(String directoryPath) {
		File dir = new File(directoryPath);

		if (dir.isDirectory() == false) {
			log.info("No existing directory at path: " + directoryPath + "  so new directory will be created");
			return;
		} else {
			File[] listFiles = dir.listFiles();
			for (File file : listFiles) {
				file.delete();
			}
			// log.info("Cleared all existing charts");
		}

	}
}
