package com.majesco.itaf.rest.utils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TreeSet;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.github.wnameless.json.flattener.JsonFlattener;

public class ValidateJson {
	private final static Logger log = LogManager.getLogger(ValidateJson.class);
	static Map<String, Object> flattenExpectedJsonMap = null;
	static Map<String, Object> flattenResponseJsonMap = null;
	static Map<String, Object> valueErrorJsonMap = null;
	static HashSet<String> extraExpectedKeys = null;
	static HashSet<String> extraResponseKeys = null;

	public static boolean main(String expectedJsonPath, String responseJsonPath, List<String> ignoreKeyList)
			throws Exception {
		boolean compareJsonFlag = false;
		// To check whether response json matches with expected json or not.
		compareJsonFlag = compareJson(expectedJsonPath, responseJsonPath, ignoreKeyList);
		if (compareJsonFlag)
			compareSheetCreator(flattenExpectedJsonMap, flattenResponseJsonMap, valueErrorJsonMap, extraExpectedKeys,
					extraResponseKeys, expectedJsonPath, responseJsonPath);
		return !compareJsonFlag;
	}

	// Below given Json code has been chnaged to incorporate BaseBilling Date
	// changes - any key that will contain Date text will now ignore the offset
	public static boolean compareJson(String expectedJsonPath, String responseJsonPath, List<String> ignoreKeyList)
			throws IOException {
		boolean compareJsonFlag;
		String expectedJsonString = new String(Files.readAllBytes(Paths.get(expectedJsonPath)));
		String responseJsonString = new String(Files.readAllBytes(Paths.get(responseJsonPath)));

		flattenExpectedJsonMap = JsonFlattener.flattenAsMap(expectedJsonString);
		valueErrorJsonMap = JsonFlattener.flattenAsMap(expectedJsonString);
		flattenResponseJsonMap = JsonFlattener.flattenAsMap(responseJsonString);

		Iterator<Entry<String, Object>> iterator = flattenExpectedJsonMap.entrySet().iterator();
		while (iterator.hasNext()) {
			// Remove key-value pairs present in ignore list
			Entry<String, Object> entry = iterator.next();
			String key = entry.getKey();
			String expectedValue = String.valueOf(entry.getValue());

			// Extract the date from the expected value if the key contains "date"
			if (key.contains("Date")) {
				String[] parts = expectedValue.split("\\+");
				if (parts.length > 1) {
					expectedValue = parts[0]; // Extract the date portion
				}
			}

			if (ignoreKeyList.contains(key.substring(key.lastIndexOf(".") + 1))) {
				flattenResponseJsonMap.remove(key);
				valueErrorJsonMap.remove(key);
				iterator.remove();
			}

			// Remove key-value pairs common in both and compare date values
			if (flattenResponseJsonMap.containsKey(key)) {
				String responseValue = String.valueOf(flattenResponseJsonMap.get(key));

				// Extract the date from the response value if the key contains "date"
				if (key.contains("Date")) {
					String[] parts = responseValue.split("\\+");
					if (parts.length > 1) {
						responseValue = parts[0]; // Extract the date portion
					}
				}

				if (responseValue.equals(expectedValue)) {
					flattenResponseJsonMap.remove(key);
					valueErrorJsonMap.remove(key);
					iterator.remove();
				}
			}
		}

		if (flattenExpectedJsonMap.isEmpty() && flattenResponseJsonMap.isEmpty()) {
			compareJsonFlag = false;
		} else {
			compareJsonFlag = true;
		}

		if (compareJsonFlag) {
			// keyset containing extra keys in response
			extraResponseKeys = new HashSet<>(flattenExpectedJsonMap.keySet());
			extraResponseKeys.addAll(flattenResponseJsonMap.keySet());
			extraResponseKeys.removeAll(flattenExpectedJsonMap.keySet());

			// keyset containing extra keys in expected
			extraExpectedKeys = new HashSet<>(flattenExpectedJsonMap.keySet());
			extraExpectedKeys.addAll(flattenResponseJsonMap.keySet());
			extraExpectedKeys.removeAll(flattenResponseJsonMap.keySet());

			// same key, different values, stores as response/expected
			for (String k : extraExpectedKeys) {
				valueErrorJsonMap.remove(k);
			}
			for (Entry<String, Object> entry1 : valueErrorJsonMap.entrySet()) {
				valueErrorJsonMap.put(entry1.getKey(), flattenResponseJsonMap.get(entry1.getKey()) + "/"
						+ flattenExpectedJsonMap.get(entry1.getKey()));
			}
		}
		return compareJsonFlag;
	}

	// creates excel sheet for comparing json
	public static String compareSheetCreator(Map<String, Object> flattenExpectedJsonMap,
			Map<String, Object> flattenResponseJsonMap, Map<String, Object> valueErrorJsonMap,
			HashSet<String> extraExpectedKeys, HashSet<String> extraResponseKeys, String expectedJsonPath,
			String responseJsonPath) throws Exception {

		try {
			log.info("Creating the Error Result xlsx sheet...");
			String filename = expectedJsonPath.substring(expectedJsonPath.lastIndexOf("\\") + 1,
					expectedJsonPath.lastIndexOf("_")) + "_Result";
			String filepath = expectedJsonPath.substring(0, expectedJsonPath.lastIndexOf("\\") + 1) + filename
					+ ".xlsx";

			int rownum = 0;
			int cellnum = 0;
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet compareSheet = workbook.createSheet(filename + " Comparison");

			Hyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.FILE);
			Hyperlink link1 = workbook.getCreationHelper().createHyperlink(HyperlinkType.FILE);
			Row row = null;
			Cell cell = null;
			CellStyle style = null;

			Map<String, Object> combineJson = flattenExpectedJsonMap;
			combineJson.putAll(flattenResponseJsonMap);
			TreeSet<String> treeSet = new TreeSet<String>(combineJson.keySet());
			HashSet<String> valueErrorKeySet = new HashSet<>(valueErrorJsonMap.keySet());
			LinkedHashMap<String, Integer> rootMap = new LinkedHashMap<String, Integer>();
			rootMapCreator(treeSet, rootMap);

			for (int i = 0; i < 6; i++) {
				row = compareSheet.createRow(i);
				for (int j = 0; j < 25; j++) {
					cell = row.createCell(j);
					style = ExcelUtility.getCellStyleWhite(workbook);
					if (i > 0 && j > 1 && j < 12 && i < 5)
						ExcelUtility.setAllBorders(style);
					cell.setCellStyle(style);
				}
			}

			// header section and sheet formatting
			compareSheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 6));
			compareSheet.addMergedRegion(new CellRangeAddress(2, 2, 1, 6));
			compareSheet.addMergedRegion(new CellRangeAddress(3, 3, 1, 6));
			compareSheet.addMergedRegion(new CellRangeAddress(4, 4, 1, 6));
			compareSheet.addMergedRegion(new CellRangeAddress(1, 1, 9, 11));
			compareSheet.addMergedRegion(new CellRangeAddress(2, 2, 9, 11));
			compareSheet.addMergedRegion(new CellRangeAddress(3, 3, 9, 11));
			compareSheet.addMergedRegion(new CellRangeAddress(4, 4, 9, 11));
			compareSheet.addMergedRegion(new CellRangeAddress(1, 4, 8, 8));

			row = compareSheet.getRow(rownum++);
			row = compareSheet.getRow(rownum++);
			cell = row.getCell(1);
			cell.setCellValue("Color Coding");
			style = ExcelUtility.getXPathStyle(workbook);
			ExcelUtility.setAllBorders(style);
			style.setAlignment(HorizontalAlignment.CENTER);
			cell.setCellStyle(style);

			cell = row.getCell(7);
			cell.setCellValue("Count");
			cell.setCellStyle(style);

			cell = row.getCell(9);
			cell.setCellValue("JSON File Path");
			cell.setCellStyle(style);

			row = compareSheet.getRow(rownum++);
			cell = row.getCell(1);
			cell.setCellValue("If key exists only in Expected Json");
			style = ExcelUtility.getGreenStyle(workbook);
			ExcelUtility.setAllBorders(style);
			style.setAlignment(HorizontalAlignment.CENTER);
			cell.setCellStyle(style);

			cell = row.getCell(7);
			cell.setCellValue(extraExpectedKeys.size());
			cell.setCellStyle(style);
			style = ExcelUtility.getXPathStyle(workbook);
			ExcelUtility.setAllBorders(style);
			style.setAlignment(HorizontalAlignment.CENTER);

			cell = row.getCell(9);
			cell.setCellValue("Expected JSON File");
			link1.setAddress(expectedJsonPath.replace("\\", "/"));
			cell.setHyperlink(link1);
			cell.setCellStyle(ExcelUtility.getHyperLinkStyle(workbook));

			row = compareSheet.getRow(rownum++);
			cell = row.getCell(1);
			cell.setCellValue("If key exists only in Response Json");
			style = ExcelUtility.getHeaderStyleAqua(workbook);
			ExcelUtility.setAllBorders(style);
			style.setAlignment(HorizontalAlignment.CENTER);
			cell.setCellStyle(style);

			cell = row.getCell(7);
			cell.setCellValue(extraResponseKeys.size());
			cell.setCellStyle(style);

			style = ExcelUtility.getXPathStyle(workbook);
			ExcelUtility.setAllBorders(style);
			style.setAlignment(HorizontalAlignment.CENTER);
			cell = row.getCell(9);
			cell.setCellValue("Response JSON File");
			link.setAddress(responseJsonPath.replace("\\", "/"));
			cell.setHyperlink(link);
			cell.setCellStyle(ExcelUtility.getHyperLinkStyle(workbook));

			row = compareSheet.getRow(rownum++);
			cell = row.getCell(1);
			cell.setCellValue("Value difference, actual_value/expected_value");
			style = ExcelUtility.getRedStyle(workbook);
			ExcelUtility.setAllBorders(style);
			style.setAlignment(HorizontalAlignment.CENTER);
			cell.setCellStyle(style);

			cell = row.getCell(7);
			cell.setCellValue(valueErrorJsonMap.size());
			cell.setCellStyle(style);

			row = compareSheet.getRow(rownum++);
			row = compareSheet.getRow(rownum++);

			// writing json Path in sheet and storing rownum corresponding to every json
			// Path
			for (String k : rootMap.keySet()) {
				style = ExcelUtility.getXPathStyle(workbook);
				cellnum = 0;
				row = compareSheet.createRow(rownum++);
				rootMap.put(k, rownum - 1);
				cell = row.createCell(cellnum++);
				cell.setCellValue("$");
				cell.setCellStyle(style);
				cell = row.createCell(cellnum);
				cell.setCellValue(k);
				cell.setCellStyle(style);
				row = compareSheet.createRow(rownum++);
				row = compareSheet.createRow(rownum++);
				row = compareSheet.createRow(rownum++);
			}

			// supply all 3 maps individually to fill sheet
			fillCompareSheet(rootMap, workbook, compareSheet, extraExpectedKeys, flattenExpectedJsonMap,
					ExcelUtility.getGreenStyle(workbook));
			fillCompareSheet(rootMap, workbook, compareSheet, extraResponseKeys, flattenResponseJsonMap,
					ExcelUtility.getHeaderStyleAqua(workbook));
			fillCompareSheet(rootMap, workbook, compareSheet, valueErrorKeySet, valueErrorJsonMap,
					ExcelUtility.getRedStyle(workbook));

			try {
				File existingResultFile = new File(filepath);
				if (existingResultFile.exists())
					existingResultFile.delete();
				FileOutputStream out = new FileOutputStream(new File(filepath));
				workbook.write(out);
				out.close();
				File f = new File(filepath);
				f.setReadOnly();
			} catch (Exception exObj) {
				System.out.println("Error occured while writing file to disk: " + exObj.getMessage());
			}
			return filepath;
		} catch (Throwable th) {
			th.printStackTrace(); // Print the stack trace to the console
			log.error("Exception occurred while creating the Excel sheet: " + th.getMessage());
			throw new Exception("Exception occurred while creating the Excel sheet", th); // Rethrow the exception to
																							// propagate it further
		}
	}

	// map that holds json Path(path before last seperator ".") and rownum where it
	// is written
	public static void rootMapCreator(TreeSet<String> treeSet, HashMap<String, Integer> rootMap) throws Exception {
		String separator = ".";
		int sepPos = 0;
		Iterator<String> responseIterator = treeSet.iterator();
		while (responseIterator.hasNext()) {
			String str = responseIterator.next();
			if (str.contains(".")) {
				sepPos = str.lastIndexOf(separator);
				String xPath = str.substring(0, sepPos);
				if (!rootMap.containsKey(xPath))
					rootMap.put("root." + xPath, 0);
			} else
				rootMap.put("root." + str, 0);
		}
	}

	// fills the sheet with values corresponding to supplied json
	// Commenting for error -Exception occurred while creating the Excel sheet: The maximum length of cell contents (text) is 32,767 characters
	/*
	 * public static void fillCompareSheet(LinkedHashMap<String, Integer> rootMap,
	 * XSSFWorkbook workbook, XSSFSheet sheet, HashSet<String> Keys, Map<String,
	 * Object> flattenedJsonMap, CellStyle style) { Iterator<String> iterator =
	 * Keys.iterator(); Row row = null; Cell cell = null; int rownum = 0; while
	 * (iterator.hasNext()) { String temp = iterator.next(); if (temp.contains("."))
	 * { rownum = rootMap.get("root." + temp.substring(0, temp.lastIndexOf(".")));
	 * row = sheet.getRow(rownum + 1); cell =
	 * row.createCell(row.getPhysicalNumberOfCells() + 1);
	 * cell.setCellValue(temp.substring(temp.lastIndexOf(".") + 1, temp.length()));
	 * row = sheet.getRow(rownum + 2); cell =
	 * row.createCell(row.getPhysicalNumberOfCells() + 1);
	 * cell.setCellValue(flattenedJsonMap.get(temp).toString());
	 * cell.setCellStyle(style); } else { rownum = rootMap.get("root." + temp); row
	 * = sheet.getRow(rownum); cell = row.getCell(1); cell.setCellValue("root");
	 * cell.setCellStyle(ExcelUtility.getXPathStyle(workbook)); row =
	 * sheet.getRow(rownum + 1); cell =
	 * row.createCell(row.getPhysicalNumberOfCells() + 1); cell.setCellValue(temp);
	 * row = sheet.getRow(rownum + 2); cell =
	 * row.createCell(row.getPhysicalNumberOfCells() + 1);
	 * cell.setCellValue(flattenedJsonMap.get(temp).toString());
	 * cell.setCellStyle(style); } } }
	 */
	public static void fillCompareSheet(LinkedHashMap<String, Integer> rootMap, XSSFWorkbook workbook, XSSFSheet sheet,
			HashSet<String> Keys, Map<String, Object> flattenedJsonMap, CellStyle style) {
		Iterator<String> iterator = Keys.iterator();
		Row row = null;
		Cell cell = null;
		int rownum = 0;

		while (iterator.hasNext()) {
			String temp = iterator.next();
			String value = String.valueOf(flattenedJsonMap.get(temp));
			if (value == null)
				value = "";

			if (temp.contains(".")) {
				rownum = rootMap.get("root." + temp.substring(0, temp.lastIndexOf(".")));
				row = sheet.getRow(rownum + 1);
				cell = row.createCell(row.getPhysicalNumberOfCells() + 1);
				cell.setCellValue(temp.substring(temp.lastIndexOf(".") + 1));
				row = sheet.getRow(rownum + 2);

// Handle long content
				if (value.length() > 32767) {
					cell = row.createCell(row.getPhysicalNumberOfCells() + 1);
					cell.setCellValue(value.substring(0, 32767)); // Truncate to max allowed length
					log.warn("Truncated cell content for key {} as it exceeded the maximum length", temp);
				} else {
					cell = row.createCell(row.getPhysicalNumberOfCells() + 1);
					cell.setCellValue(value);
				}

				cell.setCellStyle(style);
			} else {
				rownum = rootMap.get("root." + temp);
				row = sheet.getRow(rownum);
				cell = row.getCell(1);
				cell.setCellValue("root");
				cell.setCellStyle(ExcelUtility.getXPathStyle(workbook));
				row = sheet.getRow(rownum + 1);
				cell = row.createCell(row.getPhysicalNumberOfCells() + 1);
				cell.setCellValue(temp);
				row = sheet.getRow(rownum + 2);

// Handle long content
				if (value.length() > 32767) {
					cell = row.createCell(row.getPhysicalNumberOfCells() + 1);
					cell.setCellValue(value.substring(0, 32767)); // Truncate to max allowed length
					log.warn("Truncated cell content for key {} as it exceeded the maximum length", temp);
				} else {
					cell = row.createCell(row.getPhysicalNumberOfCells() + 1);
					cell.setCellValue(value);
				}

				cell.setCellStyle(style);
			}
		}
	}

}