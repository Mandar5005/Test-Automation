package com.majesco.itaf.report.utils;

import java.util.*;

//import org.apache.logging.log4j.LogManager;//Sel4
//import org.apache.logging.log4j.Logger;//Sel4
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

public class SummarySheetCreator {
	//private final static Logger log = LogManager.getLogger(SummarySheetCreator.class.getName());
	static int passed_ID = 0;
	static int failed_ID = 0;
	static int pass_trans = 0;
	static int fail_trans = 0;
	static int pending_trans = 0;
	static int total_trans = 0;

	static void fillSummarySheet(XSSFWorkbook workbook, LinkedHashMap<String, SummaryObject> summaryMap,
			String sheetname) {
		XSSFSheet sheet = workbook.getSheet(sheetname);
		int rownum = 0;
		int cellnum = 0;
		CellStyle style = StyleGenerator.getHeaderStyleAqua(workbook);
		CellStyle Sumstyle = StyleGenerator.getRegularSummaryStyle(workbook);
		String[] arr = { "ID", "Total", "Pass", "Fail", "Pending", "Time(HH:MM:SS)" };
		StyleGenerator.setFontBold(workbook, style, 12);
		Row row = sheet.createRow(rownum++);
		Cell cell;
		for (int i = 0; i < arr.length; i++) {
			cell = row.createCell(cellnum++);
			cell.setCellValue(arr[i]);
			cell.setCellStyle(style);
		}

		SummaryObject obj = null;
		for (Map.Entry<String, SummaryObject> m : summaryMap.entrySet()) {
			obj = (SummaryObject) m.getValue();
			pass_trans = pass_trans + obj.getPass();
			fail_trans = fail_trans + obj.getFail();
			pending_trans = pending_trans + obj.getPending();
			total_trans = total_trans + obj.getTotal();

			if (obj.getFail() == 0 & obj.getPass() > 0) {
				passed_ID++;
			}

			if (obj.getFail() > 0) {
				failed_ID++;
			}
		}

		for (Map.Entry<String, SummaryObject> m : summaryMap.entrySet()) {
			obj = (SummaryObject) m.getValue();
			row = sheet.createRow(rownum++);
			cellnum = 0;
			cell = row.createCell(cellnum++);
			cell.setCellValue(m.getKey().toString());
			cell.setCellStyle(Sumstyle);
			cell = row.createCell(cellnum++);
			cell.setCellValue(obj.getTotal());
			cell.setCellStyle(Sumstyle);
			cell = row.createCell(cellnum++);
			cell.setCellValue(obj.getPass());
			cell.setCellStyle(Sumstyle);
			cell = row.createCell(cellnum++);
			cell.setCellValue(obj.getFail());
			cell.setCellStyle(Sumstyle);
			cell = row.createCell(cellnum++);
			cell.setCellValue(obj.getPending());
			cell.setCellStyle(Sumstyle);
			cell = row.createCell(cellnum++);
			cell.setCellValue(Functions.convertSecondsToHMmSs(obj.getEtime() / 1000));
			cell.setCellStyle(Sumstyle);

		}

		cellnum = 0;
		row = sheet.createRow(rownum++);
		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue("Total Time");
		cell = row.createCell(cellnum + 4);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue(Functions.totalTime());
		row = sheet.createRow(rownum++);
		row = sheet.createRow(rownum++);
		row = sheet.createRow(rownum++);
		row = sheet.createRow(rownum++);

		fillOverview(0, sheet, Sumstyle, style);

		for (int i = 0; i < 20; i++) {
			sheet.autoSizeColumn(i);
		}
		// log.info("SummaryResults: Summarized sheet written successfully on disk");

	}

	public static void fillOverview(int rownum, XSSFSheet sheet, CellStyle Sumstyle, CellStyle style) {

		int cellnum = 8;
		Row row = sheet.getRow(rownum);
		Cell cell = row.createCell(cellnum);

		sheet.addMergedRegion(new CellRangeAddress(rownum, rownum, cellnum, cellnum + 3));
		cell = row.createCell(cellnum++);
		cell.setCellValue("Overview");
		cell.setCellStyle(style);
		cell = row.createCell(cellnum + 2);
		cell.setCellStyle(style);
		rownum += 1;

		cellnum = 8;
		row = sheet.getRow(rownum);
		sheet.addMergedRegion(new CellRangeAddress(rownum, rownum, cellnum, cellnum + 1));
		cell = row.createCell(cellnum++);
		cell.setCellValue("Scenarios");
		cell.setCellStyle(style);
		cell = row.createCell(cellnum++);
		cell.setCellStyle(style);

		cellnum = 10;
		row = sheet.getRow(rownum);
		sheet.addMergedRegion(new CellRangeAddress(rownum, rownum, cellnum, cellnum + 1));
		cell = row.createCell(cellnum++);
		cell.setCellValue("Transactions");
		cell.setCellStyle(style);
		cell = row.createCell(cellnum++);
		cell.setCellStyle(style);

		cellnum = 8;
		rownum += 1;
		row = sheet.getRow(rownum);
		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue("Passed Scenarios");
		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue(passed_ID);

		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue("Passed transactions");
		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue(pass_trans);

		rownum += 1;
		cellnum = 8;
		row = sheet.getRow(rownum);
		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue("Failed Scenarios");
		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue(failed_ID);

		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue("Failed transactions");
		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue(fail_trans);

		rownum += 1;
		cellnum = 8;
		row = sheet.getRow(rownum);

		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue("Pending Scenarios");
		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue("NA");

		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue("Pending transactions");
		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue(pending_trans);

		rownum += 1;
		cellnum = 8;
		row = sheet.getRow(rownum);

		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue("Total Scenarios");
		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue(MapGenerator.SummaryMap.size());

		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue("Total transactions");
		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		cell.setCellValue(total_trans);

		rownum += 1;
		cellnum = 8;
		row = sheet.getRow(rownum);

		sheet.addMergedRegion(new CellRangeAddress(rownum, rownum, cellnum, cellnum + 1));
		cell = row.createCell(cellnum++);
		cell.setCellValue("Total time(HH:MM:SS)");
		cell.setCellStyle(Sumstyle);
		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
		sheet.addMergedRegion(new CellRangeAddress(rownum, rownum, cellnum, cellnum + 1));
		cell = row.createCell(cellnum++);
		cell.setCellValue(Functions.totalTime());
		cell.setCellStyle(Sumstyle);
		cell = row.createCell(cellnum++);
		cell.setCellStyle(Sumstyle);
	}
}
