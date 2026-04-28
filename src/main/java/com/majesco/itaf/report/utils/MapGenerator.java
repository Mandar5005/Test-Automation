package com.majesco.itaf.report.utils;

import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.commons.lang.StringUtils;
//import org.apache.logging.log4j.LogManager;//Sel4
//import org.apache.logging.log4j.Logger;//Sel4
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import com.majesco.itaf.main.Config;

public class MapGenerator {
	//private final static Logger log = LogManager.getLogger(MapGenerator.class.getName());//Sel4
	static int IDColumnNumber = 0;
	static int StatusColumnNumber = 0;
	static int TimeColumnNumber = 0;
	static int RowCount = 0;
	static int ID_flag = 0;
	static String TestID = null;

	static LinkedHashMap<String, SummaryObject> SummaryMap = null;

	public static void generateSummaryMap(XSSFSheet sheet) throws IOException {
		HashMap<String, Integer> columnNumberMap = Functions.getValueFromHashMap(sheet);
		StatusColumnNumber = columnNumberMap.get("Status") + 1;
		for (String timeColumnHeader : columnNumberMap.keySet()) {
			if (timeColumnHeader.equalsIgnoreCase("ElapsedTime"))
				TimeColumnNumber = columnNumberMap.get("ElapsedTime") + 1;
			else if (timeColumnHeader.equalsIgnoreCase("ExecutionTime(Seconds)"))
				TimeColumnNumber = columnNumberMap.get("ExecutionTime(Seconds)") + 1;
		}
		if (Config.executionApproach.equalsIgnoreCase("Linear") && VisualReport.sortByGroup) {
			IDColumnNumber = columnNumberMap.get("GroupName") + 1;
		} else {
			IDColumnNumber = columnNumberMap.get("TestCaseID") + 1;
		}

		Iterator<Row> rowIterator1 = sheet.iterator();
		RowCount = 0;
		SummaryMap = new LinkedHashMap<String, SummaryObject>();
		int idEntryExistFlag = 0;
		while (rowIterator1.hasNext()) {
			int cellno = 0;
			RowCount = RowCount + 1;
			Row row = rowIterator1.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			idEntryExistFlag = 0;
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				cellno = cellno + 1;
				/* new change */
				if (cellno == IDColumnNumber && RowCount > 1) {
					TestID = cell.getStringCellValue();

					for (Map.Entry<String, SummaryObject> m : SummaryMap.entrySet()) {
						if (TestID.equalsIgnoreCase(m.getKey().toString())) {
							idEntryExistFlag = 1;
						}
					}
					if (idEntryExistFlag == 0) {
						if (!Functions.ignoreTestCaseList.contains(cell.getStringCellValue())) {
							if (Config.executionApproach.equalsIgnoreCase("TimeTravel")) {
								if (TestID.contains("_")) {
									SummaryMap.put(StringUtils.substringBefore(TestID, "_"), new SummaryObject());
								} else {
									SummaryMap.put(TestID, new SummaryObject());
								}
							}

							else
								SummaryMap.put(TestID, new SummaryObject());
							idEntryExistFlag = 0;
						}

					}
				}
			}
		}
		SummaryObject obj = null;
		for (Map.Entry<String, SummaryObject> m : SummaryMap.entrySet()) {
			String tempID = m.getKey().toString();
			obj = (SummaryObject) m.getValue();
			Iterator<Row> rowIterator2 = sheet.iterator();
			RowCount = 0;

			while (rowIterator2.hasNext()) {
				ID_flag = 0;
				int cellno = 0;
				RowCount = RowCount + 1;
				Row row = rowIterator2.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					cellno = cellno + 1;
					if (cellno == IDColumnNumber && RowCount > 1) {
						if (Config.executionApproach.equalsIgnoreCase("TimeTravel")) {
							String tempID1 = cell.getStringCellValue();
							if (tempID1.contains("_")) {
								tempID1 = StringUtils.substringBefore(cell.getStringCellValue(), "_");
							} else {
								tempID1 = cell.getStringCellValue();
							}
							if (tempID.equals(tempID1)) {
								ID_flag = 1;
							}
						} else if (Config.executionApproach.equalsIgnoreCase("Linear")) {
							String temp2 = cell.getStringCellValue();
							if (tempID.equals(temp2)) {
								ID_flag = 1;
							}
						}

					}

					if (cellno == StatusColumnNumber && ID_flag == 1) {
						if (cell.getStringCellValue().equalsIgnoreCase("PASS")) {
							obj.setPass(obj.getPass() + 1);
						}

						if (cell.getStringCellValue().equalsIgnoreCase("FAIL")) {
							obj.setFail(obj.getFail() + 1);
						}
					}

					if (cellno == TimeColumnNumber && ID_flag == 1) {
						String time = null;
						if (cell.getCellType() == CellType.STRING)
							time = cell.getStringCellValue();
						else if (cell.getCellType() == CellType.NUMERIC)
							time = String.valueOf(cell.getNumericCellValue());
						long milliseconds = Functions.converttomillis(time);
						obj.setEtime(obj.getEtime() + milliseconds);
					}
				}
			}
		}

		HashMap<String, Integer> totalCount = Functions.getTotalCount();
		for (Entry<String, SummaryObject> m : MapGenerator.SummaryMap.entrySet()) {
			for (Entry<String, Integer> m1 : totalCount.entrySet()) {
				if (m.getKey().equalsIgnoreCase(m1.getKey())) {
					obj = (SummaryObject) m.getValue();
					obj.setTotal(m1.getValue());
					int pending = obj.getTotal() - obj.getPass() - obj.getFail();
					obj.setPending(pending);
					m.setValue(obj);
				}
			}
		}
	}
}
