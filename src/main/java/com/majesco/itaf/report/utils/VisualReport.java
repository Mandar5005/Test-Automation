package com.majesco.itaf.report.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import javafx.application.Application;
//import org.apache.logging.log4j.LogManager;//Sel4
//import org.apache.logging.log4j.Logger;//Sel4
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.majesco.itaf.main.Config;

public class VisualReport {
	//private final static Logger log = LogManager.getLogger(VisualReport.class.getName());
	static boolean sortByGroup = false;
	static {
		if (Config.reportSortByGroupID != null && Config.reportSortByGroupID.equalsIgnoreCase("True"))
			sortByGroup = true;
	}
	static FileInputStream file = null;
	static XSSFWorkbook SummaryResults = new XSSFWorkbook();

	public static void generateVisualReport() throws Exception {
		Functions.clearDirectory(Config.resultFilePath + "\\" + "charts");
		new File(Config.resultFilePath + "\\" + "charts").mkdir();

		String generatedXlsxFilePath = Functions.convertCsvToXlsx(SummaryResults, Config.resultFilePath,
				Config.resultOutput);
		File visualReport = new File(generatedXlsxFilePath.trim());
		if (visualReport.exists())
			visualReport.delete();
		XSSFSheet summarySheet = SummaryResults.getSheetAt(SummaryResults.getActiveSheetIndex());
		MapGenerator.generateSummaryMap(summarySheet);
		if (MapGenerator.SummaryMap.size() > 100)
			throw new Exception("Size exceeded");
		SummaryResults.createSheet("Visualised");
		SummaryResults.createSheet("Summarized");
		Functions.setBackgroundWhite(SummaryResults, "Visualised", 0, 0, 200, 200);
		SummarySheetCreator.fillSummarySheet(SummaryResults, MapGenerator.SummaryMap, "Summarized");
		Application.launch(VisualSheetCreator.class);
		VisualSheetCreator.visualSheetLayout(SummaryResults);
		FileOutputStream out = new FileOutputStream(new File(generatedXlsxFilePath.trim()));
		SummaryResults.write(out);
		out.close();
		File f = new File(generatedXlsxFilePath.trim());
		f.setReadOnly();
		// log.info("File saved to disk");
		SummaryResults.close();
	}
}
