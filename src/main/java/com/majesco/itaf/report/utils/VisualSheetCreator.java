package com.majesco.itaf.report.utils;

import java.io.File;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.LinkedHashMap;
import java.util.Map;
import javax.imageio.ImageIO;

//import org.apache.logging.log4j.LogManager;//Sel4
//import org.apache.logging.log4j.Logger;//Sel4
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.majesco.itaf.main.Config;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.embed.swing.SwingFXUtils;
import javafx.scene.Scene;
import javafx.scene.chart.*;
import javafx.scene.image.WritableImage;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import org.apache.poi.ss.util.CellRangeAddress;

public class VisualSheetCreator extends Application {
	//private final static Logger log = LogManager.getLogger(VisualSheetCreator.class.getName());
	static int barGraphCount = 0;

	public void start(Stage primaryStage) throws Exception {
		int scenarioCount = 0;
		String chartLocation = Config.resultFilePath + "//" + "charts";
		LinkedHashMap<String, SummaryObject> SummaryMap = MapGenerator.SummaryMap;

		primaryStage.setTitle("Visualised");
		CategoryAxis xAxis = new CategoryAxis();
		xAxis.setLabel("Test Scenarios");
		xAxis.getCategories().addAll("Pass", "Fail", "Pending");
		xAxis.autosize();

		@SuppressWarnings("rawtypes")
		Axis yAxis = new NumberAxis();
		yAxis.setLabel("Count");

		CategoryAxis timexAxis = new CategoryAxis();
		timexAxis.setLabel("Test Scenarios");

		@SuppressWarnings("rawtypes")
		Axis timeYAxis = new NumberAxis();
		timeYAxis.setLabel("Elapsed Time(minutes)");

		XYChart.Series<String, Long> TimeSeries = new XYChart.Series<String, Long>();
		TimeSeries.setName("Elapsed Time in minutes");

		@SuppressWarnings("unchecked")
		LineChart<String, Long> timeChart = new LineChart<String, Long>(timexAxis, timeYAxis);
		timeChart.setMinHeight(450);
		timeChart.getData().add(TimeSeries);
		timeChart.setTitle("Time Comparision between various test scenarios");
		timeChart.setAnimated(false);

		XYChart.Series<String, Integer> PassSeries = new XYChart.Series<String, Integer>();
		PassSeries.setName("Pass");

		XYChart.Series<String, Integer> FailSeries = new XYChart.Series<String, Integer>();
		FailSeries.setName("Fail");

		XYChart.Series<String, Integer> PendingSeries = new XYChart.Series<String, Integer>();
		PendingSeries.setName("Pending");

		@SuppressWarnings("unchecked")
		StackedBarChart<String, Integer> chart = new StackedBarChart<String, Integer>(xAxis, yAxis);
		chart.setMinHeight(450);
		chart.setTitle("Comparision between various test scenarios");
		chart.getData().add(PassSeries);
		chart.getData().add(FailSeries);
		chart.getData().add(PendingSeries);
		PassSeries.getData().clear();
		FailSeries.getData().clear();
		PendingSeries.getData().clear();

		for (Map.Entry<String, SummaryObject> m : MapGenerator.SummaryMap.entrySet()) {
			scenarioCount++;
			SummaryObject obj = (SummaryObject) m.getValue();
			PassSeries.getData().add(new XYChart.Data<String, Integer>(m.getKey(), (Integer) obj.getPass()));
			FailSeries.getData().add(new XYChart.Data<String, Integer>(m.getKey(), (Integer) obj.getFail()));
			PendingSeries.getData().add(new XYChart.Data<String, Integer>(m.getKey(), (Integer) obj.getPending()));
			TimeSeries.getData().add(new XYChart.Data<String, Long>(m.getKey(), (obj.getEtime() / 60000)));
			if (scenarioCount % 50 == 0 || scenarioCount == SummaryMap.size()) {
				barGraphCount++;

				if (scenarioCount == SummaryMap.size()) {
					for (int i = 0; i < (50 - (scenarioCount % 50)); i++) {
						PassSeries.getData()
						.add(new XYChart.Data<String, Integer>("NA".concat(Integer.toString(i)), 0));
						FailSeries.getData()
						.add(new XYChart.Data<String, Integer>("NA".concat(Integer.toString(i)), 0));
						PendingSeries.getData()
						.add(new XYChart.Data<String, Integer>("NA".concat(Integer.toString(i)), 0));
					}
				}
				String filename = "StackedBarChart_";
				filename = filename.concat(Integer.toString(barGraphCount));
				filename = filename.concat(".png");
				chart.setAnimated(false);
				saveBarAsImage(chart, filename, primaryStage, chartLocation);
				PassSeries.getData().clear();
				FailSeries.getData().clear();
				PendingSeries.getData().clear();
				switch (barGraphCount) {
				case 1:
					Functions.importImage(chartLocation, VisualReport.SummaryResults, "Visualised", filename, 1, 0, 23,
							10);
					break;
				case 2:
					Functions.importImage(chartLocation, VisualReport.SummaryResults, "Visualised", filename, 24, 0, 46,
							10);
					break;
				default:

				}
			}
		}
		saveBarAsImage(timeChart, "TimeGraph.png", primaryStage, chartLocation);

		Stage pieStage = new Stage();
		PieChart TransPieChart = new PieChart();
		TransPieChart.setTitle("Total Transaction Summary");
		PieChart.Data slice1 = new PieChart.Data("Passed transactions", SummarySheetCreator.pass_trans);
		PieChart.Data slice2 = new PieChart.Data("Failed transactions", SummarySheetCreator.fail_trans);
		PieChart.Data slice3 = new PieChart.Data("Pending transactions", SummarySheetCreator.pending_trans);
		TransPieChart.getData().add(slice1);
		TransPieChart.getData().add(slice2);
		TransPieChart.getData().add(slice3);
		TransPieChart.setAnimated(false);
		savePieAsImage(TransPieChart, "TransactionsPiechart.png", pieStage, chartLocation);

		PieChart SCPieChart = new PieChart();
		SCPieChart.setTitle("Total SC Summary");
		PieChart.Data IDPassSlice = new PieChart.Data("Passed scenarios", SummarySheetCreator.passed_ID);
		PieChart.Data IDFailSlice = new PieChart.Data("Failed scenarios", SummarySheetCreator.failed_ID);
		// PieChart.Data IDPendingSlice = new PieChart.Data("Pending scenarios" ,
		// SummarySheetCreator.pending_ID);
		SCPieChart.getData().add(IDPassSlice);
		SCPieChart.getData().add(IDFailSlice);
		// SCPieChart.getData().add(IDPendingSlice);
		SCPieChart.setAnimated(false);
		savePieAsImage(SCPieChart, "ScenariosPiechart.png", pieStage, chartLocation);

		if (barGraphCount > 1) {
			Functions.importImage(chartLocation, VisualReport.SummaryResults, "Visualised", "timeGraph.png", 47, 0, 68,
					10);
			Functions.importImage(chartLocation, VisualReport.SummaryResults, "Visualised", "TransactionsPiechart.png",
					70, 0, 88, 10);
			Functions.importImage(chartLocation, VisualReport.SummaryResults, "Visualised", "ScenariosPiechart.png", 70,
					10, 88, 21);
		} else {
			Functions.importImage(chartLocation, VisualReport.SummaryResults, "Visualised", "timeGraph.png", 24, 0, 45,
					10);
			Functions.importImage(chartLocation, VisualReport.SummaryResults, "Visualised", "TransactionsPiechart.png",
					47, 0, 63, 10);
			Functions.importImage(chartLocation, VisualReport.SummaryResults, "Visualised", "ScenariosPiechart.png", 47,
					10, 63, 21);
		}
		Platform.exit();
		// log.info("SummaryResults: Visualised sheet written successfully on disk.");
	}

	/* create stackedbarchart image */
	public void saveBarAsImage(StackedBarChart<String, Integer> chart, String filename, Stage stage, String imgLocation)
			throws IOException {
		VBox vbox1 = new VBox(chart);
		Scene scene1 = new Scene(vbox1, 1125, 450);
		stage.setScene(scene1);
		saveImage(scene1, imgLocation, filename);
	}

	/* create linechart image */
	public void saveBarAsImage(LineChart<String, Long> chart, String filename, Stage stage, String imgLocation)
			throws IOException {
		VBox vbox1 = new VBox(chart);
		Scene scene1 = new Scene(vbox1, 1125, 450);
		stage.setScene(scene1);
		saveImage(scene1, imgLocation, filename);
	}

	/* create pie chart image */
	public void savePieAsImage(PieChart chart, String filename, Stage stage, String imgLocation) throws IOException {
		VBox vbox1 = new VBox(chart);
		Scene scene1 = new Scene(vbox1, 562, 360);
		stage.setScene(scene1);
		saveImage(scene1, imgLocation, filename);
	}

	/* save chart as image to given location */
	public void saveImage(Scene scene, String imgLocation, String filename) throws IOException {
		WritableImage snapShot = scene.snapshot(null);
		ImageIO.write(SwingFXUtils.fromFXImage(snapShot, null), "png", new File(imgLocation + "\\" + filename));
	}

	/* sets entire layout of visualised sheet in SumaryResults */
	public static void visualSheetLayout(XSSFWorkbook workbook) throws IOException {
		int rownum = 0;
		int cellnum = 0;
		XSSFSheet sheet = workbook.getSheet("Visualised");
		Row row = sheet.createRow(rownum);
		Cell cell = row.createCell(cellnum);
		CellStyle style = StyleGenerator.getHeaderStyleAqua(workbook);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		StyleGenerator.setFontBold(workbook, style, 16);
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 19));
		cell.setCellStyle(style);
		String reportHeading = null;
		if (VisualReport.sortByGroup)
			reportHeading = "GRAPHICAL OVERVIEW OF SUMMARY RESULTS(By GroupName)";
		else
			reportHeading = "GRAPHICAL OVERVIEW OF SUMMARY RESULTS(By TestCaseID)";
		cell.setCellValue(reportHeading);

		float pass_trans = (float) SummarySheetCreator.pass_trans;
		float fail_trans = (float) SummarySheetCreator.fail_trans;
		float pending_trans = (float) SummarySheetCreator.pending_trans;
		float total = (float) (pass_trans + fail_trans + pending_trans);

		float pass_percentage = ((pass_trans / total) * 100);
		float fail_percentage = ((fail_trans / total) * 100);
		float pending_percentage = ((pending_trans / total) * 100);

		float passed_ID = SummarySheetCreator.passed_ID;
		float failed_ID = SummarySheetCreator.failed_ID;
		// float pending_ID= SummarySheetCreator.pending_ID;
		float total_ID = MapGenerator.SummaryMap.size();

		float passed_ID_percentage = ((passed_ID / total_ID) * 100);
		float failed_ID_percentage = ((failed_ID / total_ID) * 100);
		// float pending_ID_percentage= ((pending_ID/total_ID)*100);

		CellStyle greenStyle = StyleGenerator.getGreenStyle(workbook);
		CellStyle orangeStyle = StyleGenerator.getOrangeStyle(workbook);
		CellStyle yellowStyle = StyleGenerator.getYellowStyle(workbook);
		DecimalFormat df = new DecimalFormat("0.00");

		if (barGraphCount > 1) {
			rownum = 88;
			cellnum = 4;
			row = sheet.getRow(rownum);
			cell = row.createCell(cellnum++);
			cell.setCellValue(df.format(pass_percentage) + "%");
			cell.setCellStyle(orangeStyle);

			cell = row.createCell(cellnum++);
			cell.setCellValue(df.format(fail_percentage) + "%");
			cell.setCellStyle(yellowStyle);

			cell = row.createCell(cellnum++);
			cell.setCellValue(df.format(pending_percentage) + "%");
			cell.setCellStyle(greenStyle);

			cellnum = 14;

			cell = row.createCell(cellnum++);
			cell.setCellValue(df.format(passed_ID_percentage) + "%");
			cell.setCellStyle(orangeStyle);

			cell = row.createCell(cellnum++);
			cell.setCellValue(df.format(failed_ID_percentage) + "%");
			cell.setCellStyle(yellowStyle);

			// cell= row.createCell(cellnum++);
			// cell.setCellValue(df.format(pending_ID_percentage) +"%");
			// cell.setCellStyle(greenStyle);
		}

		else {
			rownum = 65;
			cellnum = 4;
			row = sheet.getRow(rownum);
			cell = row.createCell(cellnum++);
			cell.setCellValue(df.format(pass_percentage) + "%");
			cell.setCellStyle(orangeStyle);

			cell = row.createCell(cellnum++);
			cell.setCellValue(df.format(fail_percentage) + "%");
			cell.setCellStyle(yellowStyle);

			cell = row.createCell(cellnum++);
			cell.setCellValue(df.format(pending_percentage) + "%");
			cell.setCellStyle(greenStyle);

			cellnum = 14;

			cell = row.createCell(cellnum++);
			cell.setCellValue(df.format(passed_ID_percentage) + "%");
			cell.setCellStyle(orangeStyle);

			cell = row.createCell(cellnum++);
			cell.setCellValue(df.format(failed_ID_percentage) + "%");
			cell.setCellStyle(yellowStyle);

			// cell= row.createCell(cellnum++);
			// cell.setCellValue(df.format(pending_ID_percentage) +"%");
			// cell.setCellStyle(greenStyle);
		}
	}

}