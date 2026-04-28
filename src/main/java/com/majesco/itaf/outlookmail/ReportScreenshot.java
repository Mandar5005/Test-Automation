package com.majesco.itaf.outlookmail;

import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.image.BufferedImage;
import java.io.File;
import java.util.Calendar;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.aspose.cells.*;
import com.majesco.itaf.main.Config;

//Below given class will export a range from the excel to an image
public class ReportScreenshot {
	private final static Logger log = LogManager.getLogger(ReportScreenshot.class);

	public static String printexcel() throws Exception {

		String summaryReportName = null;
		String dataDir = getSharedDataDir(ReportScreenshot.class) + "Outlook/";
		String ExecutionReport = Config.executionStatusReportUtility;
		Workbook workbook = new Workbook(ExecutionReport);
		Worksheet worksheet = workbook.getWorksheets().get(1);
		worksheet.setZoom(100);
		worksheet.getPageSetup().setPrintTitleColumns("$B:$L");
		worksheet.getPageSetup().setPrintTitleColumns("$1:$2");
		worksheet.getPageSetup().setLeftMargin(0);
		worksheet.getPageSetup().setRightMargin(0);
		worksheet.getPageSetup().setTopMargin(0);
		worksheet.getPageSetup().setBottomMargin(0);

		log.info("Converting image to best fit outlook...");
		// Set OnePagePerSheet option as true
		ImageOrPrintOptions options = new ImageOrPrintOptions();
		options.setOnePagePerSheet(true);
		options.setImageType(ImageType.PNG);
		options.setHorizontalResolution(150);
		options.setVerticalResolution(150);
		// options.getQuality();
		options.setQuality(100);

		// Setting the print quality of the worksheet to 180 dpi
		PageSetup pageSetup = worksheet.getPageSetup();
		// pageSetup.setFitToPagesWide(0);
		// pageSetup.setZoom(100);
		pageSetup.setOrientation(PageOrientationType.PORTRAIT);
		// pageSetup.setPaperSize(PaperSizeType.PAPER_A_4);
		// pageSetup.setPrintQuality(800);

		Calendar cd = Calendar.getInstance();
		int currMonth = cd.get((Calendar.MONTH)) + 1;
		String s_prefix = "_" + cd.get(Calendar.DAY_OF_MONTH) + "_" + currMonth + "_" + cd.get(Calendar.YEAR) + "_"
				+ cd.get(Calendar.HOUR) + "_" + cd.get(Calendar.MINUTE) + "_" + cd.get(Calendar.SECOND);

		// Take the image of your worksheet
		SheetRender sr = new SheetRender(worksheet, options);
		sr.toImage(0, dataDir + "SummaryReport" + s_prefix + ".png");
		summaryReportName = ("SummaryReport" + s_prefix + ".png");

		// System.out.println("Screenprint taken successfully -" + summaryReportName);
		log.info("Screenprint taken successfully -" + summaryReportName);

		return summaryReportName;
	}

	public static String getSharedDataDir(Class<?> c) {

		File dir = new File(Config.executionStatusReportPath);
		File theDir = new File(Config.executionStatusReportPath + "/Outlook/");
		if (!theDir.exists()) {
			theDir.mkdir();
		}
		return dir.toString() + File.separator;
	}

	@SuppressWarnings("unused")
	private static BufferedImage resize(BufferedImage img, int height, int width) {
		Image tmp = img.getScaledInstance(width, height, Image.SCALE_SMOOTH);
		BufferedImage resized = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);

		Graphics2D g2d = resized.createGraphics();
		g2d.drawImage(tmp, 0, 0, null);
		g2d.dispose();
		return resized;
	}

}
