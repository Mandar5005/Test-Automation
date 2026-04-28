package com.majesco.itaf.util;

import java.io.File;
import java.io.IOException;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.majesco.compare.CompareUtil;
import com.majesco.compare.ResultBean.SummaryAttributes;
import com.majesco.compare.pdf.PDFResultBean;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.ITAFWebDriver;
import com.majesco.itaf.main.MainController;
import com.majesco.itaf.main.ObjectFactory;

interface PrintInterface {
	public void printPDFSummary(String compareResult, String FormPageNum);
}

public class PDFComparisonUtil {
	private final static Logger log = LogManager.getLogger(PDFComparisonUtil.class.getName());
	public static String StringToIgnore = "";
	public static String comparisonMode = "";
	private static MainController controller = ObjectFactory.getMainController();
	public static ITAFWebDriver webDriverPDF = ITAFWebDriver.getInstance();

	public static void setStringToIgnore(String str) {
		if (str.isEmpty()) {
			// StringToIgnore = "";
			log.info("No input is provided to ignore during PDF comparison");
		} else {
			StringToIgnore = str.trim();
			log.info("String to ignore during PDF comparison is set to '" + StringToIgnore + "'");
		}
	}

	public static void setPDFComparisonMode(String str) {
		if (str.isEmpty()) {
			comparisonMode = "";
			log.info("PDF comparison mode set to default - 'CompareAllBookmarks'");
		} else {
			comparisonMode = str.trim();
			log.info("PDF comparison mode set to '" + comparisonMode + "'");
		}
	}

	public static Object[] PDFCompare(String logicalName, String controltype, String ctrlValue) throws Exception {

		long expectedFilePageCount = 0, actualFilePageCount = 0;

		// get path of files for comparison
		String[] temp = ctrlValue.split(";");
		String expFilePath = Config.expPdfFilePath + "\\" + temp[1];
		String actualFilePath = Config.actualPdfDownloadPath + "\\" + temp[1];

		// create TC_Trasnaction folder if not present
		String folderPath = Config.pdfCompResultPath + "\\" + controller.controllerTestCaseID + "\\"
				+ controller.controllerTransactionType;
		File createDir = new File(folderPath);
		if (!createDir.exists()) {
			createDir.mkdirs();
		}

		// create subfolder with timestamp inside folder created above
		String timeStampForFolder = new SimpleDateFormat("ddMMMMyyyy_HH_mm_ss").format(new Date());
		String subFolderPathForOutput = folderPath + "\\" + timeStampForFolder;
		File subFolder = new File(subFolderPathForOutput);
		if (!subFolder.exists()) {
			subFolder.mkdirs();
		}

		String[] args = new String[7];

		// file comparison
		args[0] = "pdf";

		// Expected file path
		args[1] = expFilePath;

		// Actual file path
		args[2] = actualFilePath;

		// path for properties file. Not useful for now
		args[3] = "";

		// Results folder path
		args[4] = subFolderPathForOutput;

		// semicolon separated string to exclude
		// args[5] = "80-CP-002810335-8;06-12-18";
		args[5] = StringToIgnore;

		// Bookmark comparison setting
		if (comparisonMode.equalsIgnoreCase("CompareAllBookmarks"))
			args[6] = "ALL";
		else if (comparisonMode.equalsIgnoreCase("CompareChildOnlyBookmarks"))
			args[6] = "LEAVES";
		else
			args[6] = "ALL";

		StringToIgnore = "";
		comparisonMode = "";

		// Check for presence of expected and actual files else throw error
		File expFile = new File(expFilePath);
		if (!expFile.exists()) {
			log.info("Expected PDF file missing - " + expFilePath);
			if (subFolder.exists()) {
				subFolder.delete();
			}
			webDriverPDF.getReport().setMessage("Expected PDF File missing - " + expFilePath);
			webDriverPDF.getReport().setStatus("FAIL");
			return (new String[0]);
			// throw new FileNotFoundException("Expected PDF File missing - "+
			// expFilePath);
		}
		File actualFile = new File(actualFilePath);
		if (!actualFile.exists()) {
			log.info("Actual PDF file missing - " + actualFilePath);
			if (subFolder.exists()) {
				subFolder.delete();
			}
			webDriverPDF.getReport().setMessage("Actual PDF File missing - " + actualFilePath);
			webDriverPDF.getReport().setStatus("FAIL");
			return (new String[0]);
			// throw new
			// FileNotFoundException("Actual PDF File missing - "+actualFilePath);
		}

		System.out.println("PDF comparison started at : " + new SimpleDateFormat("ddMMMM HH:mm:ss").format(new Date()));
		log.info("PDF comparison started at : " + new SimpleDateFormat("ddMMMM HH:mm:ss").format(new Date()));
		System.out.println("PDF comparison is in progess");
		log.info("PDF comparison is in progess");

		PDFResultBean pdfResultBean = new PDFResultBean();
		try {
			System.out.println("Time before calling Garbage Collection: " + new Timestamp(System.currentTimeMillis()));
			System.gc();
			System.out.println("Time after calling Garbage Collection: " + new Timestamp(System.currentTimeMillis()));
			pdfResultBean = (PDFResultBean) CompareUtil.compareFiles(args);
		} catch (Throwable e) {
			e.printStackTrace();
			log.error(e.getMessage(), e);
			webDriverPDF.getReport().setMessage("Error during PDF Comparison - " + e);
			webDriverPDF.getReport().setStatus("FAIL");
			return (new String[0]);
			// throw new Exception("Error during PDF Comparison - " +e);
		}
		System.out
				.println("PDF comparison completed at: " + new SimpleDateFormat("ddMMMM HH:mm:ss").format(new Date()));
		log.info("PDF comparison completed at: " + new SimpleDateFormat("ddMMMM HH:mm:ss").format(new Date()));

		expectedFilePageCount = pdfResultBean.getExpectedFilePageCount();
		actualFilePageCount = pdfResultBean.getActualFilePageCount();

		String returnVal1 = "", returnVal2 = "";
		if (pdfResultBean.isError()) {
			returnVal1 = "Mismatch found in " + pdfResultBean.getSummary().get(SummaryAttributes.NotMatched)
					+ " page(s) of PDF document.";
			returnVal2 = "\nExpected page Count: " + expectedFilePageCount + ", ";
			returnVal2 = returnVal2 + "\nActual page Count: " + actualFilePageCount + ", ";
			returnVal2 = returnVal2 + "\nTotal Pages Compared: "
					+ pdfResultBean.getSummary().get(SummaryAttributes.Compared) + ", ";
			returnVal2 = returnVal2 + "\nTotal Pages Matched: "
					+ pdfResultBean.getSummary().get(SummaryAttributes.Matched) + ", ";
			returnVal2 = returnVal2 + "\nResults saved at: " + subFolderPathForOutput;
		} else {

			if (expectedFilePageCount == 0 || actualFilePageCount == 0
					|| pdfResultBean.getSummary().get(SummaryAttributes.NotMatched) == null
					|| pdfResultBean.getSummary().get(SummaryAttributes.Compared) == null
					|| pdfResultBean.getSummary().get(SummaryAttributes.Matched) == null) {
				returnVal1 = "Issue in PDF Comparison.";
				returnVal2 = "\nExpected page Count: " + expectedFilePageCount + ", ";
				returnVal2 = returnVal2 + "\nActual page Count: " + actualFilePageCount + ", ";
				returnVal2 = returnVal2 + "\nTotal Pages Compared: "
						+ pdfResultBean.getSummary().get(SummaryAttributes.Compared) + ", ";
				returnVal2 = returnVal2 + "\nTotal Pages Not Matched: "
						+ pdfResultBean.getSummary().get(SummaryAttributes.NotMatched) + ", ";
				returnVal2 = returnVal2 + "\nTotal Pages Matched: "
						+ pdfResultBean.getSummary().get(SummaryAttributes.Matched) + ", ";
				returnVal2 = returnVal2 + "\nCheck if any result is generated at: " + subFolderPathForOutput;
			} else {
				returnVal1 = "PDF Comparison is successful, NO mismatch found.";
				returnVal1 = returnVal1 + "\nExpected page Count: " + expectedFilePageCount + ", ";
				returnVal1 = returnVal1 + "\nActual page Count: " + actualFilePageCount + ", ";
				returnVal1 = returnVal1 + "\nTotal Pages Compared: "
						+ pdfResultBean.getSummary().get(SummaryAttributes.Compared) + ", ";
				returnVal1 = returnVal1 + "\nTotal Pages Matched: "
						+ pdfResultBean.getSummary().get(SummaryAttributes.Matched) + ", ";
				returnVal2 = returnVal1;
				if (subFolder.exists()) {
					subFolder.delete();
				}
			}
		}

		Object[] result = new Object[4];
		// result[0] = String.valueOf(pdfResultBean.isError());
		result[0] = pdfResultBean;
		result[1] = returnVal2;
		result[2] = returnVal1;
		result[3] = subFolderPathForOutput;
		return result;
	}

	public static void writeToPDFComparisonSummary(String GroupName, String TestCaseID, String TransactionType,
			String CurDate, Object[] PDFResultSummary) throws IOException {
		PDFResultBean pdfResultBean = new PDFResultBean();
		pdfResultBean = (PDFResultBean) PDFResultSummary[0];
		String filePath = Config.resultFilePath + "\\PDFComparisonSummary.xlsx";

		try {
			ExcelUtility.createExcelIfNotPresent(filePath, "PDFComparisonSummary");
			int usedRow = ExcelUtility.getUsedRowCount(filePath, "PDFComparisonSummary");
			if (usedRow == 0) {
				String[] headers = { "GroupName", "TestCaseID", "TransactionType", "CurrentDate", "CompareResult",
						"FormName\\PageNum", "ResultPath" };
				ExcelUtility.writeRowToExcel(filePath, "PDFComparisonSummary", 0, 0, headers);
			}

			PrintInterface p = new PrintInterface() {
				public void printPDFSummary(String compareResult, String FormPageNum) {
					try {
						String[] temp = { GroupName, TestCaseID, TransactionType, CurDate, compareResult, FormPageNum,
								PDFResultSummary[3].toString() };
						int usedRow = ExcelUtility.getUsedRowCount(filePath, "PDFComparisonSummary");
						ExcelUtility.writeRowToExcel(filePath, "PDFComparisonSummary", usedRow, 0, temp);
					} catch (Exception e) {
						log.info("Error while writing to PDFComparisonSummary " + e.getStackTrace());
					}
				}
			};

			if (!(StringUtils.equalsIgnoreCase(pdfResultBean.getAdditionalBookmarks(), null)
					|| StringUtils.equalsIgnoreCase(pdfResultBean.getAdditionalBookmarks(), ""))) {
				String[] temp = pdfResultBean.getAdditionalBookmarks().split(",");
				for (int i = 0; i < temp.length; i++) {
					p.printPDFSummary("Additional Bookmark", temp[i]);
				}
			}
			if (!(StringUtils.equalsIgnoreCase(pdfResultBean.getMissingBookmarks(), null)
					|| StringUtils.equalsIgnoreCase(pdfResultBean.getMissingBookmarks(), ""))) {
				String[] temp = pdfResultBean.getMissingBookmarks().split(",");
				for (int i = 0; i < temp.length; i++) {
					p.printPDFSummary("Missing Bookmark", temp[i]);
				}
			}
			if (!(StringUtils.equalsIgnoreCase(pdfResultBean.getPositionMismatchBookmarks(), null)
					|| StringUtils.equalsIgnoreCase(pdfResultBean.getPositionMismatchBookmarks(), ""))) {
				String[] temp = pdfResultBean.getPositionMismatchBookmarks().split(",");
				for (int i = 0; i < temp.length; i++) {
					p.printPDFSummary("Bookmark Position Change", temp[i]);
				}
			}
			if (!(StringUtils.equalsIgnoreCase(pdfResultBean.getBookmarksWithDifference(), null)
					|| StringUtils.equalsIgnoreCase(pdfResultBean.getBookmarksWithDifference(), ""))) {
				String[] temp = pdfResultBean.getBookmarksWithDifference().split(",");
				for (int i = 0; i < temp.length; i++) {
					p.printPDFSummary("Content Mismatch", temp[i]);
				}
			}
			if (!(StringUtils.equalsIgnoreCase(pdfResultBean.getPageNoWithDifference(), null)
					|| StringUtils.equalsIgnoreCase(pdfResultBean.getPageNoWithDifference(), ""))) {
				String[] temp = pdfResultBean.getPageNoWithDifference().split(",");
				for (int i = 0; i < temp.length; i++) {
					p.printPDFSummary("Content Mismatch", "Bookmark not present. Difference on page : " + temp[i]);
				}
			}
		} catch (Throwable e) {
			log.error("Error while writing to PDFComparisonSummary " + e.getMessage());
			e.printStackTrace();
		}

	}

}
