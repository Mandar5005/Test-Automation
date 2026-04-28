package com.majesco.itaf.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.Map.Entry;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.ITAFWebDriver;
import com.majesco.itaf.main.MainController;
import com.majesco.itaf.main.MainControllerBilling;
import com.majesco.itaf.main.MainControllerSuite;
import com.majesco.itaf.main.ObjectFactory;
import com.majesco.itaf.main.TransactionMapping;
import com.majesco.itaf.main.WebHelper;
import com.majesco.itaf.main.WebHelperCS;
import com.majesco.itaf.main.WebHelperClaims;
import com.majesco.itaf.main.WebHelperDM;
import com.majesco.itaf.main.WebHelperDS;
import com.majesco.itaf.main.WebHelperLNA;
import com.majesco.itaf.main.WebHelperPAS;
import com.majesco.itaf.main.WebHelperSuite;
import com.majesco.itaf.main.WebHelperUtil;
import com.majesco.itaf.rest.utils.JsonUtility;
import com.majesco.itaf.verification.WebVerification;
import com.majesco.itaf.vo.Reporter;
import com.majesco.itaf.webservice.WebService;

public class ExcelUtility {

	private final static Logger log = LogManager.getLogger(ExcelUtility.class.getName());
	public static char myChar = 34;
	public static int firstRow = 1;
	public static int DResult = 1;
	public static int dynamicNum = 0;
	public static ArrayList<Integer> TIvaluesheetrows = new ArrayList<Integer>();
	public static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	private static MainController controller = ObjectFactory.getMainController();
	private static boolean isClaimsApplication = ITAFWebDriver.isClaimsApplication();
	public static String tempStatus = null;
	public static int failedcount = 0;

	//For Billing --- TimeTravel Approach
	/**
	 * Retrieves data based on provided parameters.
	 * If TestCaseID or cycleDate is empty, handles the scenario accordingly.
	 *
	 * @param structurePath   Path to the structure file.
	 * @param FilePath        Path to the input file.
	 * @param TestCaseID      Test case identifier (can be empty).
	 * @param TransactionType Type of transaction.
	 * @param cycleDate       Date for the cycle (can be empty).
	 * @param operationType   Type of operation.
	 * @return The row of expected data.
	 */
	public static Row GetDataFromValues(String structurePath, String FilePath, String TestCaseID,
			String TransactionType, String cycleDate, String operationType)
			throws IOException, InterruptedException, Exception
	// For XLS to XLSX
	{
		Row expectedRow = null;
		String tempRecovery_Scenario = null;
		Cell transactionType = null;
		HashMap<String, Integer> headerValues = new HashMap<>();
		Cell cycleDate_Values = null;
		Cell executeFlag_Values = null;
		Cell testCaseID_Values = null;
		Boolean firstRowFound = false;
		int singlerowindex = 0;
		int executeFlagcount = 0;

		try {
			Row currentRow = null;
			Sheet valuesSheet = WebHelperUtil.getSheet(FilePath, "Values");

			firstRowFound = false;
			ArrayList<Integer> valuesheetrows = new ArrayList<Integer>();

			int rowCount = valuesSheet.getLastRowNum() + 1;
			@SuppressWarnings("unused")
			int rowCount1 = valuesSheet.getPhysicalNumberOfRows();
			log.info("Number of rows in value sheet are:" + rowCount);
			log.info("Testcase ID:" + TestCaseID);
			log.info("Transaction Type:" + TransactionType);
			log.info("Business Date:" + cycleDate);
			if (rowCount == 0) {
				// MainController.pauseExecution = true;
				controller.pauseFun("No data for transaction " + TransactionType + " TestCase" + TestCaseID
						+ " Cycle date" + cycleDate + " in Values sheet");
			}
			for (int rowIndex = 1; rowIndex < rowCount && !controller.pauseExecution; rowIndex++) {

				// Time Travel START
				currentRow = valuesSheet.getRow(rowIndex);

				if (headerValues.isEmpty() == true) {
					headerValues = WebHelperUtil.getValueFromHashMap(valuesSheet);
					// log.info("headerValues : " + headerValues);//Sel4
				}
				String TestCaseID_Values_Str = WebHelperUtil.getCellData("TestCaseID", valuesSheet, rowIndex,
						headerValues);
				if (TestCaseID_Values_Str != "") {

					testCaseID_Values = currentRow.getCell(headerValues.get("TestCaseID"));
					cycleDate_Values = currentRow.getCell(headerValues.get("CycleDate"));
					try {
						executeFlag_Values = currentRow.getCell(headerValues.get("ExecuteFlag"));
					} catch (Exception ex) {
						log.error("ExecuteFlag is null. Reading value from Execute_Status");
						executeFlag_Values = currentRow.getCell(headerValues.get("Execute_Status"));
					}
					transactionType = currentRow.getCell(headerValues.get("TransactionType"));
					SimpleDateFormat cdFormat = new SimpleDateFormat("dd-MMM-yyyy");
					DateFormat cycleDateFormat = new SimpleDateFormat("MM/dd/yyyy");
					String CycleDate_Valuessheet = null;
					String CycleDate_MainController = null;
					if (!transactionType.toString().equalsIgnoreCase("Login")
							&& !transactionType.toString().equalsIgnoreCase("ChangeBusinessDate")) {
						if (cycleDate_Values.toString().contains("-")) {
							Date CycleDate_Values = cdFormat.parse(cycleDate_Values.toString());
							CycleDate_Valuessheet = cycleDateFormat.format(CycleDate_Values);
						} else if (cycleDate_Values.toString().equalsIgnoreCase("NA")
								|| cycleDate_Values.toString().equalsIgnoreCase("")) {
							CycleDate_Valuessheet = "NA";
						} else {
							CycleDate_Valuessheet = cycleDate_Values.toString();
						}

						if (controller.cycleDateCellValue.toString().contains("-")) {
							Date CycleDate_Main_Controller = cycleDateFormat
									.parse(controller.cycleDateCellValue.toString());
							CycleDate_MainController = cycleDateFormat.format(CycleDate_Main_Controller);
						} else if (controller.cycleDateCellValue.toString().equalsIgnoreCase("NA"))

						{
							CycleDate_MainController = "NA";
						}

						else {
							CycleDate_MainController = controller.cycleDateCellValue.toString();
						}

						if ((testCaseID_Values.toString().equalsIgnoreCase(controller.controllerTestCaseID.toString()))
								&& ((CycleDate_Valuessheet.equals(CycleDate_MainController)))
								&& (!executeFlag_Values.toString().equalsIgnoreCase("Y")) && (transactionType.toString()
										.equalsIgnoreCase(controller.controllerTransactionType.toString()))) {
							executeFlagcount = executeFlagcount + 1;
							log.info("executeFlagcount : " + executeFlagcount);
						}

						if ((!testCaseID_Values.toString().equalsIgnoreCase(controller.controllerTestCaseID.toString()))
								|| (!(CycleDate_Valuessheet.equals(CycleDate_MainController)))
								|| (!executeFlag_Values.toString().equalsIgnoreCase("Y"))
								|| (!transactionType.toString()
										.equalsIgnoreCase(controller.controllerTransactionType.toString()))) {
							if (rowIndex == rowCount) {
								//
							}
							continue;
						}
					}
					if (firstRowFound == false) {
						expectedRow = valuesSheet.getRow(rowIndex);
						valuesheetrows.add(rowIndex);
						singlerowindex = rowIndex;
					} else {
						valuesheetrows.add(rowIndex);
					}
					// log.info("row found :" + rowIndex);//Sel4
					firstRowFound = true;

				}
			}
			if (executeFlagcount == 1) {
				webDriver.getReport().setMessage("ExecuteFLag is N in value sheet");
				if (Config.recovery_scenario.equalsIgnoreCase("TRUE")) {
					tempRecovery_Scenario = Config.recovery_scenario;
					Config.recovery_scenario = "FALSE";
				}
				controller.pauseFun(webDriver.getReport().getMessage());
				webDriver.getReport().setMessage(webDriver.getReport().getMessage());
				ExcelUtility.writeReport(webDriver.getReport());
				executeFlagcount = 0;
				if (tempRecovery_Scenario == "TRUE") {
					Config.recovery_scenario = tempRecovery_Scenario;
				}
			} else

			if (valuesheetrows.isEmpty()) {
				webDriver.getReport().setMessage("No data found in value sheet");
				log.info("Test case will be marked 'N' as there is No data found in the" + " " + transactionType + " "
						+ "Transaction sheet");// 24/06/23
				controller.pauseFun(webDriver.getReport().getMessage());
				webDriver.getReport().setMessage(webDriver.getReport().getMessage());
				ExcelUtility.writeReport(webDriver.getReport());
			} else {
				TIvaluesheetrows = valuesheetrows;
				WebHelperUtil.GetCellInfo(structurePath, FilePath, expectedRow, singlerowindex, rowCount,
						TransactionType, TestCaseID, operationType, valuesheetrows);
			}
		}

		catch (IOException we) {
			
			
			log.error(we.getMessage(), we);
			webDriver.getReport().setMessage(we.getMessage());
		} catch (Exception e) {
			String errorMessage = e.getMessage();
			//String firstLine = errorMessage.split("\\r?\\n")[0];
			log.error(errorMessage);
			webDriver.getReport().setMessage(errorMessage);
		}
		return expectedRow;
	}

	//For PAS, Claims, DM --- Linear Approach - Row-wise execution
	// Reads the Values sheet from Input Excel and returns the row
	public static Row GetDataFromValues(String FilePath, String TestCaseID, String TransactionType)
			throws IOException, InterruptedException, Exception {
		Row expectedRow = null;
		Sheet valuesSheet = GetSheet(FilePath, "Values");

		int rowCount = valuesSheet.getLastRowNum() + 1;

		int endRow = 0;
		if (TransactionMapping.isTransactionStructureCommon) {
			firstRow = 1;
			endRow = 1;
		} else {
			endRow = getRowCount(valuesSheet, "");
		}
		if (endRow == 0) {
			controller.pauseExecution = true;
		}

		for (int rowIndex = firstRow; rowIndex < firstRow + endRow && !controller.pauseExecution; rowIndex++) {
			expectedRow = valuesSheet.getRow(rowIndex);
			if (ITAFWebDriver.isDMApplication() || (ITAFWebDriver.isSuiteApplication()
					&& MainControllerSuite.transactionSilo.equalsIgnoreCase("DM")))
				WebHelperDM.GetCellInfo(FilePath, expectedRow, rowIndex, rowCount);
			else if (ITAFWebDriver.isClaimsApplication() || (ITAFWebDriver.isSuiteApplication()
					&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims")))
				WebHelperClaims.GetCellInfo(FilePath, expectedRow, rowIndex, rowCount);
			else if (ITAFWebDriver.isLNAApplication())
				WebHelperLNA.GetCellInfo(FilePath, expectedRow, rowIndex, rowCount);
			else if (ITAFWebDriver.isPLATApplication())
				WebHelperDS.GetCellInfo(FilePath, expectedRow, rowIndex, rowCount);
			else if (ITAFWebDriver.isCSApplication())
				WebHelperCS.GetCellInfo(FilePath, expectedRow, rowIndex, rowCount);
			/*
			 * else if (ITAFWebDriver.isSuiteApplication())//Mandar - TO Test And Delete
			 * WebHelperClaims.GetCellInfo(FilePath, expectedRow, rowIndex, rowCount);
			 */
			/*
			 * Reporter report =
			 * TransactionMapping.TransactionInputData(controllerTestCaseID,
			 * controllerTransactionType, filePath); else if
			 * ((TransactionMapping.TransactionInputData(controllerTestCaseID)
			 * WebHelperCS.GetCellInfo(FilePath, expectedRow, rowIndex, rowCount);
			 */
			
			else if (ITAFWebDriver.isSuiteApplication())
				WebHelperSuite.GetCellInfo(FilePath, expectedRow, rowIndex, rowCount);
			else
				WebHelperPAS.GetCellInfo(FilePath, expectedRow, rowIndex, rowCount);

		}
		return expectedRow;
	}

	public static int getRowCount(Sheet valSheet, String cTransactionType, String operationtype) throws IOException {
		Cell testCaseID = null;
		Cell transactionType = null;
		int loopRowCount = 0;

		firstRow = 1;
		Boolean isFirstFound = false;
		int rowCount = valSheet.getLastRowNum() + 1;
		for (int rowIndex = 1; rowIndex < rowCount; rowIndex++) {
			Row row = valSheet.getRow(rowIndex);
			testCaseID = row.getCell(0);
			String tCase = null;

			if (testCaseID == null) {
				tCase = "";
			} else {
				tCase = testCaseID.toString();
			}
			transactionType = row.getCell(1);
			String tType = null;
			if (transactionType == null) {
				break;
			}
			tType = transactionType.toString();
			if (transactionType.toString().equalsIgnoreCase("Login")
					|| transactionType.toString().equalsIgnoreCase("ChangeBusinessDate")) {
				if (firstRow == 1 && !isFirstFound) {
					firstRow = rowIndex;
					isFirstFound = true;
				}
				loopRowCount++;
			} else {
				log.info("Value TestCase ID is : " + tCase);
				log.info("MainController TestCase ID is :" + controller.controllerTestCaseID.toString());
				log.info("Value Transaction is :" + tType);
				log.info("MainController Transaction is :" + controller.controllerTransactionType.toString());
				if ((tCase.equalsIgnoreCase(controller.controllerTestCaseID.toString())
						&& tType.equals(controller.controllerTransactionType.toString()))
						|| (tType.equalsIgnoreCase(cTransactionType.toString()) && operationtype == "Capture")) {
					if (firstRow == 1 && !isFirstFound) {
						firstRow = rowIndex;
						isFirstFound = true;
					}
					loopRowCount++;
				} else if ((!tCase.equalsIgnoreCase(controller.controllerTestCaseID.toString())
						|| !tType.equalsIgnoreCase(controller.controllerTransactionType.toString()))
						&& (!isFirstFound && rowIndex == rowCount - 1)) {
					controller.pauseFun("TestCaseID Or Transaction Didn't Match " + controller.controllerTestCaseID
							+ " " + controller.controllerTransactionType);
					if (ITAFWebDriver.isClaimsApplication())
						ExcelUtility.writeReportPAS(webDriver.getReport());
					else
						ExcelUtility.writeReport(webDriver.getReport());
					break;

				}
			}
		}
		return loopRowCount;
	}

	public static int getRowCount(Sheet valSheet, String purpose) throws IOException {
		Cell testCaseID = null;
		Cell transactionType = null;

		HashMap<String, Integer> headerValues = new HashMap<>();
		headerValues = WebHelperUtil.getValueFromHashMap(valSheet);

		int loopRowCount = 0;
		int temp = 0;
		firstRow = 1;
		Boolean isFirstFound = false;
		Boolean isDefaultFound = false;
		int rowCount = valSheet.getLastRowNum() + 1;

		for (int rowIndex = 1; rowIndex < rowCount; rowIndex++) {
			Row row = valSheet.getRow(rowIndex);

			if (row == null) {
				continue;
			}
			if (ITAFWebDriver.isClaimsApplication() || (ITAFWebDriver.isSuiteApplication()//Added suite condition - Mandar 12/10/2024
					&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims")))
				testCaseID = row.getCell(headerValues.get("TestCaseID"));
			else
				testCaseID = row.getCell(0);

			if (testCaseID == null || StringUtils.equalsIgnoreCase(testCaseID.toString().trim(), "")) {
				continue;
			}

			if (ITAFWebDriver.isClaimsApplication() || (ITAFWebDriver.isSuiteApplication()
					&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims")))
				transactionType = row.getCell(headerValues.get("TransactionType"));
			else
				transactionType = row.getCell(1);

			// bhaskar
			String tCase = testCaseID.toString();
			String tType = transactionType.toString();
			if (tCase.toString().equalsIgnoreCase("DEFAULT") && tType.toString().equalsIgnoreCase("Login")) {
				isDefaultFound = true;
				temp = rowIndex;
			}
			if (tCase.equalsIgnoreCase(controller.controllerTestCaseID.toString())
					&& tType.equalsIgnoreCase(controller.controllerTransactionType.toString())) {
				if (firstRow == 1 && !isFirstFound) {
					firstRow = rowIndex;
					isFirstFound = true;
				}
				loopRowCount++;

			} else if ((StringUtils.equalsIgnoreCase(purpose, "Verification"))
					&& (ITAFWebDriver.isBillingApplication() || (ITAFWebDriver.isSuiteApplication()
							&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Billing")))) {

				if (firstRow == 1 && !isFirstFound) {
					firstRow = rowIndex;
					isFirstFound = true;
				}
				loopRowCount++;

			} else if ((!tCase.equalsIgnoreCase(controller.controllerTestCaseID.toString())
					|| !tType.equalsIgnoreCase(controller.controllerTransactionType.toString()))
					&& (!isFirstFound && rowIndex == rowCount - 1)) {
				if (isDefaultFound == true) {
					firstRow = temp;
					isFirstFound = true;
					loopRowCount++;
				} else {
					if ((StringUtils.equalsIgnoreCase(purpose, "Verification"))) {
						loopRowCount = 0;
					} else {
						controller.pauseFun("TestCaseID Or Transaction Didn't Match " + controller.controllerTestCaseID
								+ " " + controller.controllerTransactionType);
						// ExcelUtility.writeReportPAS(webDriver.getReport());//For DM team - 17/04/2024
						break;
					}
				}
			}
		}
		return loopRowCount;

	}

	public static Sheet GetSheet(String FilePath, String SheetName) throws IOException
	{
		Workbook workBook = null;
		Sheet workSheet = null;
		
		try {
			log.info("FilePath is : " + FilePath + " and SheetName is : " + SheetName);// Sel4
			InputStream myXls = new FileInputStream(FilePath);
			// ZipSecureFile.setMinInflateRatio(-1.0d);
			if (FilePath.endsWith(".xls")) {
				workBook = new HSSFWorkbook(myXls);
			} else if (FilePath.endsWith(".xlsx") | FilePath.endsWith(".xlsm")) {
				workBook = new XSSFWorkbook(myXls);
			}
			workSheet = workBook.getSheet(SheetName);
		} catch (IOException e) {
			//log.error(e.getMessage(), e);
			String errorMessage = e.getMessage();
			String firstLine = errorMessage.split("\\r?\\n")[0];//Mandar
			log.error("File Not Found - " + firstLine);
			
			controller.pauseFun("Failed to access  file " + FilePath + " Sheet:  " + SheetName);
			log.error("Failed to access  file " + FilePath + "Sheet:  " + SheetName + "<-|-> LocalizeMessage "
					+ e.getLocalizedMessage() + " <-|-> Message " + e.getMessage() + " <-|-> Cause " + e.getCause());
			throw new IOException("Failed to access  file :" + FilePath + "Sheet:  " + SheetName
					+ "  <-|-> LocalizeMessage " + e.getLocalizedMessage() + " <-|-> Message" + e.getMessage()
					+ " <-|-> Cause " + e.getCause());
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			log.error("Failed to access  file " + FilePath + "Sheet:  " + SheetName + "<-|-> LocalizeMessage "
					+ e.getLocalizedMessage() + " <-|-> Message " + e.getMessage() + " <-|-> Cause " + e.getCause());
			controller.pauseFun("Failed to access  file " + FilePath + "Sheet:  " + SheetName);
			return null;
		} finally {
			workBook.close();
		}
		return workSheet;
	}

	// Writes SummaryResults
	public static void writeReport(Reporter report) throws IOException {
		if (ITAFWebDriver.isSuiteApplication()) {
			report.setTestDescription(controller.testDescription);
			writeReportSuite(report);
		} else {
			PrintStream print = null;
			try {
				// log.info(report.strMessage);
				report.setReport(report);
				// String frmDate = report.getFromDate();//Sel4
				File file = new File(Config.resultOutput);
				report = report.getReport();

				String lasttoDate = Config.dtFormat.format(new Date());
				report.setToDate(lasttoDate);

				// TM:19/01/2015-Changes made to remove ==null
				if (StringUtils.isBlank(report.getMessage()))
					report.setMessage("");

				// TM:19/01/2015-Changes made to remove ==null
				if (StringUtils.isBlank(report.getTestDescription()))
					report.setTestDescription("");

				if (report.getStatus() == "FAIL") {
					if (report.getTrasactionType().equalsIgnoreCase("RunBatch_JBEAM")) {
						System.out.println("Do not go in Recovery");
					} else if (report.getTrasactionType().startsWith("Verify")) {
						System.out.println("Do not go in Recovery");
					} else if (report.getTrasactionType().endsWith("Verify")) { // Vinay B- For verification transaction
						System.out.println("Do not go in Recovery");
						failedcount = 0;
					} else {
						failedcount = 0;
					}

				}

				if (report.getTrasactionType().equalsIgnoreCase("ChangeBusinessdate")) {
					report.setTestcaseId("Common");
				}
				if (StringUtils.isBlank(report.getGroupName()))
					report.setGroupName("");

				if (WebService.isNoResponseFileTrue()) {
					report.setMessage(WebHelper.description);
				}

				if (WebHelper.faultstring == true) {
					if (WebService.isNodeAvailable()) {
						report.setMessage("<" + WebService.getreportNodeValue() + "> " + WebHelper.description);
					} else {
						report.setMessage(WebHelper.description);
					}
				}

				// if(MainController.controllerTransactionType.toString().startsWith("WebService"))
				if (report.getTrasactionType().startsWith("WebService")) {
					report.setScreenShot("");
				}

				if (WebHelper.success == true) {
					report.setMessage(WebHelper.description);
				}

				if (file.exists() == false) {
					print = new PrintStream(file);
				}
				String timeElapsed = calTimeElapsed(report.getFromDate(), report.getToDate());
				int usedRows = WebHelperUtil.count(file);
				if (usedRows == 0) {
					if (ITAFWebDriver.isBillingApplication()) {
						print.print(
								"Iteration,TestCaseID,CycleDate,TransactionType,VerificationType,TestCaseDesription,StartDate,EndDate,ElapsedTime,Status,ExecutionDescription,Screenshot");
					}

					else {
						print.print(
								"Iteration,TestCaseID,CycleDate,TransactionType,TestCaseDesription,StartDate,EndDate,ElapsedTime,Status,ExecutionDescription,Screenshot");
					}

					print.println();

				}
				usedRows = WebHelperUtil.count(file);
				print = new PrintStream(new FileOutputStream(file, true));

				// it will remove millisecond as macro does not support it.
				if (MainControllerBilling.frmDateNew == null)
					MainControllerBilling.frmDateNew = new Date();

				report.setFromDate(Config.dtFormat.format(MainControllerBilling.frmDateNew));
				String fromDate = report.getFromDate().substring(0, (report.getFromDate().lastIndexOf(":")));
				String toDate = report.getToDate().substring(0, (report.getToDate().lastIndexOf(":")));

				// print in summary result with time elapsed.
				if (ITAFWebDriver.isBillingApplication()) {

					if (report.getVerificationType() == null) {
						report.setVerificationType("");
					}
				}

				if (ITAFWebDriver.isBillingApplication()) {// report.getMessage()
					String escapedMessage = report.getMessage().replace(",", ";"); // Replace commas with semicolons to
																					// avoid splitting
					print.print(myChar + Config.cycleNumber + myChar + "," + myChar + report.getTestcaseId() + myChar
							+ "," + myChar + report.getCycleDate() + myChar + "," + myChar + report.getTrasactionType()
							+ myChar + "," + myChar + report.getVerificationType() + myChar + "," + myChar
							+ report.getTestDescription() + myChar + "," + myChar + fromDate + myChar + "," + myChar
							+ toDate + myChar + "," + myChar + timeElapsed + myChar + "," + myChar + report.getStatus()
							+ myChar + "," + myChar + escapedMessage + myChar + "," + myChar + report.getScreenShot()
							+ myChar);

				} else {
					String escapedMessage = report.getMessage().replace(",", ";");
					print.print(myChar + Config.cycleNumber + myChar + "," + myChar + report.getTestcaseId() + myChar
							+ "," + myChar + report.getCycleDate() + myChar + "," + myChar + report.getTrasactionType()
							+ myChar + "," + myChar + report.getTestDescription() + myChar + "," + myChar + fromDate
							+ myChar + "," + myChar + toDate + myChar + "," + myChar + timeElapsed + myChar + ","
							+ myChar + report.getStatus() + myChar + "," + myChar + escapedMessage + myChar + ","
							+ myChar + report.getScreenShot() + myChar);
				}
				print.println();
				if (report.getTrasactionType().equalsIgnoreCase("Login")
						&& report.getStatus().equalsIgnoreCase("Fail")) {
					controller.pExecution = true;
				}

				if (report.getTrasactionType().equalsIgnoreCase("ChangeBusinessdate")
						&& report.getStatus().equalsIgnoreCase("Fail")) {
					controller.pExecution = true;
				}

				if (report.getStatus().equalsIgnoreCase("Fail")) {
				}
			}

			catch (IOException ie) {
				log.error(ie.getMessage(), ie);
				controller.pauseFun(ie.getMessage());
			} finally {
				webDriver.getReport().setScreenShot("");
			}
			MainControllerBilling.frmDateNew = new Date();
		}
	}

	// Writes SummaryResults
	public static void writeReportPAS(Reporter report) throws IOException {
		if (ITAFWebDriver.isSuiteApplication()) {
			report.setTestDescription(controller.testDescription);
			writeReportSuite(report);
		} else {
			PrintStream print = null;
			try {

				File file = new File(Config.resultOutput);

				if (file.exists() == true && StringUtils.equalsIgnoreCase(Config.appendResultOutput, "false")
						&& DResult == 1) {
					file.delete();
				}

				report.setReport(report);
				report = report.getReport();
				// Quote ID for PAS report requirement
				if (controller.controllerQuoteId != null) {
					report.setStrQuoteId(controller.controllerQuoteId.toString());
				} else {
					report.setStrQuoteId("");
				}
				// TM:19/01/2015-Changes made to remove ==null
				if (StringUtils.isBlank(report.getMessage()))
					report.setMessage("");

				if (StringUtils.contains(report.getMessage(), "\"")) {
					report.setMessage(report.getMessage().replaceAll("\"", "'"));
				}
				if (StringUtils.contains(report.getMessage(), ",")) {
					report.setMessage("\"\"" + report.getMessage() + "\"\"");
				}

				// TM:19/01/2015-Changes made to remove ==null
				if (StringUtils.isBlank(report.getTestDescription()))
					report.setTestDescription("");

				// TM:19/01/2015-Added for GroupName Blank
				if (StringUtils.isBlank(report.getGroupName()))
					report.setGroupName("");

				Date finalToDate = new Date();// Added by PAS COE -26/09/2019
				webDriver.getReport().setToDate(Config.dtFormat.format(finalToDate));

				String timeElapsed = calTimeElapsed(report.getFromDate(), report.getToDate());// Added by PAS COE -
				// 25/09/2019
				print = new PrintStream(new FileOutputStream(file, true));

				int usedRows = WebHelperUtil.count(file);
				if (usedRows == 0) {
					if (ITAFWebDriver.isClaimsApplication())
						print.print(
								"GroupName,Iteration,TestCaseID,TransactionType,TestCaseDesription,StartDate,EndDate,ExecutionTime(Seconds),Status,ExecutionDescription,Screenshot");
					else
						print.print(
								"GroupName,Iteration,TestCaseID,TransactionType,TestCaseDesription,StartDate,EndDate,ExecutionTime(Seconds),Status,Description,Screenshot,Quote Number");
					print.println();
				}
				usedRows = WebHelperUtil.count(file);
				print = new PrintStream(new FileOutputStream(file, true));
				String fromDate = report.getFromDate().substring(0, (report.getFromDate().lastIndexOf(":")));
				String toDate = report.getToDate().substring(0, (report.getToDate().lastIndexOf(":")));

				// 19/09/2022 - Below given condition for Claims Report added controllerSC_No
				// instead of Config.cycleNumber
				if (ITAFWebDriver.isClaimsApplication()) {
					print.print(myChar + report.getGroupName() + myChar + "," + myChar + controller.controllerSC_No
							+ myChar + "," + myChar + report.getTestcaseId() + myChar + "," + myChar
							+ report.getTrasactionType() + myChar + "," + myChar + report.getTestDescription() + myChar
							+ "," + myChar + fromDate + myChar + "," + myChar + toDate + myChar + "," + myChar
							+ timeElapsed + myChar + "," + myChar + report.getStatus() + myChar + "," + myChar
							+ report.getMessage() + myChar + "," + myChar + report.getScreenShot() + myChar + ","
							+ myChar + report.getStrQuoteId() + myChar);
				}
				// ***

				else {

					print.print(myChar + report.getGroupName() + myChar + "," + myChar + Config.cycleNumber + myChar
							+ "," + myChar + report.getTestcaseId() + myChar + "," + myChar + report.getTrasactionType()
							+ myChar + "," + myChar + report.getTestDescription() + myChar + "," + myChar + fromDate
							+ myChar + "," + myChar + toDate + myChar + "," + myChar + timeElapsed + myChar + ","
							+ myChar + report.getStatus() + myChar + "," + myChar + report.getMessage() + myChar + ","
							+ myChar + report.getScreenShot() + myChar + "," + myChar + report.getStrQuoteId()
							+ myChar);
				}

				print.println();
				DResult++;
			} catch (IOException ie) {
				log.error(ie.getMessage(), ie);
				controller.pauseFun(ie.getMessage());
			} finally {
				webDriver.getReport().setScreenShot("");
			}
		}
	}

	// Compares the Actual and Expected Sheets cell-wise
	// ***For PAS - DM***
	public static Reporter CompareExcel(Sheet actualSheet, Sheet expectedSheet, List<String> columns,
			List<String> columnsData, String testCaseID, String transactionType, String operationtype, String cycleDate)
			throws IOException {

		List<String> status = new ArrayList<String>();
		List<String> rowStatus = new ArrayList<String>();
		List<String> actualValue = new ArrayList<String>();
		List<List<String>> actualRows = new ArrayList<List<String>>();
		List<Integer> passCounts = new ArrayList<Integer>();
		List<Integer> failCounts = new ArrayList<Integer>();
		String expectedHeaderValue = null;
		String actualHeaderValue = null;

		boolean isrowFound = false;
		int expSheetRowCount = ExcelUtility.getRowCount(expectedSheet, transactionType, operationtype); // expectedSheet.getPhysicalNumberOfRows();
		Reporter report = new Reporter();
		report.setReport(report);
		int passCount = 0;
		int failCount = 0;
		int colCount = 0;

		for (int rowIndex = firstRow; rowIndex < firstRow + expSheetRowCount; rowIndex++) {
			passCount = 0;
			failCount = 0;
			int currentRow = ++WebVerification.currentRowIndex;

			Row actualRow = actualSheet.getRow(currentRow);
			Row expectedRow = expectedSheet.getRow(rowIndex);
			Row actualRowHeader = actualSheet.getRow(0);
			Row expectedRowHeader = expectedSheet.getRow(0);
			if (isClaimsApplication || (ITAFWebDriver.isSuiteApplication()
					&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims"))) {
				try {
					for (int i = 0; i <= 3; i++) {
						System.out.println(actualRow.getCell(i).toString());
						System.out.println(expectedRow.getCell(i).toString());
						if (actualRow.getCell(i).toString() == null) {
							break;
						}
						if (expectedRow.getCell(i).toString() == null) {
							break;
						}

					}
				} catch (Exception e) {
					log.error(e.getMessage(), e);
					break;
				}
			}
			if (actualRow.getCell(0).toString().equals(expectedRow.getCell(0).toString())
					&& actualRow.getCell(1).toString().equals(expectedRow.getCell(1).toString())
					&& actualRow.getCell(2).toString().equals(expectedRow.getCell(2).toString())
					&& ((!(isClaimsApplication) && !(ITAFWebDriver.isSuiteApplication()
							&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims")))
							|| actualRow.getCell(3).toString().equals(expectedRow.getCell(3).toString()))) {

				isrowFound = true;
				actualValue = new ArrayList<String>();
				if (actualRow == null || expectedRow == null) {
					break;
				}
				colCount = expectedRow.getPhysicalNumberOfCells();
				int startColIndex = 3;
				if ((ITAFWebDriver.isSuiteApplication()
						&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims")) || isClaimsApplication)
					startColIndex = 5;

				for (int columnIndex = startColIndex; columnIndex < colCount; columnIndex++) {
					Cell actualCell = actualRow.getCell(columnIndex);
					DataFormatter fmt = new DataFormatter();
					Cell expectedCell = expectedRow.getCell(columnIndex);
					Cell actualCellHeader = actualRowHeader.getCell(columnIndex);
					DataFormatter fmt1 = new DataFormatter();
					Cell expectedCellHeader = expectedRowHeader.getCell(columnIndex);

					if (actualCellHeader != null || expectedCellHeader != null) {
						expectedHeaderValue = fmt1.formatCellValue(expectedCellHeader);
						actualHeaderValue = fmt1.formatCellValue(actualCellHeader);
						log.info("Dashboard value of actual sheet:" + expectedHeaderValue);
						log.info("Dashboard value of expected sheet:" + actualHeaderValue);
						if (expectedHeaderValue == actualHeaderValue) {
							log.info("Column Headers of actual and expected sheet are :" + expectedHeaderValue);
						}
					}
					// TM: Following 'if' is replacement of the above
					if (actualCell != null || expectedCell != null) {
						String expectedValue = fmt.formatCellValue(expectedCell);
						String stractualValue = fmt.formatCellValue(actualCell);
						log.info("Dashboard value of actual sheet:" + stractualValue);
						log.info("Dashboard value of expected sheet:" + expectedValue);
						if (!actualCell.toString().equalsIgnoreCase(expectedValue)) {
							report.setStatus("FAIL");
							report.setStatus(report.getStatus());
							failCount += 1;
							report.setActualValue("FAIL |" + expectedValue + "|" + actualCell.toString());
							report.setMessage("Values Not Matched");
							ExcelUtility.WriteToCompareDetailResults(testCaseID, transactionType, columns, actualRows,
									expSheetRowCount, colCount, report, expectedValue, stractualValue,
									actualHeaderValue, operationtype, cycleDate);
						} else {
							passCount += 1;
							report.setStatus("PASS");
							report.setStatus(report.getStatus());
							report.setActualValue(actualCell.toString());
							log.info(actualCell.toString());
							report.setMessage("Values Matched");
							ExcelUtility.WriteToCompareDetailResults(testCaseID, transactionType, columns, actualRows,
									expSheetRowCount, colCount, report, expectedValue, stractualValue,
									actualHeaderValue, operationtype, cycleDate);
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
				rowStatus.add(report.getStatus());
				passCounts.add(passCount);
				failCounts.add(failCount);
				actualRows.add(actualValue);
				report.setReport(report);
			} else if (isrowFound == false) {
				continue;
			}
		}
		if (rowStatus.contains("FAIL")) {
			report.setStatus("FAIL");
			if (isClaimsApplication || (ITAFWebDriver.isSuiteApplication()
					&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims"))) {
				report.setMessage("Values Not Matched");
			}
		} else {
			if (isClaimsApplication || (ITAFWebDriver.isSuiteApplication()
					&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims"))) {
				report.setStatus("PASS");
			}
		}
		if (!isClaimsApplication && !(ITAFWebDriver.isSuiteApplication()
				&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims"))) {
			cycleDate = null;
		}
		ExcelUtility.WriteToDetailResults(testCaseID, transactionType, columns, actualRows, passCounts, failCounts,
				expSheetRowCount, colCount, report, rowStatus, operationtype, cycleDate);
		passCounts.clear();
		failCounts.clear();
		return report;
	}

	// Compares the Actual and Expected Sheets cell-wise.

	// ***For Billing***
	public static Reporter CompareExcelBilling(Sheet actualSheet, Sheet expectedSheet, List<String> columns,
			List<String> columnsData, String testCaseID, String transactionType, String cycleDate) throws IOException {
		Reporter report = new Reporter();// Cycle Date - Added by Mandar
		report.setReport(report);

		try {
			boolean isrowFound = false;
			int expSheetRowCount = getRowCount(expectedSheet, "Verification"); // expectedSheet.getPhysicalNumberOfRows();
			int actualSheetRowsForComparison = actualSheet.getLastRowNum() - WebVerification.currentRowIndex;

			if (expSheetRowCount <= 0) {
				report.setStatus("FAIL");
				report.setMessage("No data resent in Expected sheet for current transaction| Actual Sheet created");
				return report;
			}
			if (expSheetRowCount != actualSheetRowsForComparison) {
				report.setStatus("FAIL");
				report.setMessage("Mismatch in Actual and Expected row count.\nExpected Row count = " + expSheetRowCount
						+ " | Actual Row count = " + actualSheetRowsForComparison + "\nCheck at - "
						+ Config.expectedValuesValuesPath);
				return report;
			}

			int passCount = 0;
			int failCount = 0;
			int colCount = 0;

			List<String> status = new ArrayList<String>();
			List<String> rowStatus = new ArrayList<String>();
			List<List<String>> actualRows = new ArrayList<List<String>>();
			List<Integer> passCounts = new ArrayList<Integer>();
			List<Integer> failCounts = new ArrayList<Integer>();

			for (int rowIndex = firstRow; rowIndex < firstRow + expSheetRowCount; rowIndex++) {
				passCount = 0;
				failCount = 0;
				int currentRow = ++WebVerification.currentRowIndex;
				Row actualRow = actualSheet.getRow(currentRow);
				Row expectedRow = expectedSheet.getRow(rowIndex);
				List<String> actualValue = new ArrayList<String>();

				if (actualRow.getCell(0).toString().equals(expectedRow.getCell(0).toString())
						&& actualRow.getCell(1).toString().equals(expectedRow.getCell(1).toString())) {

					isrowFound = true;
					actualValue = new ArrayList<String>();
					if (actualRow == null || expectedRow == null) {
						break;
					}
					if (ITAFWebDriver.isDMApplication() || (ITAFWebDriver.isSuiteApplication()
							&& MainControllerSuite.transactionSilo.equalsIgnoreCase("DM"))) {// Minaakshi : 01-01-2020
						colCount = expectedRow.getLastCellNum();
					} else {
						colCount = expectedRow.getPhysicalNumberOfCells();
					}
					for (int columnIndex = 3; columnIndex < colCount; columnIndex++) {
						Cell actualCell = actualRow.getCell(columnIndex);
						DataFormatter fmt = new DataFormatter();
						Cell expectedCell = expectedRow.getCell(columnIndex);
						if (actualCell != null || expectedCell != null) {
							String expectedValue = fmt.formatCellValue(expectedCell);
							if (ITAFWebDriver.isDMApplication() || (ITAFWebDriver.isSuiteApplication()
									&& MainControllerSuite.transactionSilo.equalsIgnoreCase("DM"))) {// Minaakshi :
																										// 01-01-2020
								if (expectedValue.contains("|")) {
									expectedValue = DMProduct.getDynamicExpectedValue(expectedValue,
											actualCell.toString());
								}
								if (expectedValue.contains("TEXT")) {
									expectedValue = DMProduct.getExpectedValueByFormula(expectedValue,
											actualCell.toString(), expectedSheet, expectedCell);
								}
							}
							if (expectedValue.equalsIgnoreCase("IGNORE")) {
								passCount += 1;
								report.setStatus("PASS");
								report.setStatus(report.getStatus());
								report.setActualValue(expectedValue + "|" + actualCell.toString());

							} else if (!actualCell.toString().equalsIgnoreCase(expectedValue)) {
								report.setStatus("FAIL");
								report.setStatus(report.getStatus());
								failCount += 1;
								report.setActualValue("FAIL |" + expectedValue + "|" + actualCell.toString());
							} else {
								passCount += 1;
								report.setStatus("PASS");
								report.setStatus(report.getStatus());
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
					rowStatus.add(report.getStatus());
					passCounts.add(passCount);
					failCounts.add(failCount);
					actualRows.add(actualValue);
					report.setReport(report);
				} else if (isrowFound == false) {
					continue;
				}
			}
			if (rowStatus.contains("FAIL")) {
				report.setStatus("FAIL");
			}
			WriteToDetailResults(testCaseID, transactionType, cycleDate, columns, actualRows, passCounts, failCounts,
					expSheetRowCount, colCount, report, rowStatus);// Cycle - Date Mandar
			passCounts.clear();
			failCounts.clear();
			return report;
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			report.setStatus("FAIL");
			report.setMessage("Issue in Excel Comparison." + e.getMessage());
			return report;

		}
	}

	// ***For PAS - DM NEW***
	/*
	 * public static Reporter CompareExcel(Sheet actualSheet, Sheet expectedSheet,
	 * List<String> columns, List<String> columnsData, String testCaseID, String
	 * transactionType) throws IOException { Reporter report = new Reporter();
	 * report.setReport(report);
	 * 
	 * try { boolean isrowFound = false; int expSheetRowCount =
	 * getRowCount(expectedSheet, "Verification"); int actualSheetRowsForComparison
	 * = actualSheet.getLastRowNum() - WebVerification.currentRowIndex;
	 * 
	 * if (expSheetRowCount <= 0) { report.setStatus("FAIL"); report.
	 * setMessage("No data resent in Expected sheet for current transaction| Actual Sheet created"
	 * ); return report; } if (expSheetRowCount != actualSheetRowsForComparison) {
	 * report.setStatus("FAIL"); report.
	 * setMessage("Mismatch in Actual and Expected row count.\nExpected Row count = "
	 * + expSheetRowCount + " | Actual Row count = " + actualSheetRowsForComparison
	 * + "\nCheck at - " + Config.expectedValuesValuesPath); return report; }
	 * 
	 * int passCount = 0; int failCount = 0; int colCount = 0;
	 * 
	 * List<String> status = new ArrayList<String>(); List<String> rowStatus = new
	 * ArrayList<String>(); List<List<String>> actualRows = new
	 * ArrayList<List<String>>(); List<Integer> passCounts = new
	 * ArrayList<Integer>(); List<Integer> failCounts = new ArrayList<Integer>();
	 * 
	 * for (int rowIndex = firstRow; rowIndex < firstRow + expSheetRowCount;
	 * rowIndex++) { passCount = 0; failCount = 0; int currentRow =
	 * ++WebVerification.currentRowIndex; Row actualRow =
	 * actualSheet.getRow(currentRow); Row expectedRow =
	 * expectedSheet.getRow(rowIndex); List<String> actualValue = new
	 * ArrayList<String>();
	 * 
	 * if (actualRow.getCell(0).toString().equals(expectedRow.getCell(0).toString())
	 * && actualRow.getCell(1).toString().equals(expectedRow.getCell(1).toString()))
	 * {
	 * 
	 * isrowFound = true; actualValue = new ArrayList<String>(); if (actualRow ==
	 * null || expectedRow == null) { break; } if (ITAFWebDriver.isDMApplication())
	 * { colCount = expectedRow.getLastCellNum(); } else { colCount =
	 * expectedRow.getPhysicalNumberOfCells(); } for (int columnIndex = 3;
	 * columnIndex < colCount; columnIndex++) { Cell actualCell =
	 * actualRow.getCell(columnIndex); DataFormatter fmt = new DataFormatter(); Cell
	 * expectedCell = expectedRow.getCell(columnIndex); if (actualCell != null ||
	 * expectedCell != null) { String expectedValue =
	 * fmt.formatCellValue(expectedCell); if (ITAFWebDriver.isDMApplication() ||
	 * (ITAFWebDriver.isSuiteApplication() &&
	 * MainControllerSuite.transactionSilo.equalsIgnoreCase("DM"))) {// Minaakshi :
	 * 01-01-2020 if (expectedValue.contains("|")) { expectedValue =
	 * DMProduct.getDynamicExpectedValue(expectedValue, actualCell.toString()); } if
	 * (expectedValue.contains("TEXT")) { expectedValue =
	 * DMProduct.getExpectedValueByFormula(expectedValue, actualCell.toString(),
	 * expectedSheet, expectedCell); } } if
	 * (expectedValue.equalsIgnoreCase("IGNORE")) { passCount += 1;
	 * report.setStatus("PASS"); report.setStatus(report.getStatus());
	 * report.setActualValue(expectedValue + "|" + actualCell.toString());
	 * 
	 * } else if (!actualCell.toString().equalsIgnoreCase(expectedValue)) {
	 * report.setStatus("FAIL"); report.setStatus(report.getStatus()); failCount +=
	 * 1; report.setActualValue("FAIL |" + expectedValue + "|" +
	 * actualCell.toString()); } else { passCount += 1; report.setStatus("PASS");
	 * report.setStatus(report.getStatus());
	 * report.setActualValue(actualCell.toString());
	 * System.out.println(actualCell.toString()); } status.add(report.getStatus());
	 * actualValue.add(report.getActualValue()); }
	 * 
	 * } if (status.contains("FAIL")) { report.setStatus("FAIL"); } else {
	 * report.setStatus("PASS"); } status.clear();
	 * rowStatus.add(report.getStatus()); passCounts.add(passCount);
	 * failCounts.add(failCount); actualRows.add(actualValue);
	 * report.setReport(report); } else if (isrowFound == false) { continue; } } if
	 * (rowStatus.contains("FAIL")) { report.setStatus("FAIL"); }
	 * WriteToDetailResults(testCaseID, transactionType, "", columns, actualRows,
	 * passCounts, failCounts, expSheetRowCount, colCount, report,
	 * rowStatus);//Passing Cycle date blank - Mandar passCounts.clear();
	 * failCounts.clear(); return report; } catch (Exception e) {
	 * log.error(e.getMessage(), e); report.setStatus("FAIL");
	 * report.setMessage("Issue in Excel Comparison." + e.getMessage()); return
	 * report;
	 * 
	 * } }
	 */

	/*
	 * public static Reporter compareExcel(Sheet actualSheet, Sheet expectedSheet,
	 * List<String> columns, List<String> columnsData, String testCaseID, String
	 * transactionType) throws IOException { Reporter report = new Reporter();
	 * report.setReport(report);
	 * 
	 * try { int expSheetRowCount = getRowCount(expectedSheet, "Verification"); int
	 * actualSheetLastRowNum = actualSheet.getLastRowNum();
	 * 
	 * if (expSheetRowCount <= 0) { report.setStatus("FAIL"); report.
	 * setMessage("No data resent in Expected sheet for current transaction| Actual Sheet created"
	 * ); return report; }
	 * 
	 * if (actualSheetLastRowNum < firstRow) { report.setStatus("FAIL");
	 * report.setMessage("No data found in Actual sheet for comparison."); return
	 * report; }
	 * 
	 * int colCount = columns.size();
	 * 
	 * int expectedRowIndex = findRowIndex(expectedSheet, testCaseID,
	 * transactionType); int actualRowIndex = actualSheetLastRowNum;
	 * 
	 * if (expectedRowIndex == -1) { report.setStatus("FAIL"); report.
	 * setMessage("No corresponding entry found in Expected sheet for testCaseID: "
	 * + testCaseID + " and transactionType: " + transactionType); return report; }
	 * 
	 * Row actualRow = actualSheet.getRow(actualRowIndex); Row expectedRow =
	 * expectedSheet.getRow(expectedRowIndex);
	 * 
	 * if (actualRow == null || expectedRow == null) { report.setStatus("FAIL");
	 * report.setMessage("Last rows in Actual or Expected sheets are null."); return
	 * report; }
	 * 
	 * List<String> actualValuesForRow = new ArrayList<>(); int passCount = 0; int
	 * failCount = 0;
	 * 
	 * for (int columnIndex = 0; columnIndex < colCount; columnIndex++) {
	 * 
	 * String columnName = columns.get(columnIndex); String actualValue = ""; String
	 * expectedValue = "";
	 * 
	 * // Skip adding testCaseID and transactionType columns to actualValuesForRow
	 * if (columnName.equalsIgnoreCase("TestCaseID") ||
	 * columnName.equalsIgnoreCase("TransactionType") ||
	 * columnName.equalsIgnoreCase("CurrentDate")) { continue; }
	 * 
	 * Cell actualCell = actualRow.getCell(columnIndex); Cell expectedCell =
	 * expectedRow.getCell(columnIndex);
	 * 
	 * if (actualCell != null) { actualValue = actualCell.toString(); }
	 * 
	 * if (expectedCell != null) { expectedValue = expectedCell.toString(); if
	 * (ITAFWebDriver.isDMApplication()) { if (expectedValue.contains("|")) {
	 * 
	 * expectedValue = DMProduct.getDynamicExpectedValue(expectedValue,
	 * actualCell.toString());
	 * 
	 * }if (expectedValue.contains("TEXT")) {// Minaakshi : // 04-01-2019
	 * expectedValue = DMProduct.getExpectedValueByFormula(expectedValue,
	 * actualCell.toString(), expectedSheet, expectedCell); } } }
	 * 
	 * if (expectedValue.equalsIgnoreCase("IGNORE")) { passCount++;
	 * actualValuesForRow.add("PASS | IGNORE"); } else if
	 * (actualValue.equalsIgnoreCase(expectedValue)) { passCount++;
	 * actualValuesForRow.add("PASS | " + actualValue);
	 * 
	 * 
	 * } else { failCount++; actualValuesForRow.add("FAIL | Expected: " +
	 * expectedValue + ", Actual: " + actualValue); } }
	 * 
	 * if (failCount > 0) { report.setStatus("FAIL"); } else {
	 * report.setStatus("PASS"); }
	 * 
	 * List<String> rowStatus = Collections.singletonList(report.getStatus());
	 * List<List<String>> actualRows =
	 * Collections.singletonList(actualValuesForRow); List<Integer> passCounts =
	 * Collections.singletonList(passCount); List<Integer> failCounts =
	 * Collections.singletonList(failCount);
	 * 
	 * WriteToDetailResults(testCaseID, transactionType, "", columns, actualRows,
	 * passCounts, failCounts, expSheetRowCount, colCount, report, rowStatus);
	 * 
	 * return report; } catch (Exception e) { log.error(e.getMessage(), e);
	 * report.setStatus("FAIL"); report.setMessage("Issue in Excel Comparison." +
	 * e.getMessage()); return report; } }
	 * 
	 * private static int findRowIndex(Sheet sheet, String testCaseID, String
	 * transactionType) { int rowCount = sheet.getLastRowNum(); for (int i =
	 * firstRow; i <= rowCount; i++) { Row row = sheet.getRow(i); if (row != null &&
	 * row.getCell(0) != null && row.getCell(1) != null &&
	 * row.getCell(0).toString().equals(testCaseID) &&
	 * row.getCell(1).toString().equals(transactionType)) { return i; } } return -1;
	 * // If not found }
	 */

	// ***For PAS - DM NEW***
	//***For PAS - DM NEW***
		public static Reporter compareExcel(Sheet actualSheet, Sheet expectedSheet, List<String> columns,
				List<String> columnsData, String testCaseID, String transactionType) throws IOException {
			Reporter report = new Reporter();
			report.setReport(report);

			try {
				boolean isrowFound = false;
				int expSheetRowCount = getRowCount(expectedSheet, "Verification"); // expectedSheet.getPhysicalNumberOfRows();
				int actualSheetRowsForComparison = actualSheet.getLastRowNum() - WebVerification.currentRowIndex;

				if (expSheetRowCount <= 0) {
					report.setStatus("FAIL");
					report.setMessage("No data resent in Expected sheet for current transaction| Actual Sheet created");
					return report;
				}
				if (expSheetRowCount != actualSheetRowsForComparison) {
					report.setStatus("FAIL");
					
					log.info("Mismatch in Actual and Expected row count.\nExpected Row count = " + expSheetRowCount
							+ " | Actual Row count = " + actualSheetRowsForComparison);
					report.setMessage("Mismatch in Actual and Expected row count.\nExpected Row count = " + expSheetRowCount
							+ " | Actual Row count = " + actualSheetRowsForComparison + "\nCheck at - "
							+ Config.expectedValuesValuesPath);
					
					
					return report;
				}

				int passCount = 0;
				int failCount = 0;
				int colCount = 0;

				List<String> status = new ArrayList<String>();
				List<String> rowStatus = new ArrayList<String>();
				List<List<String>> actualRows = new ArrayList<List<String>>();
				List<Integer> passCounts = new ArrayList<Integer>();
				List<Integer> failCounts = new ArrayList<Integer>();

				for (int rowIndex = firstRow; rowIndex < firstRow + expSheetRowCount; rowIndex++) {
					passCount = 0;
					failCount = 0;
					int currentRow = ++WebVerification.currentRowIndex;
					Row actualRow = actualSheet.getRow(currentRow);
					Row expectedRow = expectedSheet.getRow(rowIndex);
					List<String> actualValue = new ArrayList<String>();

					if (actualRow.getCell(0).toString().equals(expectedRow.getCell(0).toString())
							&& actualRow.getCell(1).toString().equals(expectedRow.getCell(1).toString())) {

						isrowFound = true;
						actualValue = new ArrayList<String>();
						if (actualRow == null || expectedRow == null) {
							break;
						}
						if (ITAFWebDriver.isDMApplication()) {// Minaakshi : 01-01-2020
							colCount = expectedRow.getLastCellNum();
						} else {
							colCount = expectedRow.getPhysicalNumberOfCells();
						}
						for (int columnIndex = 3; columnIndex < colCount; columnIndex++) {
							Cell actualCell = actualRow.getCell(columnIndex);
							DataFormatter fmt = new DataFormatter();
							Cell expectedCell = expectedRow.getCell(columnIndex);
							// TM: commented the code to find replacement of continue
							/*
							 * if(actualCell == null || expectedCell == null) { continue; }
							 */
							// TM: Following 'if' is replacement of the above
							if (actualCell != null || expectedCell != null) {
								String expectedValue = fmt.formatCellValue(expectedCell);
								if (ITAFWebDriver.isDMApplication() || (ITAFWebDriver.isSuiteApplication() && MainControllerSuite.transactionSilo.equalsIgnoreCase("DM"))) {// Minaakshi : 01-01-2020
									if (expectedValue.contains("|")) {

										// ***Change done by DM Team*** - 14-11-2018
										// expectedValue =
										// DMProduct.getDynamicExpectedValue(expectedValue);
										expectedValue = DMProduct.getDynamicExpectedValue(expectedValue,
												actualCell.toString());
										// ***//

									}
									if (expectedValue.contains("TEXT")) {// Minaakshi :
										// 04-01-2019
										expectedValue = DMProduct.getExpectedValueByFormula(expectedValue,
												actualCell.toString(), expectedSheet, expectedCell);
									}
								}

								// ***Change done by DM Team*** - 14-11-2018
								if (expectedValue.equalsIgnoreCase("IGNORE")) {// Minaakshi
									// :
									// 14-11-2018
									passCount += 1;
									report.setStatus("PASS");
									report.setStatus(report.getStatus());
									report.setActualValue(expectedValue + "|" + actualCell.toString());

								} else if (!actualCell.toString().equalsIgnoreCase(expectedValue)) {
									report.setStatus("FAIL");
									report.setStatus(report.getStatus());
									failCount += 1;
									report.setActualValue("FAIL |" + expectedValue + "|" + actualCell.toString());
								} else {
									passCount += 1;
									report.setStatus("PASS");
									report.setStatus(report.getStatus());
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
						rowStatus.add(report.getStatus());
						passCounts.add(passCount);
						failCounts.add(failCount);
						actualRows.add(actualValue);
						report.setReport(report);
					} else if (isrowFound == false) {
						continue;
						/*
						 * MainController.pauseFun("No Rows Found For Comparision"); break;
						 */
					}
				}
				if (rowStatus.contains("FAIL")) {
					report.setStatus("FAIL");
				}
				WriteToDetailResults(testCaseID, transactionType, "", columns, actualRows, passCounts, failCounts,
						expSheetRowCount, colCount, report, rowStatus);//Passing Cycle date blank - Mandar
				passCounts.clear();
				failCounts.clear();
				return report;
			} catch (Exception e) {
				log.error(e.getMessage(), e);
				report.setStatus("FAIL");
				report.setMessage("Issue in Excel Comparison." + e.getMessage());
				return report;

			}
		}

	
	public static void WriteToDetailResults(String testCaseID, String transactionType, List<String> columns,
			List<List<String>> actualRows, List<Integer> passCounts, List<Integer> failCounts, int rowCount,
			int colCount, Reporter report, List<String> status, String operationtype, String cycleDate)
			throws IOException {
		PrintStream print = null;
		try {
			report = report.getReport();
			report.setFromDate(Config.dtFormat.format(WebHelper.frmDate));
			report.setTestcaseId(controller.controllerTestCaseID.toString());
			if (operationtype == "Capture") {
				report.setTrasactionType(transactionType.toString());
			} else {
				report.setTrasactionType(controller.controllerTransactionType.toString());
			}
			report.setStatus(report.getStatus());

			if (WebHelper.file.exists() == false) {
				// log.info("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$");
				print = new PrintStream(WebHelper.file);
			}
			columns.remove("TestCaseID");
			columns.remove("TransactionType");
			columns.remove("CurrentDate");
			print = new PrintStream(new FileOutputStream(WebHelper.file, true));
			int usedRows = WebHelperUtil.count(WebHelper.file);
			if (usedRows == 0) {
				if (isClaimsApplication) {
					print.println(
							"Iteration,CyclDate,TestCaseID,TransactionType,CurrentDate,RowType,Status,PassCount,FailCount");
				} else {
					print.println(
							"Iteration,TestCaseID,TransactionType,CurrentDate,RowType,Status,PassCount,FailCount");
				}
			}
			usedRows = WebHelperUtil.count(WebHelper.file);

			String strCycleDate = isClaimsApplication ? myChar + cycleDate + myChar + "," : "";

			print.print(myChar + Config.cycleNumber + myChar + "," + strCycleDate + myChar + report.getTestcaseId()
					+ myChar + "," + myChar + report.getTrasactionType() + myChar + "," + myChar + report.getFromDate()
					+ myChar + "," + myChar + "Header" + myChar + "," + myChar + report.getStatus() + myChar + ","
					+ myChar + "" + myChar + "," + myChar + "" + myChar);
			int counter = 0;
			while (columns.isEmpty() == false) {
				if (counter != columns.size()) {
					print.print("," + myChar + columns.get(counter) + myChar);
					counter++;
				} else {
					break;
				}
			}
			print.println();
			rowCount = actualRows.size();
			for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
				print.print(myChar + Config.cycleNumber + myChar + "," + strCycleDate + myChar + report.getTestcaseId()
						+ myChar + "," + myChar + report.getTrasactionType() + myChar + "," + myChar
						+ report.getFromDate() + myChar + "," + myChar + "Data" + myChar + "," + myChar
						+ status.get(rowIndex).toString() + myChar + "," + myChar + passCounts.get(rowIndex) + myChar
						+ "," + myChar + failCounts.get(rowIndex) + myChar);
				counter = 0;
				while (actualRows.isEmpty() == false) {
					if (counter != actualRows.get(rowIndex).size()) {
						System.out.print(actualRows.get(rowIndex).get(counter));
						print.print("," + myChar + actualRows.get(rowIndex).get(counter) + myChar);
						counter++;
					} else {
						break;
					}
				}
				print.println();

			}
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			controller.pauseFun(e.getMessage());

		} finally {
			actualRows.clear();
			status.clear();
			columns.clear();

		}

	}

	// Writes WebVerification Results to the Excel Sheet
	// For Billing - DM
	public static void WriteToDetailResults(String testCaseID, String transactionType, String cycleDate,
			List<String> columns, List<List<String>> actualRows, List<Integer> passCounts, List<Integer> failCounts,
			int rowCount, int colCount, Reporter report, List<String> status) throws IOException // Mandar -Cycledate
																									// added
	{
		PrintStream print = null;
		try {
			report = report.getReport();
			report.setFromDate(Config.dtFormat.format(WebHelper.frmDate));
			report.setTestcaseId(controller.controllerTestCaseID.toString());

			tempStatus = report.getStatus();
			if (tempStatus.contains("FAIL") && ITAFWebDriver.isBillingApplication()) {
				failedcount++;
			}

			if (ITAFWebDriver.isBillingApplication()) {
				report.setTrasactionType(controller.controllerTransactionType.toString());
				report.setCycleDate(controller.cycleDateValue.toString());
				report.setVerificationType(transactionType);

			} else {
				report.setTrasactionType(transactionType);
			}
			report.setStatus(report.getStatus());
			if (ITAFWebDriver.isPASApplication() || ITAFWebDriver.isDMApplication())// Minaakshi : 01-11-2019
				report.setGroupName(controller.controllerGroupName.toString());

			if (WebHelper.file.exists() == false) {
				System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$");
				print = new PrintStream(WebHelper.file);
			}
			columns.remove("TestCaseID");
			columns.remove("TransactionType");
			columns.remove("CurrentDate");
			columns.remove("CycleDate");
			columns.remove("VerificationType");
			// }

			print = new PrintStream(new FileOutputStream(WebHelper.file, true));
			int usedRows = WebHelperUtil.count(WebHelper.file);
			if (usedRows == 0) {
				if (ITAFWebDriver.isPASApplication() || ITAFWebDriver.isDMApplication()) {// Minaakshi : 01-11-2019
					print.println(
							"GroupName,TestCaseID,TransactionType,CurrentDate,RowType,Status,PassCount,FailCount,Compare Result,Quote Number");
				} else if (ITAFWebDriver.isBillingApplication()) {
					print.println(
							"TestCaseID,TransactionType,VerificationType,CycleDate,RowType,Status,PassCount,FailCount,Compare Result");
				} else
					print.println(
							"Iteration,TestCaseID,TransactionType,CurrentDate,RowType,Status,PassCount,FailCount");
			}
			usedRows = WebHelperUtil.count(WebHelper.file);

			if (ITAFWebDriver.isPASApplication() || ITAFWebDriver.isDMApplication()) // Minaakshi : 01-01-2020
			{
				print.print(myChar + report.getGroupName() + myChar + "," + myChar + report.getTestcaseId() + myChar
						+ "," + myChar + report.getTrasactionType() + myChar + "," + myChar + report.getFromDate()
						+ myChar + "," + myChar + "Header" + myChar + "," + myChar + report.getStatus() + myChar + ","
						+ myChar + "" + myChar + "," + myChar + "" + myChar);

			} else if (ITAFWebDriver.isBillingApplication()) {
				print.print(myChar + report.getTestcaseId() + myChar + "," + myChar + report.getTrasactionType()
						+ myChar + "," + myChar + report.getVerificationType() + myChar + "," + myChar
						+ report.getCycleDate() + myChar + "," + myChar + "Header" + myChar + "," + myChar + "" + myChar
						+ "," + myChar + "" + myChar + "," + myChar + "" + myChar);
			} else {
				print.print(myChar + Config.cycleNumber + myChar + "," + myChar + report.getTestcaseId() + myChar + ","
						+ myChar + report.getTrasactionType() + myChar + "," + myChar + report.getFromDate() + myChar
						+ "," + myChar + "Header" + myChar + "," + myChar + report.getStatus() + myChar + "," + myChar
						+ "" + myChar + "," + myChar + "" + myChar);
			}

			int counter = 0;
			while (columns.isEmpty() == false) {
				if (counter != columns.size()) {
					print.print("," + myChar + columns.get(counter) + myChar);
					counter++;
				} else {
					break;
				}
			}
			print.println();
			rowCount = actualRows.size();
			for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
				if (ITAFWebDriver.isPASApplication() || ITAFWebDriver.isDMApplication())// Minaakshi : 01-10-2019 -
				// Added condition for DM
				{
					print.print(myChar + report.getGroupName() + myChar + "," + myChar + report.getTestcaseId() + myChar
							+ "," + myChar + report.getTrasactionType() + myChar + "," + myChar + report.getFromDate()
							+ myChar + "," + myChar + "Data" + myChar + "," + myChar + status.get(rowIndex).toString()
							+ myChar + "," + myChar + passCounts.get(rowIndex) + myChar + "," + myChar
							+ failCounts.get(rowIndex) + myChar);

				} else if (ITAFWebDriver.isBillingApplication()) {
					print.print(myChar + "" + myChar + "," + myChar + "" + myChar + "," + myChar + "" + myChar + ","
							+ myChar + "" + myChar + "," + myChar + "Data" + myChar + "," + myChar
							+ status.get(rowIndex).toString() + myChar + "," + myChar + passCounts.get(rowIndex)
							+ myChar + "," + myChar + failCounts.get(rowIndex) + myChar);

				} else
					print.print(myChar + Config.cycleNumber + myChar + "," + myChar + report.getTestcaseId() + myChar
							+ "," + myChar + report.getTrasactionType() + myChar + "," + myChar + report.getFromDate()
							+ myChar + "," + myChar + "Data" + myChar + "," + myChar + status.get(rowIndex).toString()
							+ myChar + "," + myChar + passCounts.get(rowIndex) + myChar + "," + myChar
							+ failCounts.get(rowIndex) + myChar);
				counter = 0;
				while (actualRows.isEmpty() == false) {
					if (counter != actualRows.get(rowIndex).size()) {
						System.out.print(actualRows.get(rowIndex).get(counter));
						print.print("," + myChar + actualRows.get(rowIndex).get(counter) + myChar);
						counter++;
					} else {
						break;
					}
				}
				print.println();

			}
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			controller.pauseFun(e.getMessage());

		} finally {

			if (!ITAFWebDriver.isDMApplication()) {
				actualRows.clear();
				status.clear();
				columns.clear();
			}

			if (failedcount > 0 && ITAFWebDriver.isBillingApplication())// Marking FAIL in summary results.--

			{
				webDriver.getReport().setStatus("FAIL");
			}

		}

	}

	public static void WriteToCompareDetailResults(String testCaseID, String transactionType, List<String> columns,
			List<List<String>> columnsData, int rowCount, int colCount, Reporter report, String expectedValue,
			String stractualValue, String actualHeaderValue, String operationtype, String cycleDate) throws IOException

	{
		if (ITAFWebDriver.isSuiteApplication()) {
			WebHelperUtil.WriteToDetailResultsSuite(expectedValue, stractualValue, actualHeaderValue);
		} else {
			PrintStream print = null;
			try {
				report = report.getReport();
				report.setFromDate(Config.dtFormat.format(WebHelper.frmDate));
				report.setTestcaseId(controller.controllerTestCaseID.toString());
				if (operationtype == "Capture") {
					report.setTrasactionType(transactionType.toString());
				} else {
					report.setTrasactionType(controller.controllerTransactionType.toString());
				}
				tempStatus = report.getStatus();
				if (tempStatus.contains("FAIL") && ITAFWebDriver.isBillingApplication()) {
					failedcount++;
				}
				report.setStatus(report.getStatus());
				report.setMessage(report.getMessage());

				if (WebHelper.dashboardfile.exists() == false) {
					// log.info("#########################################");
					print = new PrintStream(WebHelper.dashboardfile);
				}

				print = new PrintStream(new FileOutputStream(WebHelper.dashboardfile, true));
				int usedRows = WebHelperUtil.count(WebHelper.dashboardfile);
				log.info("Dashboard execution" + usedRows);
				if (usedRows == 0) {
					// TM: Added println instead of print
					print.println(
							"Iteration,CycleDate,TransactionName,TestCaseID,Date,RowType,Status,Message,FieldName,ExpectedValue,ActualValue");
				}
				usedRows = WebHelperUtil.count(WebHelper.dashboardfile);
				log.info("Dashboard execution" + usedRows);
				print.println(myChar + Config.cycleNumber + myChar + "," + myChar + cycleDate + myChar + "," + myChar
						+ report.getTrasactionType() + myChar + "," + myChar + report.getTestcaseId() + myChar + ","
						+ myChar + report.getFromDate() + myChar + "," + myChar + report.getTrasactionType() + "_1"
						+ myChar + "," + myChar + report.getStatus() + myChar + "," + myChar + report.getMessage()
						+ myChar + "," + myChar + actualHeaderValue + myChar + "," + myChar + expectedValue + myChar
						+ "," + myChar + stractualValue + myChar);
			}

			catch (Exception e) {
				log.error(e.getMessage(), e);
				controller.pauseFun(e.getMessage());

			} finally {
				columns.clear();
				columnsData.clear();
				if (failedcount > 0 && ITAFWebDriver.isBillingApplication())//// Vinay B- Marking FAIL in summary
																			//// results.
				{
					webDriver.getReport().setStatus("FAIL");
				}
			}
		}

	}

	/**
	 * This method copy the content of current workbook to destination workbook
	 * 
	 * @param studentsSheet - Source workbook
	 * @param name          - Current worksheet name
	 * @param rows          - number of rows in current worksheet
	 * @param output_file   - destination file outputstream to write the content
	 * @param destWorksheet - destination worksheet
	 * @throws IOException
	 */
	public static void copyFileToCreateNewOne(Workbook studentsSheet, String name, int rows,
			FileOutputStream output_file, Sheet destWorksheet) throws IOException {
		for (int i = 0; i <= rows; i++) {
			Sheet worksheet = studentsSheet.getSheet(name);
			if (i == 0) {

				createOnlyHeader(worksheet, destWorksheet);

			} else {
				copyToFileOnlyData(worksheet, destWorksheet, i);

			}
			studentsSheet.write(output_file);
		}
	}

	/**
	 * This method copied the header row from source excel worksheet to destination
	 * worksheet
	 * 
	 * @param currentWorksheet - source worksheet
	 * @param destWorksheet    - destination worksheet
	 */
	public static void createOnlyHeader(Sheet currentWorksheet, Sheet destWorksheet) {
		// create the header
		HashMap<String, Integer> map = WebHelperUtil.getValueFromHashMap(currentWorksheet);

		Set<Entry<String, Integer>> set = map.entrySet();

		Iterator<Entry<String, Integer>> iter = set.iterator();
		Row destRow = destWorksheet.createRow(0);
		while (iter.hasNext()) {
			Entry<String, Integer> key = iter.next();
			Cell cellA1 = destRow.createCell(key.getValue().intValue());
			cellA1.setCellValue(key.getKey());
		}
	}

	/**
	 * This method creates the header for report excel file
	 * 
	 * @param currentWorksheet - String array of header details
	 * @param destWorksheet    - Worksheet where the header will be populated
	 */
	public static void createOnlyHeaderFromInput(String[] currentWorksheet, Sheet destWorksheet) {
		// create the header
		Row destRow = destWorksheet.createRow(0);
		for (int col = 0; col < currentWorksheet.length; col++) {
			Cell cellA1 = destRow.createCell(col);
			cellA1.setCellValue(currentWorksheet[col]);
		}

	}

	/**
	 * This method copy the report data to Report excel sheet without header row.
	 * 
	 * @param worksheet     - current worksheet
	 * @param destWorksheet - destination worksheet
	 * @param row           - row number where data to be written
	 */

	public static void copyToFileOnlyData(Sheet worksheet, Sheet destWorksheet, int row) {
		// Create the records
		Row values = worksheet.getRow(row);
		int totalColumnCount = values.getLastCellNum();
		System.out.println("Total columns in records: " + totalColumnCount);
		Row destRow = destWorksheet.createRow(row);

		for (int col = 0; col < totalColumnCount; col++) {
			Cell destCell = destRow.createCell(col);
			Cell sourceCell = values.getCell(col);
			DataFormatter formatter = new DataFormatter();

			String value = "";
			if (sourceCell != null) {
				CellType cellType = sourceCell.getCellType(); // For Apache POI 3.15+
				// or
				// CellType cellType = sourceCell.getCellType(); // For Apache POI 4.0.0+

				switch (cellType) {
				case STRING:
					value = sourceCell.getStringCellValue();
					break;
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(sourceCell)) {
						value = formatter.formatCellValue(sourceCell);
					} else {
						value = formatter.formatCellValue(sourceCell);
					}
					break;
				case BOOLEAN:
					value = String.valueOf(sourceCell.getBooleanCellValue());
					break;
				case FORMULA:
					value = sourceCell.getCellFormula();
					break;
				default:
					value = ""; // Handle other types as needed
				}
			}

			System.out.println(value);
			destCell.setCellValue(value);
		}
	}

	/**
	 * This method copy the report data to Report excel sheet without header row.
	 * 
	 * @param destWorksheet - Destination worksheet where data to be written
	 * @param report        - value object to store business data to be written to
	 *                      report
	 * @param reportType    - Type of report whether simple or detailed
	 */

	public static void copyToFileOnlyData(Sheet destWorksheet, Reporter report, String reportType) {
		// Create the records
		int row = destWorksheet.getLastRowNum();
		Row currentRow = destWorksheet.getRow(row);
		int totalColumnCount = destWorksheet.getRow(0).getLastCellNum();
		System.out.println("Total columns in records: " + totalColumnCount);
		Row destRow = destWorksheet.createRow(row + 1);

		for (int col = 0; col < totalColumnCount; col++) {
			Cell destCell = destRow.createCell(col);

			Cell sourceCell = currentRow.getCell(col);
			DataFormatter formatter = new DataFormatter();

			String value = "";
			if (sourceCell != null) {
				CellType cellType = sourceCell.getCellType(); // For Apache POI 3.15+
				// or
				// CellType cellType = sourceCell.getCellType(); // For Apache POI 4.0.0+

				String reportValue = selectReportData(col, report, reportType);

				value = selectDataType(sourceCell, formatter, reportValue, cellType);
			}

			System.out.println(value);
			destCell.setCellValue(value);
		}
	}

	/**
	 * This method returns the value to written to each cell in excel worksheet
	 * 
	 * @param col    - column of the worksheet
	 * @param report - value object to store report data
	 * @param type   - to be generated report type
	 * @return - return string representation of the data to written to worksheet
	 */
	private static String selectReportData(int col, Reporter report, String type) {
		if (type.equalsIgnoreCase("reports")) {
			if (col == 0) {
				return report.getGroupName();
			} else if (col == 1) {
				return Config.cycleNumber;
			} else if (col == 2) {
				return report.getTestcaseId();
			} else if (col == 3) {
				return report.getTrasactionType();
			} else if (col == 4) {
				return report.getTestDescription();
			} else if (col == 5) {
				return report.getFromDate();
			} else if (col == 6) {
				return report.getToDate();
			} else if (col == 7) {
				return report.getStatus();
			} else if (col == 8) {
				return report.getMessage();
			} else if (col == 9) {
				return report.getScreenShot();
			} else if (col == 10) {
				return report.getStrQuoteId();
			}
		} else if (type.equalsIgnoreCase("detailed")) {
			if (col == 0) {
				return Config.cycleNumber;
			} else if (col == 1) {
				return report.getTestcaseId();
			} else if (col == 2) {
				return report.getTrasactionType();
			} else if (col == 3) {
				return report.getToDate();
			} else if (col == 4) {
				return report.getColumnName();
			} else if (col == 5) {
				return report.getStatus();
			} else if (col == 6) {
				return report.getPassCount();
			} else if (col == 7) {
				return report.getFailCount();
			} else if (col == 8) {
				return report.getActualValue();
			} else if (col == 9) {
				return report.getStrQuoteId();
			}
		}
		return null;
	}

	/**
	 * This method format the data type to display properly in excel worksheet
	 * 
	 * @param root         - Cell of workssheet where data to be found
	 * @param fmt          - Date formatter used
	 * @param displayValue - raw data fethed from reported value object
	 * @param type         - data type of the column in report worksheet
	 * @return formatted output string to be written to report
	 */

	private static String selectDataType(Cell root, DataFormatter fmt, String displayValue, CellType type) {
		switch (type) {
		case BLANK:
			displayValue = "";
			break;
		case NUMERIC:
			displayValue = fmt.formatCellValue(root);
			break;
		case STRING:
			displayValue = root.getStringCellValue();
			break;
		case BOOLEAN:
			displayValue = Boolean.toString(root.getBooleanCellValue());
			break;
		case ERROR:
			displayValue = "error";
			break;
		case FORMULA:
			displayValue = root.getCellFormula();
			break;
		case _NONE:
			break;
		default:
			break;
		}
		return displayValue;
	}

	private static String calTimeElapsed(String dateStart, String dateStop) {
		SimpleDateFormat format = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss:SSSS");

		Date d1 = null;
		Date d2 = null;
		StringBuilder timeElapsed = new StringBuilder();
		try {
			d1 = format.parse(dateStart);
			d2 = format.parse(dateStop);

			// in milliseconds
			long diff = d2.getTime() - d1.getTime();
			// long diffMilliSeconds = diff % 1000;//Sel4
			long diffSeconds = diff / 1000 % 60;
			long diffMinutes = diff / (60 * 1000) % 60;
			long diffHours = diff / (60 * 60 * 1000) % 24;
			// long diffDays = diff / (24 * 60 * 60 * 1000);//Sel4

			long tempArray[] = { diffHours, diffMinutes, diffSeconds };
			for (long tempDiff : tempArray) {
				timeElapsed.append(String.format("%02d", tempDiff) + ":");
			}
			if (ITAFWebDriver.isDMApplication() || ITAFWebDriver.isClaimsApplication()
					|| ITAFWebDriver.isSuiteApplication())// Changes done by PAS COE team - 25/09/2019
			{
				// return Long.toString(diffSeconds);//Changes done by PAS COE team - 26/09/2019
				double sec = diff / 1000.0;
				return Double.toString(sec);
			} else // if(ITAFWebDriver.isPASApplication() )
			{
				return timeElapsed.toString();
			}
			// else timeElapsed.append(String.format("%04d", diffMilliSeconds));

		} catch (Exception e) {
			e.printStackTrace();
		}
		return timeElapsed.toString();

	}

	public static void fillValidationValueSheet() throws Exception {
		String testCaseID = controller.controllerTestCaseID.toString();

		String transDirFilePath = getTranDirPAth();
		transDirFilePath = Config.inputDataFilePath + transDirFilePath;

		Workbook workBook = null;
		Sheet valueSheet = null;
		InputStream myXls = new FileInputStream(transDirFilePath);
		ZipSecureFile.setMinInflateRatio(-1.0d);
		if (transDirFilePath.endsWith(".xls")) {
			workBook = new HSSFWorkbook(myXls);
		} else if (transDirFilePath.endsWith(".xlsx") | transDirFilePath.endsWith(".xlsm")) {
			workBook = new XSSFWorkbook(myXls);
		}
		valueSheet = workBook.getSheet("Values");

		HashMap<String, Integer> headerValues;
		headerValues = WebHelperUtil.getValueFromHashMap(valueSheet);

		int tranRowNo = 0;
		for (int i = 0; i <= valueSheet.getLastRowNum(); i++) {
			String tempTestCaseID = valueSheet.getRow(i).getCell(0).getStringCellValue();
			if (tempTestCaseID.equals(testCaseID)) {
				tranRowNo = i;
				break;
			}
		}

		for (HashMap.Entry<String, Integer> entry : headerValues.entrySet()) {
			String colName = entry.getKey();
			int colIndex = entry.getValue();
			Row testCaseIdRow = valueSheet.getRow(tranRowNo);
			if (testCaseIdRow.getCell(colIndex) == null
					|| testCaseIdRow.getCell(colIndex).getStringCellValue().equals("")) {
				Cell tempCell = testCaseIdRow.createCell(colIndex);
				String tempCellValue = JsonUtility.readFromUniqueNumberSheet(testCaseID, colName);
				tempCell.setCellValue(tempCellValue);
			}
		}
		FileOutputStream out = new FileOutputStream(new File(transDirFilePath));
		workBook.write(out);
		workBook.close();
	}

	public static String getTranDirPAth() throws IOException {
		HashMap<String, Integer> headerValues;
		Sheet transInputSheet = GetSheet(Config.transactionInputFilePath, "Web_Transaction_Input_Files");
		headerValues = WebHelperUtil.getValueFromHashMap(transInputSheet);

		String transactionType = controller.controllerTransactionType.toString();
		int transactionTypeIndex = headerValues.get("TransactionType");

		Row tempRow;
		for (int i = 0; i <= transInputSheet.getLastRowNum(); i++) {
			tempRow = transInputSheet.getRow(i);
			String tempTransType = tempRow.getCell(transactionTypeIndex).getStringCellValue();
			if (tempTransType.equals(transactionType)) {
				String tempDirPath = tempRow.getCell(transactionTypeIndex + 1) + "\\"
						+ tempRow.getCell(transactionTypeIndex + 2);
				return tempDirPath;
			}
		}
		return null;
	}

	public static int getUsedRowCount(String filePath, String tabName) throws IOException {
		try (FileInputStream excelFile = new FileInputStream(new File(filePath));
				Workbook workbook = new XSSFWorkbook(excelFile)) {
			Sheet sheet = workbook.getSheet(tabName);
			int rowTotal = sheet.getLastRowNum();
			if (rowTotal > 0 || sheet.getPhysicalNumberOfRows() > 0) {
				rowTotal++;
			}
			return rowTotal;
		}
	}

	public static void createExcelIfNotPresent(String FilePath, String TabName) throws IOException {
		File file = new File(FilePath);
		if (!(file.isFile() && file.exists())) {
			XSSFWorkbook workbook = new XSSFWorkbook();
			FileOutputStream out = new FileOutputStream(file);
			workbook.createSheet(TabName);
			workbook.write(out);
			out.close();
			workbook.close();
		}
	}

	public static void writeRowToExcel(String FilePath, String TabName, int rowNum, int colStart, String[] rowValues)
			throws IOException {
		FileInputStream excelFile = new FileInputStream(new File(FilePath));
		Workbook workbook = new XSSFWorkbook(excelFile);
		Sheet sheet = workbook.getSheet(TabName);

		Row curRow = sheet.createRow(rowNum);
		for (int i = colStart, rVal = 0; rVal < rowValues.length; i++, rVal++) {
			Cell cell = curRow.createCell(i);
			cell.setCellValue(rowValues[rVal]);
		}
		FileOutputStream fileOut = new FileOutputStream(new File(FilePath));
		workbook.write(fileOut);
		fileOut.close();
		workbook.close();
	}

	public static void writeReportSuite(Reporter report) throws IOException {
		PrintStream print = null;
		try {

			File file = new File(Config.resultOutput);
			if (file.exists() == true && StringUtils.equalsIgnoreCase(Config.appendResultOutputGeneric, "false")
					&& MainControllerSuite.DResult == 1) {
				file.delete();
			}
			MainControllerSuite.DResult++;
			report.setReport(report);
			report = report.getReport();
			report.setSilo(MainControllerSuite.transactionSilo);
			report.setIteration(Config.cycleNumber);
			report.setFromDate(Config.dtFormat.format(MainControllerSuite.frmDate));
			report.setToDate(Config.dtFormat.format(new Date()));
			report.setTransactionVerificationType(MainControllerSuite.transactionVerificationType);
			if (StringUtils.isBlank(report.getMessage()))
				report.setMessage("");

			if (StringUtils.contains(report.getMessage(), "\"")) {
				report.setMessage(report.getMessage().replaceAll("\"", "'"));
			}
			if (StringUtils.contains(report.getMessage(), ",")) {
				report.setMessage("\"\"" + report.getMessage() + "\"\"");
			}

			if (StringUtils.isBlank(report.getTestDescription()))
				report.setTestDescription("");

			if (StringUtils.isBlank(report.getGroupName()))
				report.setGroupName("");

			if (report.getTrasactionType().equalsIgnoreCase("ChangeBusinessdate")) {
				report.setTestcaseId("Common");
				report.setGroupName("Common");
				report.setSilo("Billing");
				report.setTransactionVerificationType("NA");
			}

			if (WebService.isNoResponseFileTrue())
				report.setMessage(WebHelper.description);

			if (WebHelper.faultstring == true) {
				if (WebService.isNodeAvailable()) {
					report.setMessage("<" + WebService.getreportNodeValue() + "> " + WebHelper.description);
				} else {
					report.setMessage(WebHelper.description);
				}
			}
			if (report.getTrasactionType().startsWith("WebService"))
				report.setScreenShot("");

			if (WebHelper.success == true)
				report.setMessage(WebHelper.description);

			String timeElapsed = calTimeElapsed(report.getFromDate(), report.getToDate());
			print = new PrintStream(new FileOutputStream(file, true));
			int usedRows = WebHelperUtil.count(file);
			if (usedRows == 0) {
				print.print(
						"GroupName,Iteration,Silo,TestCaseID,CycleDate,TransactionType,TransactionVerificationType,TestCaseDescription,StartDate,EndDate,ExecutionTime(Seconds),Status,ExecutionDescription,Screenshot");
				print.println();
			}
			usedRows = WebHelperUtil.count(file);
			print = new PrintStream(new FileOutputStream(file, true));
			// String fromDate = report.getFromDate().substring(0,
			// (report.getFromDate().lastIndexOf(":")));//Sel4
			// String toDate = report.getToDate().substring(0,
			// (report.getToDate().lastIndexOf(":")));//Sel4

			print.print(myChar + report.getGroupName() + myChar + "," + myChar + report.getIteration() + myChar + ","
					+ myChar + report.getSilo() + myChar + "," + myChar + report.getTestcaseId() + myChar + "," + myChar
					+ report.getCycleDate() + myChar + "," + myChar + report.getTrasactionType() + myChar + "," + myChar
					+ report.getTransactionVerificationType() + myChar + "," + myChar + report.getTestDescription()
					+ myChar + "," + myChar + report.getFromDate() + myChar + "," + myChar + report.getToDate() + myChar
					+ "," + myChar + timeElapsed + myChar + "," + myChar + report.getStatus().toUpperCase() + myChar
					+ "," + myChar + report.getMessage() + myChar + "," + myChar + report.getScreenShot() + myChar);
			print.println();

			if (report.getTrasactionType().equalsIgnoreCase("ChangeBusinessdate")
					&& report.getStatus().equalsIgnoreCase("Fail")) {
				controller.stopExecution = true;
			}

		} catch (IOException ie) {
			log.error("Issue while writing report : " + ie.getMessage());
			controller.pauseFun("Issue while writing report : " + ie.getMessage());
		} finally {
			webDriver.getReport().setScreenShot("");

		}
	}

}
