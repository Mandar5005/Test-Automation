package com.majesco.itaf.main;

import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.net.URL;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;
import java.util.Set;
import javax.net.ssl.HttpsURLConnection;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.interactions.Coordinates;
import org.openqa.selenium.interactions.Locatable;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.sikuli.script.App;
import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;
import com.majesco.itaf.util.CalendarSnippet;
import com.majesco.itaf.util.CommonExpectedConditions;
import com.majesco.itaf.util.Constants;
import com.majesco.itaf.util.DMProduct;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.util.JDBCConnection;
import com.majesco.itaf.util.Jacob;
import com.majesco.itaf.util.WaitTool;
import com.majesco.itaf.verification.WebVerification;
import com.majesco.itaf.vo.Reporter;
import com.majesco.itaf.webservice.WebService;

public class WebHelperDM {

	private final static Logger log = LogManager.getLogger(WebHelperDM.class.getName());
	public static Cell transactionType = null;
	public static int TotalpassCount;
	public static int TotalfailCount;
	public static int DResult = 1;
	boolean stillChanging = true;
	public static Date toDate = null;
	public static String testCase = null;
	public static Screen sikuliScreen = null;
	public static List<String> searchValue1 = null;
	public static Boolean pageLoaded = false;
	private static List<String> formList = new ArrayList<String>();
	private static List<String> presentForms = new ArrayList<String>();
	private static List<String> additionalForms = new ArrayList<String>();
	private static List<String> missingForms = new ArrayList<String>();
	private static List<String> expectedForms = new ArrayList<String>();
	private static boolean isDMApplication = ITAFWebDriver.isDMApplication();
	public static String sqlQuery = "";
	public static String readFromColName = "";
	public static String toBeFetchedDBColName = "";
	public static String expectedDBStatus = "";
	public static String writeToColName = "";
	public static Wait<WebDriver> waitForElementPresence;
	public static Wait<WebDriver> dmgebtWait;
	final static long NANOSEC_PER_SEC = 1000l * 1000 * 1000;
	public static long startTime_batch;
	static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	static MainController controller = ObjectFactory.getMainController();

	static {
		WebHelper.wait = new WebDriverWait(Automation.driver, Duration.ofSeconds(Integer.parseInt(Config.timeOut)));
	}

	static void implementWait() {

		dmgebtWait = new FluentWait<>(Automation.driver)
				.withTimeout(Duration.ofSeconds(Integer.parseInt(Config.timeOut))).pollingEvery(Duration.ofSeconds(3))
				.ignoring(NoSuchElementException.class);

		waitForElementPresence = new FluentWait<>(Automation.driver).withTimeout(Duration.ofSeconds(30))
				.pollingEvery(Duration.ofSeconds(3)).ignoring(NoSuchElementException.class);

	}
	public static void GetCellInfo(String FilePath, Row rowValues, int valuesRowIndex, int valuesRowCount)
			throws IOException
	{
		boolean reportWrittenAfterException = false;
		try {

			WebHelper.frmDate = new Date();
			WebHelper.isDynamicNumFound = true;
			List<WebElement> controlList = null;
			// String testCase = null;
			String ctrlValue = null;
			InputStream myXls = new FileInputStream(FilePath);
			Workbook workBook = null; // new HSSFWorkbook(myXls);
			if (FilePath.endsWith(".xls")) {
				workBook = new HSSFWorkbook(myXls);
			} else if (FilePath.endsWith(".xlsx") | FilePath.endsWith(".xlsm")) {
				workBook = new XSSFWorkbook(myXls);
			}
			WebHelper.format = workBook.createDataFormat();
			Sheet sheetStructure = workBook.getSheet("Structure");
			log.info("TestCaseID: " + controller.controllerTestCaseID.toString() + ", TransactionName: "
					+ controller.controllerTransactionType.toString());
			int rowCount = sheetStructure.getLastRowNum() + 1;
			Sheet headerValues = ExcelUtility.GetSheet(FilePath, "Values");
			String fromDate = Config.dtFormat.format(WebHelper.frmDate);
			webDriver.getReport().setFromDate(fromDate);
			WebHelper.structureHeader = WebHelperUtil.getValueFromHashMap(sheetStructure);
			WebHelper.columnName = null;
			int dynamicIndexNumber;// Added for Action Loop
			String imageType, indexVal, controlName, executeFlag, action, logicalName, controltype, controlID,
					dynamicIndex, newDynamicIndex, rowNo, colNo;// newly //columnName1, CompareText//Sel4
			webDriver.getReport().setMessage("");
			webDriver.getReport().setStatus("PASS");
			for (int rowIndex = 1; rowIndex < rowCount && !controller.pauseExecution; rowIndex++) {
				// structureRow = sheetStructure.getRow(rowIndex);
				controlName = WebHelperUtil.getCellData("ControlName", sheetStructure, rowIndex,
						WebHelper.structureHeader);// structureRow.getCell(3);
				executeFlag = WebHelperUtil.getCellData("ExecuteFlag", sheetStructure, rowIndex,
						WebHelper.structureHeader);// structureRow.getCell(0);

				if (executeFlag.toString().equals("Y")) {
					WebElement webElement = null;
					imageType = WebHelperUtil.getCellData("ImageType", sheetStructure, rowIndex,
							WebHelper.structureHeader);

					action = WebHelperUtil.getCellData("Action", sheetStructure, rowIndex, WebHelper.structureHeader);// structureRow.getCell(1);
					logicalName = WebHelperUtil.getCellData("LogicalName", sheetStructure, rowIndex,
							WebHelper.structureHeader);// structureRow.getCell(2);

					controltype = WebHelperUtil.getCellData("ControlType", sheetStructure, rowIndex,
							WebHelper.structureHeader);// structureRow.getCell(4);
					controlID = WebHelperUtil.getCellData("ControlID", sheetStructure, rowIndex,
							WebHelper.structureHeader);// structureRow.getCell(6);

					indexVal = WebHelperUtil.getCellData("Index", sheetStructure, rowIndex, WebHelper.structureHeader);// structureRow.getCell(7);
					WebHelper.columnName = WebHelperUtil.getCellData("ColumnName", sheetStructure, rowIndex,
							WebHelper.structureHeader);

					// columnName1 = getCellData("CompareText", sheetStructure,
					// rowIndex, structureHeader);
					rowNo = WebHelperUtil.getCellData("RowNo", sheetStructure, rowIndex, WebHelper.structureHeader);
					colNo = WebHelperUtil.getCellData("ColumnNo", sheetStructure, rowIndex, WebHelper.structureHeader);
					dynamicIndex = WebHelperUtil.getCellData("DynamicIndex", sheetStructure, rowIndex,
							WebHelper.structureHeader);

					// Added code for loop
					log.info("iTAF:" + logicalName + " " + rowIndex);
					//log.info("iTAF:" + logicalName + " " + rowIndex);

					/*
					 * Below code has been written To Handle condition for multiple rows in excel
					 * sheet
					 */

					if (imageType.equalsIgnoreCase("START_SCRIPT")) {

						try {
							webElement = getElementByType(controlID, controlName, WebHelper.control, imageType,
									ctrlValue);

							if (webElement.isDisplayed()) {
								log.info("Proceed");
							}
						} catch (Exception e) {
							log.error(e.getMessage(), e);
							for (int searchrow = rowIndex; searchrow < rowCount; searchrow++) {
								imageType = WebHelperUtil.getCellData("ImageType", sheetStructure, rowIndex,
										WebHelper.structureHeader);

								rowIndex++;
								if (imageType.equalsIgnoreCase("END_SCRIPT")) {

									break;
								}
							}
						}
					}

					if (action.equalsIgnoreCase("LOOP")) {
						WebHelper.loopRow = rowIndex + 1;
					}

					// if rownum != 1 , then do below steps
					if ((valuesRowIndex != ExcelUtility.firstRow) && (dynamicIndex.length() > 0)) // valuesRowIndex
					{

						dynamicIndexNumber = Integer
								.parseInt(dynamicIndex.substring(dynamicIndex.length() - 1, dynamicIndex.length()));

						if (ExcelUtility.dynamicNum == 0) {
							ExcelUtility.dynamicNum = dynamicIndexNumber + 1;
							WebHelper.isDynamicNumFound = false;

						} else if (ExcelUtility.dynamicNum != 0 && WebHelper.isDynamicNumFound) {
							ExcelUtility.dynamicNum = ExcelUtility.dynamicNum + 1;
							WebHelper.isDynamicNumFound = false;
						}

						newDynamicIndex = dynamicIndex.replace(String.valueOf(dynamicIndexNumber),
								String.valueOf(ExcelUtility.dynamicNum));

						controlName = controlName.replace(dynamicIndex, newDynamicIndex);
					}

					/**
					 * Stop the execution of the current test case unexpected alert
					 **/

					if (!action.equalsIgnoreCase("LOOP") && !action.equalsIgnoreCase("END_LOOP")
							&& !action.equalsIgnoreCase("TableInput_End")) {

						// boolean isControlValueFound =false;
						if (WebHelper.valuesHeader.isEmpty() == true) {
							WebHelper.valuesHeader = WebHelperUtil.getValueFromHashMap(headerValues);
						}
						Object actualValue = null;
						if (logicalName != null) {
							actualValue = WebHelper.valuesHeader.get(logicalName.toString());
						} // headerRow.getCell(colIndex);
						if (actualValue == null) {
							log.info("actualValue is Null");
						} else {
							// int colIndex =
							// Integer.parseInt(actualValue.toString());
							// controlValue = rowValues.getCell(colIndex);

							ctrlValue = WebHelperUtil.getCellData(logicalName, headerValues, valuesRowIndex,
									WebHelper.valuesHeader);

							// controlValue=getCellData(logicalName,headerValues,
							// valuesRowIndex, valuesHeader);

							WebHelper.testcaseID = rowValues
									.getCell(Integer.parseInt(WebHelper.valuesHeader.get("TestCaseID").toString()));

							if (WebHelper.testcaseID == null) {
								testCase = "";
							} else {
								testCase = WebHelper.testcaseID.toString();
							}
							transactionType = rowValues.getCell(WebHelper.valuesHeader.get("TransactionType"));
						}

						if ((action.equals("I") && !ctrlValue.isEmpty()) || (action.equals("V") && !ctrlValue.isEmpty())
						// Mrinmayee 06-12-2018
								|| (action.equals("ActionClick") && !ctrlValue.isEmpty()) // Minaakshi
								// :
								// 01-03-2019
								|| (action.equals("V_availabilityStatus") && !ctrlValue.isEmpty())
								|| (action.equals("V_text") && !ctrlValue.isEmpty())
								|| (action.equals("V_attributeValue") && !ctrlValue.isEmpty())
								|| (action.equals("V_edit") && !ctrlValue.isEmpty())
								|| (action.equals("V_disableStatus") && !ctrlValue.isEmpty())
								|| (action.equals("V_checkboxStatus") && !ctrlValue.isEmpty())
								|| (action.equals("VFPresence") && !ctrlValue.isEmpty())
								|| (action.equals("VFormPresence") && !ctrlValue.isEmpty())
								|| (action.equals("InputifExist") && !ctrlValue.isEmpty())
								|| (action.equals("VFormCount") && !ctrlValue.isEmpty())
								|| (!action.equals("I") && !action.equals("V") && !action.equals("V_availabilityStatus")
										&& !action.equals("V_edit") && !action.equals("V_text")
										&& !action.equals("V_attributeValue") && !action.equals("V_disableStatus")
										&& !action.equals("VFPresence") && !action.equals("V_checkboxStatus")
										&& !action.equals("VFormCount") && !action.equals("VerifyPopUpElement")
										&& !action.equals("ActionClick"))) {

							if (!controltype.startsWith("Sikuli")) {
								if (!action.equalsIgnoreCase("LOOP") && !controltype.equalsIgnoreCase("Wait")
										&& !controltype.equalsIgnoreCase("Wait_DM")
										// Minaakshi : 05-02-2019
										&& !controltype.equalsIgnoreCase("OpenAPI")// Minaakshi
										// :
										// 01-06-2019
										&& !action.equalsIgnoreCase("END_LOOP")
										&& !controltype.equalsIgnoreCase("Browser")
										&& !controltype.equalsIgnoreCase("CloseBrowser")
										&& !controltype.equalsIgnoreCase("AttributeisPresent")
										&& !controltype.equalsIgnoreCase("CloseAndLaunchNewBrowser")
										&& !controltype.equalsIgnoreCase("PageRefresh")
										&& !controltype.equalsIgnoreCase("WaitForPageToLoad")
										&& !controltype.equalsIgnoreCase("WaitToLoad")
										&& !controltype.equalsIgnoreCase("WaitUntilElementInvisible")
										&& !controltype.equalsIgnoreCase("filedownload")
										&& !controltype.equalsIgnoreCase("NewBrowser")
										&& !controltype.equalsIgnoreCase("Window")
										&& !controltype.equalsIgnoreCase("Alert")
										&& !controltype.equalsIgnoreCase("URL")
										&& !controltype.equalsIgnoreCase("WaitForJS") && !controltype.contains("Robot")
										&& !controltype.equalsIgnoreCase("Calendar")
										&& !controltype.equalsIgnoreCase("CalendarNew")
										&& !controltype.equalsIgnoreCase("CalendarIPF")
										&& !controltype.equalsIgnoreCase("CalendarEBP")
										&& (!action.equalsIgnoreCase("Read") || ((action.equalsIgnoreCase("Read")
												&& !StringUtils.isEmpty(controlName)
												&& !ctrlValue.equalsIgnoreCase("IGNORE"))))// Minaakshi

										&& !controltype.equalsIgnoreCase("JSScript")
										&& !controltype.equalsIgnoreCase("DB")
										&& !controltype.equalsIgnoreCase("Database")// Minaakshi
										&& !controltype.equalsIgnoreCase("Pagination")// Minaakshi
																						// :
																						// 01-01-2020
										&& !controltype.equalsIgnoreCase("CreateDynamicData")// Minaakshi
																								// :
																								// 03-10-2018
										&& !controltype.equalsIgnoreCase("CreateJsonFile")
										&& !controltype.equalsIgnoreCase("FlatFile")
										&& !controltype.equalsIgnoreCase("FileUpload_DM")// Minaakshi
										&& !controltype.equalsIgnoreCase("CheckBatchStatus")// Minaakshi
																							// 01-05-2019
										&& !controlID.equalsIgnoreCase("XML") && !controltype.startsWith("Process")
										&& !controltype.startsWith("Destroy") && !controltype.startsWith("ReadSikuli")
										&& !controltype.equalsIgnoreCase("WebService") && !action.equalsIgnoreCase("VA")
										&& (ctrlValue == null || !ctrlValue.equalsIgnoreCase("IGNORE"))// Minaakshi
										&& !controltype.equalsIgnoreCase("IFrame") && !action.equalsIgnoreCase("LOOP")
										&& !controltype.equalsIgnoreCase("RowNumbersToExecute")
										&& !controltype.equalsIgnoreCase("WaitForElementToVisible")
										&& !controltype.equalsIgnoreCase("Wait_IfValue")
										&& !action.equalsIgnoreCase("Billing")
										&& !controltype.equalsIgnoreCase("D_Wait")
										&& !action.equalsIgnoreCase("UpdateDBScript")
										&& !action.equalsIgnoreCase("VFormPresence")
										&& !action.equalsIgnoreCase("FormVerificationDetails")

										&& !controltype.equalsIgnoreCase("RenameDocument")
										&& !controltype.equalsIgnoreCase("CopyFromServerLocation")
										&& !controltype.equalsIgnoreCase("IgnoreString")
										&& !controltype.equalsIgnoreCase("PDFDocumentCompare")
										&& !controltype.equalsIgnoreCase("PDFComparisonMode")
										&& !controltype.equalsIgnoreCase("MoveDocument")
										&& !controltype.equalsIgnoreCase("WaitTillFileDownload")
										&& !controltype.equalsIgnoreCase("WindowAlertOk")) {
									if ((indexVal.equalsIgnoreCase("") || indexVal.equalsIgnoreCase("0"))
											&& !controlID.equalsIgnoreCase("TagValue")
											&& !controlID.equalsIgnoreCase("TagText")) {
										webElement = getElementByType(controlID, controlName, WebHelper.control,
												imageType, ctrlValue);

										// webElement =
										// getElementByType(controlName);
									} else {
										controlList = WebHelperUtil.getElementsByType(controlID, controlName,
												WebHelper.control, imageType, ctrlValue);

										if (controlList != null && controlList.size() > 1) {
											webElement = WebHelperUtil.GetControlByIndex(indexVal, controlList,
													controlID, controlName, WebHelper.control, ctrlValue); // ,
											// ISelenium
											// selenium)
										} else {
											break;
										}
									}
								}
							} else {
								sikuliScreen = new Screen();
								// bhaskar
								// sikuliapp = Automation.SikuliScr;
								// bhaskar
							}
						}

						/*** Perform action on the identified control ***/
						// added by sheetal for new control type:
						// XPathValueMultiWithInput & XPathValueMultiWithInput_p

						if (controlID.equalsIgnoreCase("XPathValueMultiWithInput")
								|| controlID.equalsIgnoreCase("XPathValueMultiWithInput_p")) {
							String tempcontrolValue = ctrlValue;
							String[] tempValues2 = tempcontrolValue.split(";;");
							int l = tempValues2.length;
							ctrlValue = tempValues2[l - 1];

							String tempReplaceString = controlName;
							int k = 1;
							for (int j = 0; j < tempValues2.length - 1; j++) {
								tempReplaceString = tempReplaceString.replace("$" + k + "value", tempValues2[j].trim());
								k++;
							}
							controlName = tempReplaceString;
						}

						if (!(action.equals("I") && ctrlValue.equalsIgnoreCase(""))) {
							doAction(imageType, controltype, controlID, controlName, ctrlValue, logicalName, action,
									webElement, true, sheetStructure, headerValues, rowIndex, rowCount, rowNo, colNo);
						}

						// log.info("ctrlValue :"+ctrlValue);
						/*** Perform action on the identified control ***/
					}

					if (action == "END_LOOP" && (valuesRowCount != valuesRowIndex)) {
						WebHelper.loopRow = 1;
						break;
					}

				} else {
					//log.info("ExecuteFlag is N");
					log.info("ExecuteFlag is N");
				}
			}

			// Setting of reporting values after execution in case of no
			// exception
			Date toDate = new Date();
			webDriver.getReport().setFromDate(Config.dtFormat.format(WebHelper.frmDate));
			webDriver.getReport().setIteration(Config.cycleNumber);
			webDriver.getReport().setTestcaseId(controller.controllerTestCaseID.toString());
			webDriver.getReport().setGroupName(controller.controllerGroupName.toString());
			webDriver.getReport().setTrasactionType(controller.controllerTransactionType.toString());
			webDriver.getReport().setTestDescription(controller.testDescription);
			webDriver.getReport().setToDate(Config.dtFormat.format(toDate));

			// Setting status for field verification failures
			if (WebHelper.fieldVerFailCount > 0) {
				webDriver.getReport().setMessage("Check Detailed Results");
				webDriver.getReport().setStatus("FAIL");
			}

			log.info("Completed Transaction : " + controller.controllerTransactionType.toString());// Mandar

		} catch (Exception e) {
			String errorMessage = e.getMessage();
			String firstLine = errorMessage.split("\\r?\\n")[0];
			log.error("Error :" + firstLine);
			//log.error(e.getMessage(), e);
			controller.pauseFun(e.getMessage());
			reportWrittenAfterException = true;
		} finally {
			WebHelper.structureHeader.clear();
			WebHelper.valuesHeader.clear();
			if (!reportWrittenAfterException)
				ExcelUtility.writeReportPAS(webDriver.getReport());
			WebHelper.fieldVerFailCount = 0;
		}
	}

	public static Reporter WriteToDetailResults(String expectedValue, String actualValue, String columnName)
			throws IOException

	{

		// below condition changed to take the value from config sheet
		// if (WebHelper.file.exists() == true && DResult == 1) {
		if (WebHelper.file.exists() == true
				&& StringUtils.equalsIgnoreCase(Config.appendVerificationResultPath, "false") && DResult == 1) {

			// print = new PrintStream(file);
			WebHelper.file.delete();
		}
		// File file= new
		// File(Automation.getConfigValue("VERIFICATIONRESULTSPATH").toString());
		Reporter report = new Reporter();
		report.setReport(report);
		report = report.getReport();
		String passCount = "";
		String failCount = "";
		report.setTestcaseId(controller.controllerTestCaseID.toString());
		report.setTrasactionType(transactionType.toString());
		report.setTestDescription(controller.testDescription);

		if (ITAFWebDriver.isDMApplication()
				|| (ITAFWebDriver.isSuiteApplication() && MainControllerSuite.transactionSilo.equalsIgnoreCase("DM")))// Minaakshi
																														// :
																														// 01-11-2019
			report.setGroupName(controller.controllerGroupName.toString());

		// Quote ID for PAS report requirement
		if (controller.controllerQuoteId != null) {
			report.setStrQuoteId(controller.controllerQuoteId.toString());
		} else {
			report.setStrQuoteId("");
		}

		// Minaakshi : 01-10-2020
		if ((columnName.toString().contains("SSN#") && (controller.controllerTestCaseID.toString().contains("SMK_")))
				|| (columnName.toString().contains("FEIN#"))
						&& (controller.controllerTestCaseID.toString().contains("SMK_"))) {
			String replacedtempExpecetedValue = expectedValue.replace("-", "");
			expectedValue = replacedtempExpecetedValue;
		}

		// Trim Actual and Expected Value : Minaakshi : 01-10-2020
		if (!StringUtils.isBlank(expectedValue)) {
			expectedValue = expectedValue.trim();
		}

		if (!StringUtils.isBlank(actualValue)) {
			actualValue = actualValue.trim();
		}

		if ((expectedValue).equalsIgnoreCase(actualValue)) {// Minaakshi :
															// 01-01-2020
			report.setActualValue(actualValue);
			report.setExpectedValue(expectedValue);
			report.setStatus("PASS");
			report.setToDate(Config.dtFormat.format(WebHelper.frmDate));
			passCount = "1";
			failCount = "0";
			TotalpassCount += 1;
			log.info(TotalpassCount);
		} else {
			report.setActualValue("FAIL|" + actualValue + "|" + expectedValue);
			log.info("ab");
			report.setExpectedValue(expectedValue);
			report.setStatus("FAIL");
			report.setToDate(Config.dtFormat.format(WebHelper.frmDate));
			failCount = "1";
			passCount = "0";
			// DS:30-05-2014
			WebHelper.fieldVerFailCount += 1;
			TotalfailCount += 1;
			log.info(TotalfailCount);
		}

		WebHelper.print = new PrintStream(new FileOutputStream(WebHelper.file, true));

		int usedRows = WebHelperUtil.count(WebHelper.file);
		if (usedRows == 0) {

			// WebHelper.print.print("Iteration,TestCaseID,TransactionType,CurrentDate,RowType,Status,PassCount,FailCount,
			// Compare Result,Quote Number");

			WebHelper.print.print(
					"GroupName,TestCaseID,TransactionType,CurrentDate,RowType,Status,PassCount,FailCount, Compare Result,Quote Number");// Minaakshi
			// :
			// 01-11-2019

			WebHelper.print.println();
		}
		usedRows = WebHelperUtil.count(WebHelper.file);


		// WebHelper.print.print(ExcelUtility.myChar + Config.cycleNumber +
		// ExcelUtility.myChar + "," + ExcelUtility.myChar +
		// report.getTestcaseId()
		WebHelper.print.print(ExcelUtility.myChar + report.getGroupName() + ExcelUtility.myChar + ","
				+ ExcelUtility.myChar + report.getTestcaseId()

				+ ExcelUtility.myChar + "," + ExcelUtility.myChar + report.getTrasactionType() + ExcelUtility.myChar
				+ "," + ExcelUtility.myChar + report.getToDate() + ExcelUtility.myChar + "," + ExcelUtility.myChar
				+ "Field: " + columnName + ExcelUtility.myChar + "," + ExcelUtility.myChar + report.getStatus()
				+ ExcelUtility.myChar + "," + ExcelUtility.myChar + passCount + ExcelUtility.myChar + ","

				+ ExcelUtility.myChar + failCount + ExcelUtility.myChar + "," + ExcelUtility.myChar
				+ report.getActualValue() + ExcelUtility.myChar + "," + ExcelUtility.myChar + report.getStrQuoteId()
				+ ExcelUtility.myChar);
		WebHelper.print.println();
		DResult++;
		return report;
	}

	/** Locating Web Element **/
	public static WebElement getElementByType(String controlId, String controlName, String controlType,
			String imageType, String controlValue) throws Exception {
		// String a=controlValue;

		// WebDriverWait wait = new WebDriverWait(Automation.driver,
		// Integer.parseInt(Config.timeOut));//Sel4
		WebDriverWait wait = new WebDriverWait(Automation.driver, Duration.ofSeconds(Integer.parseInt(Config.timeOut)));

		WebHelperDM.implementWait();

		WebElement controlList = null;
		Constants.ControlIdEnum controlID = Constants.ControlIdEnum.valueOf(controlId);
		WebDriverWait wait2 = new WebDriverWait(Automation.driver, (Duration.ofSeconds(2)));

		if (StringUtils.contains(controlValue, "pickFromUniqueNumbers")) {
			String dynamicValue = "";
			dynamicValue = WebHelperUtil.ReadFromExcel(controlValue, WebHelper.columnName);
			if (!dynamicValue.equalsIgnoreCase("")) {
				controlValue = controlValue.replaceAll("pickFromUniqueNumbers", dynamicValue);
			}
		}

		// log.info("bhaski controlID:"+controlID);
		try {
			switch (controlID) {
			case doNothing:
				break;
			case PageToLoad:

			case Id:

				// WebHelper.getText(Automation.driver,Automation.driver.findElement(By.xpath("")));
			case HTMLID:
				// log.info(Automation.driver.getCurrentUrl());

				controlList = wait.until(ExpectedConditions.elementToBeClickable(By.id(controlName)));
				break;

			case XPath_ctrvalue:

			case XPath:

				// JavascriptExecutor js =
				// (JavascriptExecutor)Automation.driver;
				if (isDMApplication || (ITAFWebDriver.isSuiteApplication()
						&& MainControllerSuite.transactionSilo.equalsIgnoreCase("DM"))) {// Minaakshi
					// :
					// 01-10-2020
					Thread.sleep(250);
					controlList = dmgebtWait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
				} else {
					Thread.sleep(1000);
					controlList = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
				}
				break;

			case XPathValue: // Minaakshi : 01-10-2020
				String tempCtrlName = controlName;
				String tempReplaceString = "";

				if (controlValue.contains("|")) {// Minaakshi : 01-05-2019
					String[] temp = controlValue.split("\\|");
					tempReplaceString = tempCtrlName.replace("$value", temp[0].toString());
				} else if (controlValue.contains(";")) {
					String[] temp = controlValue.split("\\;");
					tempReplaceString = tempCtrlName.replace("$value", temp[0].toString());
				} else {
					tempReplaceString = tempCtrlName.replace("$value", controlValue);
				}

				if (isDMApplication || (ITAFWebDriver.isSuiteApplication()
						&& MainControllerSuite.transactionSilo.equalsIgnoreCase("DM"))) {

					controlList = dmgebtWait
							.until(ExpectedConditions.presenceOfElementLocated(By.xpath(tempReplaceString)));
				} else {

					controlList = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString)));

				}

				break;

			case XPathValue_dy: // Minaakshi : 01-10-2020
				String tempCtrlValue = "";
				String tempReplaceString5 = "";

				if (controlValue.contains("|")) {
					String[] temp = controlValue.split("\\|");
					tempCtrlValue = DMProduct.ReadFromExcelUsingColumnName(temp[0].toString(), WebHelper.columnName);
				} else {
					tempCtrlValue = DMProduct.ReadFromExcelUsingColumnName(controlValue, WebHelper.columnName);
				}

				String tempCtrlName5 = controlName;
				tempReplaceString5 = tempCtrlName5.replace("$value", tempCtrlValue);

				controlList = dmgebtWait
						.until(ExpectedConditions.presenceOfElementLocated(By.xpath(tempReplaceString5)));

				break;
			// Minaakshi : 01-03-2021 : $value1;@value2
			case XPathValue_Multi:
				String tempcontrolValue1 = controlValue;
				String[] tempValues1 = tempcontrolValue1.split("//;");
				String tempReplaceString7 = controlName;
				int a = 1;
				for (String temp : tempValues1) {
					tempReplaceString7 = tempReplaceString7.replace("$" + "value" + a, temp.trim());
					a++;
				}

				controlList = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString7)));
				break;

			case XPathValue_cpp: // Subhasis : 27-11-2018
				// case added for PAS requirement from Sandesh Kumbhar
				try {
					String tempCtrlNamecbp = controlName;
					String tempReplaceStringcbp = tempCtrlNamecbp.replace("$value", controlValue);
					controlList = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceStringcbp)));
				} catch (Exception e) {
					log.info(e.getMessage());
					return controlList;
				}
				break;

			case XPath_VFCount: // Sheetal:2/13/2019 - for form count
				// verification in Policy2015
				try {
					String tempCtrlName2 = controlName;
					int tempControlValue = Integer.parseInt(controlValue);
					tempControlValue = tempControlValue + 4;
					String extraFormLocator = Integer.toString(tempControlValue);

					String tempReplaceString2 = tempCtrlName2.replace("$value", extraFormLocator);
					controlList = Automation.driver.findElement(By.xpath(tempReplaceString2));
				} catch (Exception e) {
					controlList = null;
				}
				break;

			case RetriveFormsCountAndName:
				missingForms.removeAll(missingForms);
				additionalForms.removeAll(additionalForms);
				presentForms.removeAll(presentForms);
				expectedForms.removeAll(expectedForms);
				formList.removeAll(formList);

				List<WebElement> forms = Automation.driver.findElements(By.xpath(controlName));

				for (int i = 0; i < forms.size(); i++) {
					formList.add(forms.get(i).getText());
				}
				break;

			case XPathValue_p: // Mrinmayee : 15-11-2018
				try {
					// WebDriverWait wait2 = new
					// WebDriverWait(Automation.driver
					// ,5);
					String tempCtrlName1 = controlName;
					String tempReplaceString1 = tempCtrlName1.replace("$value", controlValue);
					controlList = wait2
							.until(ExpectedConditions.presenceOfElementLocated(By.xpath(tempReplaceString1)));

				} catch (Exception e) {

					controlList = null;
				}

				break;
			// Sheetal:2/19/2019
			case XPathValueMulti:
				String tempcontrolValue = controlValue;
				String[] tempValues = tempcontrolValue.split(";;");
				String tempReplaceString3 = controlName;
				int i = 1;
				for (String temp : tempValues) {
					tempReplaceString3 = tempReplaceString3.replace("$" + i + "value", temp.trim());
					i++;
				}

				controlList = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString3)));
				break;

			case XPathValueMultiWithInput: // Mrinmayee 12-12-2018
				String tempcontrolValue2 = controlValue;
				String[] tempValues2 = tempcontrolValue2.split(";;");
				String tempReplaceString4 = controlName;
				int k = 1;
				for (int j = 0; j < tempValues2.length - 1; j++) {
					tempReplaceString4 = tempReplaceString4.replace("$" + k + "value", tempValues2[j].trim());
					k++;
				}
				controlList = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString4)));
				break;

			case XPathValueMultiWithInput_p:
				try {
					String tempcontrolValue3 = controlValue;
					String[] tempValues3 = tempcontrolValue3.split(";;");
					String tempReplaceString6 = controlName;
					int s = 1;
					for (int j = 0; j < tempValues3.length - 1; j++) {
						tempReplaceString6 = tempReplaceString6.replace("$" + s + "value", tempValues3[j].trim());
						s++;
					}
					controlList = wait2
							.until(ExpectedConditions.presenceOfElementLocated(By.xpath(tempReplaceString6)));
				}

				catch (Exception e) {
					controlList = null;
				}
				break;

			case XPathValueMulti_p:
				String tempcontrolVal = controlValue;
				String[] tempVal = tempcontrolVal.split(";;");
				String tempReplaceStr = controlName;
				int l = 1;
				for (String temp : tempVal) {
					tempReplaceStr = tempReplaceStr.replace("$" + l + "value", temp.trim());
					l++;
				}
				try {
					controlList = wait2.until(ExpectedConditions.presenceOfElementLocated(By.xpath(tempReplaceStr)));
				} catch (Exception e) {

					controlList = null;
				}
				break;

			case XPath_dy:// Minaakshi : 01-10-2020 Added this case while
							// handling
				// xpath which requires ctrlvalue from unique number
				String UNSheetValue = "";
				if (controlValue.equalsIgnoreCase("Yes")) {// Minaakshi :
					// 6-12-2018
					UNSheetValue = DMProduct.ReadFromExcelUsingColumnName("", WebHelper.columnName);
				} else {
					UNSheetValue = DMProduct.ReadFromExcelUsingColumnName(controlValue, WebHelper.columnName);

				}

				String tempCtrlName1 = controlName;
				String tempReplaceString1 = tempCtrlName1.replace("$value", UNSheetValue);

				controlList = dmgebtWait
						.until(ExpectedConditions.presenceOfElementLocated(By.xpath(tempReplaceString1)));

				break;
			// Added by sheetal
			case XPath_Wait:
				WebDriverWait wait3 = new WebDriverWait(Automation.driver, (Duration.ofSeconds(5)));
				int p, max = 600;
				for (p = 1; p < max; p += 10) {
					try {
						controlList = wait3.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
						log.info("Found element after waiting for approx: " + p + " seconds.");
						break;
					} catch (Exception e) {
						Thread.sleep(10000);
					}
				}
				if (p >= max) {
					log.info("Time-out occured after wating for approx." + max + " seconds. Element: "
							+ controlName + " is still not loaded.");
					controlList = null;
				}
				break;

			case Name: // Minaakshi : 01-10-2020
				if (isDMApplication || (ITAFWebDriver.isSuiteApplication()
						&& MainControllerSuite.transactionSilo.equalsIgnoreCase("DM"))) {
					controlList = dmgebtWait.until(ExpectedConditions.presenceOfElementLocated(By.name(controlName)));
				} else {
					controlList = wait.until(ExpectedConditions.elementToBeClickable(By.name(controlName)));
				}

				break;

			case ClassName:
				controlList = wait.until(ExpectedConditions.elementToBeClickable(By.className(controlName)));
				break;

			case LinkText:
				// js.executeScript("return document.readyState").toString().equals("complete");
				controlList = wait.until(ExpectedConditions.elementToBeClickable(By.linkText(controlName)));
				break;

			case LinkValue:
				controlList = wait.until(ExpectedConditions.elementToBeClickable(By.linkText(controlValue)));
				break;

			case TagText:
			case TagValue:
			case TagOuterText:
				controlList = wait.until(ExpectedConditions.elementToBeClickable(By.tagName(imageType)));
				break;

			case CSSSelector:
				controlList = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(controlName)));
				break;

			// GAIC AJAX controls - TM:02/02/2015

			case AjaxPath: // Minaakshi : 01-10-2020
				if (isDMApplication || (ITAFWebDriver.isSuiteApplication()
						&& MainControllerSuite.transactionSilo.equalsIgnoreCase("DM"))) {
					Thread.sleep(500);// Minaakshi : 01-06-2020
					controlList = dmgebtWait.until(ExpectedConditions
							.elementToBeClickable(By.xpath(controlName + "[contains(text(),'" + controlValue + "')]")));
				} else {

					controlList = wait.until(ExpectedConditions
							.elementToBeClickable(By.xpath(controlName + "[contains(text(),'" + controlValue + "')]")));
				}
				break;

			case AjaxPath_ExactValue:// Minaakshi : 01-10-2020 : Added this
				// case
				// for selecting Exact value from
				// Drop-Down
				// controlList =
				if (isDMApplication || (ITAFWebDriver.isSuiteApplication()
						&& MainControllerSuite.transactionSilo.equalsIgnoreCase("DM"))) {
					Thread.sleep(500);
					controlList = dmgebtWait.until(ExpectedConditions
							.elementToBeClickable(By.xpath(controlName + "[text()='" + controlValue + "']")));
				} else {

					controlList = wait.until(ExpectedConditions
							.elementToBeClickable(By.xpath(controlName + "[contains(text(),'" + controlValue + "')]")));

				}

				break;

			case AjaxPath_dy:// Minaakshi : 01-10-2020 Added this case while
								// handling
				// dyanamic xpath which requires ctrlvalue from
				// unique number sheet
				String UNSheetValue2 = "";
				if (controlValue.equalsIgnoreCase("Yes")) {
					UNSheetValue2 = DMProduct.ReadFromExcelUsingColumnName("", WebHelper.columnName);
				} else {
					UNSheetValue2 = DMProduct.ReadFromExcelUsingColumnName(controlValue, WebHelper.columnName);
				}

				// String temp2 = controlName + "[contains(text(),'" + UNSheetValue2 +
				// "')]";//Sel4

				Thread.sleep(500);
				controlList = dmgebtWait
						.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName + "[contains(text(),'"

								+ UNSheetValue2 + "')]")));

				break;

			case Id_p:
			case HTMLID_p:
				controlList = wait.until(ExpectedConditions.presenceOfElementLocated(By.id(controlName)));
				break;

			case XPath_H:

				/*
				 * boolean controlList1 = wait
				 * .until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(controlName))
				 * );
				 */// Sel4
				break;

			case XPath_p: // Minaakshi : 01-10-2020
				if (isDMApplication || (ITAFWebDriver.isSuiteApplication()
						&& MainControllerSuite.transactionSilo.equalsIgnoreCase("DM"))) {

					controlList = dmgebtWait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
				} else {

					controlList = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

				}

				break;

			case Xpath_ctrlvalue: // Minaakshi : 01-10-2020
				try {

					if (!controlValue.isEmpty()) {

						controlName = controlName.replace("$value", controlValue);

						controlList = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));

					} else {
						controlList = null;
					}
				} catch (Exception e) {
					log.error(e.getMessage(), e);
					controlList = null;
				}

				break;
			case XPath_R:

				// controlList =
				// wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlValue)));

				try {

					if (!controlValue.isEmpty()) {

						controlList = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlValue)));

					} else {
						controlList = null;
					}
				} catch (Exception e) {
					log.error(e.getMessage(), e);
					controlList = null;
				}

				break;

			case XPath_value:

				if (!controlValue.isEmpty()) {

					try {

						if (Automation.driver.findElement(By.xpath(controlName)).isDisplayed()) {

							List<WebElement> webElement = Automation.driver.findElements(By.xpath(controlName));

							for (WebElement element : webElement) {

								log.info(element.getAttribute("value"));

								if (element.getAttribute("value").equalsIgnoreCase(controlValue)) {

									String Id = element.getAttribute("id");

									String elementTagName = element.getTagName();

									String controlName1 = "//" + elementTagName + "[@id='" + Id + "'" + "]/"
											+ imageType;

									controlName = controlName1;

									controlList = Automation.driver.findElement(By.xpath(controlName));

									break;
								}

								else {

									log.info("Element value doesn't match with control value");
								}

							}

							controlList = wait2.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
							// controlList=fluentWait(By.xpath(controlName));
						}

					} catch (Exception e) {
						log.error(e.getMessage(), e);

						controlList = null;
					}
				} else {

					System.out.print("don't perform any action");
				}
				break;

			// Mrinmayee - Wait till visibility of element
			case XPath_vis:

				Thread.sleep(1000);
				controlList = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(controlName)));

				break;
			// Mrinmayee - Wait till visibility of element
			case XPath_if:

				try {

					// WebDriverWait wait2 = new
					// WebDriverWait(Automation.driver,5);
					// controlList =
					// wait2.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
					if (Automation.driver.findElement(By.xpath(controlName)).isDisplayed()) {
						controlList = Automation.driver.findElement(By.xpath(controlName));
					}
				} catch (Exception e) {
					// log.error(e.getMessage(), e);
					controlList = null;
				}
				break;
			case XPath_Rif:// Sel4
				// Code for handling XPath_Rif
				break;

			case XPath_Menu:// Sel4
				// Code for handling XPath_Rif
				break;

			case XPath_Editif:// Sel4
				// Code for handling XPath_Rif
				break;
			case FluentWait:// Sel4
				break;
			case TagName:// Sel4
				break;
			case XPathValueEval:// Sel4
				break;
			default:
				break;

			}
			return controlList;
		} catch (Exception e) {
			// log.info("bhaskar in catch block");
			log.error(e.getMessage(), e);
			throw new Exception(e.getMessage());
		}
	}

	@SuppressWarnings("incomplete-switch")
	public static String doAction(String imageType, String controlType, String controlId, String controlName,
			String ctrlValue, String logicalName, String action, WebElement webElement, Boolean Results,
			Sheet strucSheet, Sheet valSheet, int rowIndex, int rowcount, String rowNo, String colNo) throws Exception {

		// new
		// WebDriverWait(Automation.driver,timeOutInSeconds).until(ExpectedConditions.invisibilityOfElementLocated(By.id("divProgress")));

		List<WebElement> WebElementList = null;
		String currentValue = null;
		String uniqueNumber = "";
		WebVerification.isFromVerification = false;
		Constants.ControlTypeEnum controlTypeEnum = Constants.ControlTypeEnum.valueOf(controlType);
		Constants.ControlTypeEnum actionName = Constants.ControlTypeEnum.valueOf(action);
		// bhaskar
		WebHelper.sikscreen = Config.SikuliScr;
		// log.info(sikscreen);
		// bhaskar
		if (controlType.contains("Robot") && !WebHelper.isIntialized) {
			log.info("In method doaction debug1");
			WebHelper.robot = new Robot();
			WebHelper.isIntialized = true;
		}

		log.info("ctrlValue doAction:" + ctrlValue);
		// - skip the step for which control value is blank : start

		JavascriptExecutor js = (JavascriptExecutor) Automation.driver;
		try {
			// getHTMLResponse();
			// log.info("In method doaction debug3");
			switch (controlTypeEnum) {

			case WebEdit:
				switch (actionName) {
				case NC:

					webElement = getElementByType(controlId, controlName, WebHelper.control, imageType, ctrlValue);

					try {
						if (webElement.isEnabled()) {
							Thread.sleep(4000);
							// webElement.sendKeys(Keys.INSERT(ctrlValue));
							webElement.click();

							// webElement.sendKeys(Keys.chord(Keys.CONTROL,"a"));

							if (webElement.getAttribute("value") != null) {
								webElement.clear();
								// webElement.sendKeys(Keys.TAB);
								Thread.sleep(1000);
							}

							webElement.sendKeys(ctrlValue);
							// webElement.sendKeys(Keys.TAB); Mrinmayee
							Thread.sleep(2000);
							if (webElement.getAttribute("value").isEmpty()) {
								webElement.sendKeys(ctrlValue);

							}
							// Automation.driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
							Automation.driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));// Sel4
							// Automation.driver.manage().timeouts().pageLoadTimeout(50, TimeUnit.SECONDS);
							Automation.driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(50));// Sel4

							Thread.sleep(2000);
						} else {
							log.info("Element Not Found");
						}
						;

					} catch (StaleElementReferenceException e) {
						log.error(e.getMessage(), e);
						log.info("Trying to recover from a stale element :" + e.getMessage());

					}

					break;
				case RD:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					// Thread.sleep(500);
					if (ctrlValue.contains("|")) {// Minaakshi :
													// 01-05-2019
						String[] temp = ctrlValue.split("\\|");
						ctrlValue = temp[1].toString();
					}

					String ValueStatus = webElement.getAttribute("value");
					if (ValueStatus.isEmpty()) {
						if (Config.projectName.equals("DistributionManagement")) {// Minaakshi
							if (!ctrlValue.equalsIgnoreCase("null")) {
								//log.info("!@#$%^&*!@#$%^&*ctrlValue:" + ctrlValue); Sel4
								webElement.clear();
								webElement.sendKeys(ctrlValue);

							} else {
								webElement.clear();
							}
						} else {
							if (ctrlValue != null) {
								// waitForPageLoadingToComplete();
								// waitForAjaxLoad(Automation.driver);
								try {
									// sheetal 2-7-2019
									((JavascriptExecutor) Automation.driver)
											.executeScript("arguments[0].scrollIntoView();", webElement);
									Thread.sleep(500);

									if (webElement.isEnabled()) {
										WebHelperUtil.InputValue(webElement, ctrlValue, controlId, controlName,
												imageType, ctrlValue);
									}

								} catch (Exception e) {
									log.error(e.getMessage(), e);
									log.info("Element not displayed");
								}
							}
						}
					} else {
						break;
					}
					break;// Minaakshi : 17-12-2018

				case JSEdit:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					// Thread.sleep(500);
					if (ctrlValue.contains("|")) {// Minaakshi :
													// 01-05-2019
						String[] temp = ctrlValue.split("\\|");
						ctrlValue = temp[1].toString();
					}

					String ValueStatus1 = webElement.getAttribute("value");
					if (ValueStatus1.isEmpty()) {
						if (Config.projectName.equals("DistributionManagement")) {// Minaakshi
							if (!ctrlValue.equalsIgnoreCase("null")) {
								//log.info("!@#$%^&*!@#$%^&*ctrlValue:" + ctrlValue);
								webElement.clear();
								webElement.sendKeys(ctrlValue);

							} else {
								webElement.clear();
							}
						} else {
							if (ctrlValue != null) {
								// waitForPageLoadingToComplete();
								// waitForAjaxLoad(Automation.driver);
								try {
									// sheetal 2-7-2019
									((JavascriptExecutor) Automation.driver)
											.executeScript("arguments[0].value=ctrlValue;", webElement);
									Thread.sleep(500);

									if (webElement.isEnabled()) {
										WebHelperUtil.InputValue(webElement, ctrlValue, controlId, controlName,
												imageType, ctrlValue);
									}

								} catch (Exception e) {
									log.error(e.getMessage(), e);
									log.info("Element not displayed");
								}
							}
						}
					} else {
						break;
					}
					break;// Minaakshi : 17-12-2018

				case Read:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					uniqueNumber = DMProduct.ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
					webElement.clear();
					webElement.sendKeys(uniqueNumber);
					break;

				case Write:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo,
							colNo);

					break;

				case Input: // Replace fields InputifExist and I by
					// Asif
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}

					try {
						if (webElement.isEnabled()) {

							if (!ctrlValue.equalsIgnoreCase("null")) {
								{
									// log.info("!@#$%^&*!@#$%^&*ctrlValue:"+ctrlValue);
									for (int i = 0; i < 5; i++) {
										webElement.sendKeys(Keys.ENTER);

										webElement.clear();

										Thread.sleep(500);
										webElement.sendKeys(ctrlValue);
										// webElement.click();
										log.info(webElement.getAttribute("value"));
										if (!webElement.getAttribute("value").isEmpty()) {
											webElement.sendKeys(Keys.TAB);
											break;
										} else {
											Thread.sleep(3000);
										}
									}
								}
							} else {
								break;
							}
						}
					} catch (Exception e) {
						log.error(e.getMessage(), e);
						log.info("Element not exist");
					}
					break;

				case WaitTodisplayElementValue:

					WebHelperUtil.WaitUntilAttributeValueEquals(webElement);

					break;

				case V:
					if (ctrlValue == null || ctrlValue.trim().equals("") || ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}

					// currentValue = webElement.getText();
					try {
						currentValue = webElement.getAttribute("value");
						// currentValue = webElement.getText();
						log.info(currentValue);
						if (currentValue.equalsIgnoreCase(ctrlValue)) {

							log.info("PASSED");
						} else {
							log.info("FAILED");
						}
					} catch (Exception e) {
						log.error(e.getMessage(), e);
						currentValue = "null";
						if (currentValue.equalsIgnoreCase(ctrlValue)) {

							log.info("PASSED");
						} else {
							log.info("FAILED");
						}
					}
					break;

				case Clear:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					// webElement.clear();
					try {
						if (webElement.isDisplayed()) {

							Coordinates coordinate = ((Locatable) webElement).getCoordinates();
							coordinate.onPage();
							coordinate.inViewPort();
							Thread.sleep(1000);
							webElement.clear();
						}
					} catch (Exception e) {
						log.info("Element does not exist");

					}
					break;
				// js1.executeScript("arguments[0].value = '';",
				// webElement);

				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					// Thread.sleep(500);
					if (ctrlValue.contains("|")) {// Minaakshi :
						// 01-05-2019
						String[] temp = ctrlValue.split("\\|");
						ctrlValue = temp[1].toString();
					}
					if (Config.projectName.equals("DistributionManagement")) {// Minaakshi
						if (!ctrlValue.equalsIgnoreCase("null")) {
							//log.info("!@#$%^&*!@#$%^&*ctrlValue:" + ctrlValue);
							webElement.clear();
							webElement.sendKeys(ctrlValue);

						} else {
							webElement.clear();
						}
					} else {
						if (ctrlValue != null) {
							// waitForPageLoadingToComplete();
							// waitForAjaxLoad(Automation.driver);
							try {
								// sheetal 2-7-2019
								((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
										webElement);

								Thread.sleep(500);

								if (webElement.isEnabled()) {
									if (ctrlValue.equalsIgnoreCase("HitTabKey")) {
										webElement.sendKeys(Keys.TAB);
									} else {
										WebHelperUtil.InputValue(webElement, ctrlValue, controlId, controlName,
												imageType, ctrlValue);
									}
								}

							} catch (Exception e) {
								log.error(e.getMessage(), e);
								log.info("Element not displayed");
							}
						}
					}
					break;// Minaakshi : 17-12-2018

				}
				break;
			case EonPagination:
				switch (actionName) {

				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					for (int i = 0; i < 2; i++) {
						if (ctrlValue.equalsIgnoreCase("Notdisplayed")) {
							break;
						}
						try {

							List<WebElement> pagination = Automation.driver
									.findElements(By.xpath("//ul[@class='pagination ice-pagination']/li/a"));
							int k = 0;
							while (pagination.size() - 1 > 0) {

								if (webElement.isDisplayed()) {
									log.info("inside if loop 1");
									webElement = getElementByType(controlId, controlName, WebHelper.control, imageType,
											ctrlValue);
									if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
										if (!webElement.isSelected()) {
											Coordinates coordinate = ((Locatable) webElement).getCoordinates();
											coordinate.onPage();
											coordinate.inViewPort();
											JavascriptExecutor executor = (JavascriptExecutor) Automation.driver;
											executor.executeScript("arguments[0].click();", webElement);
											if (webElement.isSelected()) {
												// Automation.driver.manage().timeouts().pageLoadTimeout(60,TimeUnit.SECONDS);
												break;
											}
											log.info("inside if loop 2");
										} else {
											log.info("inside if loop 3");
											break;
										}
									} else {
										log.info("inside if loop 4");
										Thread.sleep(500);
										if (!webElement.isSelected()) {

											break;
										} else {
											JavascriptExecutor executor = (JavascriptExecutor) Automation.driver;
											executor.executeScript("arguments[0].click();", webElement);

											break;
										}
									}
								} else {

									// List<WebElement> pagination
									// =Automation.driver.findElements(By.xpath("//ul[@class='pagination
									// ice-pagination']/li/a"));
									// checkif pagination link exists

									k++;
									String s = String.valueOf(k);
									// click on pagination link
									String newxpath = "//ul[@class='pagination ice-pagination']/li/a" + "[" + s + "]";
									log.info("newxpath" + newxpath);
									Automation.driver.findElements(By.xpath(newxpath));

								}
							}

						}

						catch (Exception e) {
							log.error(e.getMessage(), e);
							try {
								Thread.sleep(500);
								webElement = getElementByType(controlId, controlName, WebHelper.control, imageType,
										ctrlValue);
								// webElement.click();
								break;
							} catch (Exception e1) {

								Thread.sleep(2000);
								log.info("Checkbox not found");
								i--;
							}

						}

					}

				}
				break;
			case ReplaceDefault:
				switch (actionName) {
				case I:
					if (ctrlValue.equalsIgnoreCase("null") || ctrlValue.equalsIgnoreCase("")) {
						break;
					}

					((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
							webElement);

					Thread.sleep(500);

					try {
						int counter = 1;
						String temp = "";
						while (!temp.equals(ctrlValue)) {

							JavascriptExecutor js1 = (JavascriptExecutor) Automation.driver;
							js1.executeScript("arguments[0].value = '';", webElement);
							webElement.sendKeys(ctrlValue);
							Thread.sleep(500);
							webElement.sendKeys(Keys.TAB);
							webElement.sendKeys(Keys.TAB);
							Thread.sleep(1000);
							temp = webElement.getAttribute("value");
							if (temp.equals("")) {
								temp = webElement.getText();
							}
							/*
							 * temp = temp.replace("$", ""); temp = temp.replace(" ", ""); temp =
							 * temp.replace(",", "");
							 */
							counter++;
							if (counter == 5) {
								log.info("Could not set correct value");
								break;
							}
						}

					} catch (StaleElementReferenceException e) {
						e.toString();
						log.info("Trying to recover from a stale element :" + e.getMessage());
					}

					break;

				}
				break;

			case WaitForElementToBeFound:
				if (webElement == null) {
					log.info("Element not found.");
				} else {
					log.info("Element found.");
				}
				break;
			case WebButton:
				switch (actionName) {
				case FileUpload:
					// Added code to handle Fileupload window using
					// AutoIt
					String autoitFileDir = Config.inputDataFilePath + TransactionMapping.directoryPathFileUpload;
					webElement.click();
					Thread.sleep(5000);
					// Runtime.getRuntime().exec(autoitFileDir +
					// "\\FileUpload.exe "+ctrlValue);
					Runtime.getRuntime().exec(autoitFileDir + "\\FileUpload.exe " + autoitFileDir + "\\" + ctrlValue);
					Thread.sleep(5000);
					break;

				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}

					if (!StringUtils.equalsIgnoreCase(ctrlValue, null) || !ctrlValue.equalsIgnoreCase("No")) {

						// waitForPageLoadingToComplete();
						// waitForAjaxLoad(Automation.driver);
						int i = 7;
						while (i >= 1) {
							try {

								webElement = getElementByType(controlId, controlName, WebHelper.control, imageType,
										ctrlValue);

								/*
								 * if(webElement==null){
								 * 
								 * break; }
								 */
								if (webElement.isEnabled() == true) {

									Thread.sleep(1000);
									// webElement.click();
									// Mrinmayee - Webbutton click with
									// conditions
									Actions MousebuilderClick = new Actions(Automation.driver);
									Action MouseclickAction = MousebuilderClick.moveToElement(webElement).clickAndHold()
											.release().build();

									MouseclickAction.perform();

									Thread.sleep(1000);
									break;
								}
							} catch (Exception e) {
								log.error(e.getMessage(), e);
								try {
									webElement = getElementByType(controlId, controlName, WebHelper.control, imageType,
											ctrlValue);

									Thread.sleep(500);
									if (webElement.isEnabled()) {
										if (i <= 5) {
											// js.executeScript("arguments[0].focus();
											// arguments[0].blur(); return
											// true",

											// webElement);
											Actions MousebuilderClick = new Actions(Automation.driver);
											Action MouseclickAction = MousebuilderClick.moveToElement(webElement)
													.clickAndHold().release().build();
											MouseclickAction.perform();
											break;
										} else {
											webElement.click();
											break;
										}

										// webElement.sendKeys(Keys.ENTER);
									}
									Thread.sleep(1000);

								} catch (Exception e1) {
									log.error(e.getMessage(), e);
									Thread.sleep(3000);
									i--;
									log.info("Element does not exist");
									// e1.printStackTrace();
								}
							}
						}
					} else {
						log.info("Don't Perform any action");
					}
					break;

				case NC:

					// waitForAjaxLoad(Automation.driver);
					// waitForPageLoadingToComplete();
					int iCtr = 7;
					while (iCtr >= 1) {

						webElement = getElementByType(controlId, controlName, WebHelper.control, imageType, ctrlValue);

						try {

							if (webElement.isEnabled()) {

								Thread.sleep(500);

								webElement.click();
								Thread.sleep(500);

							}
							break;
						}

						catch (Exception e) {
							log.error(e.getMessage());
							try {
								Thread.sleep(500);
								webElement = getElementByType(controlId, controlName, WebHelper.control, imageType,
										ctrlValue);

								if (webElement.isEnabled()) {

									if (iCtr <= 6) {
										// js.executeScript("arguments[0].focus(); arguments[0].blur(); return true",
										// webElement);
										Actions MousebuilderClick = new Actions(Automation.driver);
										Action MouseclickAction = MousebuilderClick.moveToElement(webElement)
												.clickAndHold().release().build();
										MouseclickAction.perform();
										log.info("catchclicked");
										log.info("catchclicked");
										Thread.sleep(500);
										break;
									} else {
										// js.executeScript("arguments[0].focus(); arguments[0].blur(); return true",
										// webElement);
										webElement.click();
										break;
									}
								}

							} catch (Exception e1) {
								// webElement =
								// getElementByType(controlId,
								// controlName,control,imageType,ctrlValue);
								if (iCtr <= 2) {
									log.error(e1.getMessage(), e1);
								} else {
									log.error(e1.getMessage());
								}
								Thread.sleep(1000 * iCtr);

								iCtr--;
								log.info("Element does not exist");
								// e1.printStackTrace();

							}
						}
					}
					Thread.sleep(1000);
					break;

				case V:
					if (ctrlValue == null || ctrlValue.trim().equals("") || ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					if (webElement.isDisplayed()) {
						if (webElement.isEnabled() == true)
							currentValue = "True";
						else
							currentValue = "False";
					}

				case ActionClick:

					// Minaakshi : 05-02-2019
					if (ctrlValue == null || ctrlValue.trim().equals("") || ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					Actions builderClick = new Actions(Automation.driver);
					Action clickAction = builderClick.moveToElement(webElement).clickAndHold().release().build();
					clickAction.perform();

					break;
				// Mrinmayee - Action click with conditions end
				case QSP:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					try {
						Automation.driver.switchTo().frame(Automation.driver.findElement(By.xpath(controlName)));
					} catch (Exception e) {
						log.error(e.getMessage(), e);
						System.out.print("Quote search page not present");
					}
					break;

				case NCIF:
					// Sheetal: 2/19/2019, commenting below line as its
					// already
					// executed in flow
					// webElement = getElementByType(controlId,
					// controlName,
					// WebHelper.control, imageType, ctrlValue);

					try {

						if (webElement.isDisplayed()) {

							// if(webElement.isEnabled() == true)

							Actions Click = new Actions(Automation.driver);
							highlightElement(webElement);
							Action Mouseclick = Click.moveToElement(webElement).clickAndHold().release().build();
							Mouseclick.perform();
							Thread.sleep(500);

							// Sheetal: 2/19/2019, commenting below line
							// as it
							// should not be part of Webbutton->NCIF
						}

					} catch (Exception e) {
						log.error(e.getMessage(), e);
						log.info("Element Not Found on Page.");
						log.info("Element Not Found on Page in NCIF case");
					}

					break;
				}
				break;
			// break;

			case WebElement:

				WebVerification.isFromVerification = true;
				// bhaskar
				switch (actionName) {

				case I: // Minaakshi
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
							|| !ctrlValue.equalsIgnoreCase("")) {

						webElement.click();
					}
					break;

				case NC:

					// waitForAjaxLoad(Automation.driver);
					// waitForPageLoadingToComplete();
					int i = 5;
					while (i > 1) {

						try {

							if (webElement.isDisplayed()) {

								webElement = getElementByType(controlId, controlName, WebHelper.control, imageType,
										ctrlValue);

								webElement.click();

								log.info("clicked");
								log.info("WebElement:NC->clicked");
								// Thread.sleep(500);
								// Automation.driver.manage().timeouts().pageLoadTimeout(60,TimeUnit.SECONDS);

								break;

							}
						}

						catch (Exception e) {
							log.error(e.getMessage(), e);
							try {

								webElement = getElementByType(controlId, controlName, WebHelper.control, imageType,
										ctrlValue);

								Thread.sleep(2000);

								webElement.click();

								log.info("catchclicked");
								log.info("WebElement:NC->catchclicked");
								break;

							} catch (Exception e1) {
								log.error(e1.getMessage(), e1);
								// webElement =
								// getElementByType(controlId,
								// controlName,control,imageType,ctrlValue);
								Thread.sleep(3000);

								i--;
								log.info("Element does not exist");
								log.info("WebElement:NC->Element does not exist");
								// e1.printStackTrace();

							}
						}
					}

					break;

				case JSClick: // Minaakshi
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
							|| !ctrlValue.equalsIgnoreCase("")) {

						JavascriptExecutor executor = (JavascriptExecutor) Automation.driver;
						executor.executeScript("arguments[0].click();", webElement);

					}
					break;

				case Read:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					uniqueNumber = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
					webElement.clear();
					webElement.sendKeys(uniqueNumber);
					break;
				case Write:
					if (ctrlValue == null) {

						WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo,
								colNo);

						break;
					}
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					if (isDMApplication || (ITAFWebDriver.isSuiteApplication()
							&& MainControllerSuite.transactionSilo.equalsIgnoreCase("DM"))) {// Minaakshi
						if (logicalName.equalsIgnoreCase("RequestID")) {
							DMProduct.writeRequestIDToExcel(ctrlValue, webElement, controlId, controlType, controlName,
									rowNo, colNo);

						} else {
							WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName,
									rowNo, colNo);

						}
						break;
					} else {
						WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo,
								colNo);

					}
					/*
					 * WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType,
					 * controlName, rowNo, colNo);
					 */
					break;

				case VD: // Minaakshi : 01-10-2020 Added for handling dynamic
							// field
							// level verification
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}

					if (ctrlValue.equalsIgnoreCase(null)) {
						ctrlValue = DMProduct.ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
					} else {
						String strTemp = ctrlValue;
						if (ctrlValue.contains("$value")) {// Minaakshi
															// :
							ctrlValue = DMProduct.ReadFromExcelUsingColumnName("", WebHelper.columnName);

						} else if (ctrlValue.contains("|*") || ctrlValue.contains("-*") || ctrlValue.contains("- *")) {// Minaakshi
																														// :
							// RGRS_SC_47_Data_01|Sales Hierarchy -
							// *value-Active-RLI Corp.-RT1004701(1)' //
							// 01-07-2019
							// RGRS_SC_47_Data_03|Hooper Pollard
							// Mccray-*value-Active-RT1004703
							// Minaakshi : 04-01-2019
							String[] temp = ctrlValue.split("\\|");
							String tempCtrlValue = DMProduct.ReadFromExcelUsingColumnName(temp[0].toString(),
									WebHelper.columnName);
							ctrlValue = temp[1].toString().replace("*value", tempCtrlValue);
							// PRDMV_PolicyData1|@PRDMV_SCPATIT_BD_Ent1@AGYName
						} else if (ctrlValue.contains("|@")) {// PRDMV_SCPATIT_BD_Ent1|@PRDMV_SCPATIT_BD_Ent1@EntityCode
																// : Minaakshi : 01-01-2021

							String[] temp = ctrlValue.split("\\@");
							String tempCtrlValue = DMProduct.ReadFromExcelUsingColumnName(temp[1].toString(),
									temp[2].toString());
							ctrlValue = tempCtrlValue;

							// TC|COl-TC|Col : AGY20740-AGY_2907203331 // Minaakshi : 01-03-2021
						} else if (ctrlValue.contains("|") && ctrlValue.contains("-") && !(ctrlValue.contains("XXX"))) {

							String[] temp1 = ctrlValue.split("\\-");

							String[] temp2 = temp1[0].split("\\|");
							String tempCtrlValue1 = DMProduct.ReadFromExcelUsingColumnName(temp2[0].toString(),
									temp2[1].toString());

							String[] temp3 = temp1[1].split("\\|");
							String tempCtrlValue2 = DMProduct.ReadFromExcelUsingColumnName(temp3[0].toString(),
									temp3[1].toString());

							ctrlValue = tempCtrlValue1 + "-" + tempCtrlValue2;

						} else if (ctrlValue.contains("|")) {// Minaakshi
																// :
																// 01-06-2019
							// Minaakshi : 04-01-2019
							String[] temp = ctrlValue.split("\\|");
							ctrlValue = temp[1].toString();
							// Amiie J-AGT20567-Active-
							// e1#d1&e2#d2&-Active-
						} else if (ctrlValue.contains("-") && ctrlValue.contains("#") && ctrlValue.contains("&")) {// Minaakshi

							// Minaakshi : 01-01-2021
							String[] temp1 = ctrlValue.split("\\&");

							String[] temp2 = temp1[0].split("\\#");
							String tempCtrlValue1 = DMProduct.ReadFromExcelUsingColumnName(temp2[0].toString(),
									temp2[1].toString());

							String[] temp3 = temp1[1].split("\\#");
							String tempCtrlValue2 = DMProduct.ReadFromExcelUsingColumnName(temp3[0].toString(),
									temp3[1].toString());

							ctrlValue = tempCtrlValue1 + "-" + tempCtrlValue2 + temp1[2].toString();

						} else if (ctrlValue.contains("-")) {// Minaakshi
							// :
							// 01-06-2019
							// Minaakshi : 01-09-2020
							String[] temp = ctrlValue.split("\\-");
							String tempCtrlValue1 = DMProduct.ReadFromExcelUsingColumnName(temp[0].toString(),
									WebHelper.columnName);
							String tempCtrlValue2 = DMProduct.ReadFromExcelUsingColumnName(temp[1].toString(),
									WebHelper.columnName);
							ctrlValue = tempCtrlValue1 + "-" + tempCtrlValue2;

						}

						else {
							ctrlValue = DMProduct.ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
						}

						/*
						 * if (ctrlValue.contains("/")&& ctrlValue.contains("$value")){//Minaakshi :
						 * 03-10-2018 ctrlValue = ctrlValue.replace("/", "-"); }
						 */

						// ***Change done by DM Team***//
						if (ctrlValue.contains("/") && strTemp.contains("$value")) {// Minaakshi:14-11-2018
							ctrlValue = ctrlValue.replace("/", "-");
						} // ***//

						if (strTemp.contains("$value")) {// Minaakshi :
							String strReplacedString = strTemp.replace("$value", ctrlValue);
							ctrlValue = strReplacedString;
							if (strTemp.contains("|")) {// Minaakshi : 01-03-2021
								String[] temp1 = ctrlValue.split("\\|");
								ctrlValue = temp1[1].toString();
							}
						}

					}

					if (WebVerification.isFromVerification == true) {
						currentValue = webElement.getText();
						if (currentValue.equalsIgnoreCase(null) || currentValue.equalsIgnoreCase("")) {
							currentValue = webElement.getAttribute("value");
						}

						break;
					}
					boolean textPresent1 = false;
					textPresent1 = webElement.getText().contains(ctrlValue);
					if (textPresent1 == false) {
						currentValue = Boolean.toString(textPresent1);

					} else {

						currentValue = ctrlValue;

					}

					break;
				case PopUpCount:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}

					// List<WebElement> a=webElement.getSize();
					List<WebElement> a = webElement.findElements(By.xpath(controlName));
					// List<WebElement> a=getElementByType(controlId,
					// controlName,control, imageType, ctrlValue);
					int b = a.size();
					currentValue = String.valueOf(b);
					if (WebVerification.isFromVerification == true) {
						if (ctrlValue.equals(currentValue)) {

							log.info("PASS");
						} else {
							log.info("FAIL");

						}
					}
					break;
				// log.info(b);
				// log.info(a);

				case V:
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {

						break;
					}

					// Minaakshi : 01-03-2021 : $value1;$value2

					if (ctrlValue.contains(";") && !(ctrlValue.contains("PRE=")) && !(ctrlValue.contains("POST="))) {
						String[] temp = ctrlValue.split("\\;");
						ctrlValue = temp[1].toString();
					}

					if (WebVerification.isFromVerification == true) {
						currentValue = webElement.getText();
						log.info(currentValue);

						// ***Change done by DM Team***/
						if (currentValue.equalsIgnoreCase(null) || currentValue.equalsIgnoreCase(""))// Minaakshi
						// :
						// 14-11-2018
						{
							currentValue = webElement.getAttribute("value");
						}
						// ***End***//

						break;
					}
					boolean textPresent = false;
					textPresent = webElement.getText().contains(ctrlValue);
					if (textPresent == false) {
						currentValue = Boolean.toString(textPresent);

					} else {
						currentValue = ctrlValue;

					}
					break;
				case V_Image:
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {

						break;
					}

					if (WebVerification.isFromVerification == true) {
						currentValue = webElement.getAttribute("src");
						log.info(currentValue);
					}

					break;
				case V_edit:
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {

						break;
					}
					if (webElement != null) {
						currentValue = webElement.getAttribute("value");
						if (StringUtils.equalsIgnoreCase(currentValue, null)) {
							currentValue = webElement.getText();
						}
						ctrlValue = "Edit/Error displayed: '" + currentValue + "'.";
						currentValue = ctrlValue;
					} else {
						currentValue = "Edit/Error not displayed. Expected at : " + ctrlValue;
						ctrlValue = currentValue;
					}
					break;

				case V_attributeValue:
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {

						break;
					}

					if (ctrlValue.equalsIgnoreCase("pickFromUniqueNumbers")) {
						ctrlValue = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
					}

					String attribute = "";
					String vType = "";

					String[] tempImageType = imageType.split(";;");
					if (!imageType.isEmpty() && tempImageType.length == 2) {
						attribute = tempImageType[1];
						vType = tempImageType[0];
					} else {
						attribute = "value";
						vType = "full";
					}

					try {
						highlightElement(webElement);
						if (attribute.equalsIgnoreCase("value")) {
							currentValue = webElement.getAttribute(attribute);
							if (StringUtils.equalsIgnoreCase(currentValue, null) || currentValue.equalsIgnoreCase(""))
								currentValue = webElement.getText();
						} else {
							currentValue = webElement.getAttribute(attribute);
						}
						if (vType.equalsIgnoreCase("partial")) {
							if (currentValue.contains(ctrlValue)) {
								// ctrlValue= ctrlValue +
								// " is present in attribute "+ attribute + " of
								// given
								// tag. Actual attribute value is " +
								// currentValue;

								currentValue = ctrlValue;
							}
						}

						if (currentValue == null) {
							currentValue = "Incorrect structure of V_attributeValue";
						}

					} catch (Exception e) {
						log.info("Error in V_attributeValue " + e.getMessage());
					}

					break;

				case V_text: // Sheetal - 2-5-2019
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {

						break;
					}

					// highlightElement(webElement);
					try {
						currentValue = webElement.getAttribute("value");

						if (StringUtils.equalsIgnoreCase(currentValue, null) || currentValue.equalsIgnoreCase("")) {
							currentValue = webElement.getText();
							if (StringUtils.equalsIgnoreCase(currentValue, null))

							{
								currentValue = "";
							}

						}

					} catch (NullPointerException e) {
						currentValue = "Element not found : " + controlName;

					}
					if (ctrlValue.contains("pickFromUniqueNumbers")) {
						// ctrlValue = ReadFromExcel(ctrlValue);
						String dynamicValue = "";
						dynamicValue = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
						ctrlValue = ctrlValue.replaceAll("pickFromUniqueNumbers", dynamicValue);
						if (dynamicValue.equalsIgnoreCase("")) {
							ctrlValue = "Value missing in UniqueNumber file";
						}
					} else if (ctrlValue.equalsIgnoreCase("blank")) {
						if (currentValue.trim().isEmpty()) {
							currentValue = ctrlValue;
						}
					} else if (ctrlValue.equalsIgnoreCase("not blank")) {
						if (!currentValue.contains("Element not found")) {
							if (!currentValue.isEmpty() || !currentValue.equalsIgnoreCase("")) {
								currentValue = ctrlValue;
							}
						}
					}

					break;

				case V_checkboxStatus: // Sheetal - 2-5-2019
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {

						break;
					}
					try {
						highlightElement(webElement);
						String checkedValue = webElement.getAttribute("checked");

						if (StringUtils.equalsIgnoreCase(checkedValue, "true")) {
							currentValue = "Checked";
						} else {
							currentValue = "Unchecked";
						}
					} catch (NullPointerException e) {
						currentValue = "Element not found : " + controlName;
					}
					if (currentValue.equalsIgnoreCase(ctrlValue)) {
						currentValue = ctrlValue;
					}
					break;

				case V_disableStatus: // Sheetal - 2-5-2019
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {

						break;
					}
					highlightElement(webElement);
					String disableStatus = webElement.getAttribute("disabled");

					if (StringUtils.equalsIgnoreCase(disableStatus, "true")) {// Minaakshi
																				// :
																				// 01-11-2019
						currentValue = "Disabled";
					} else {
						currentValue = "Enabled";
					}
					if (currentValue.equalsIgnoreCase(ctrlValue)) {
						currentValue = ctrlValue;
					}
					break;

				case VerifyPopUpElement:

					if (WebVerification.isFromVerification == true) {

						try {
							webElement = getElementByType(controlId, controlName, WebHelper.control, imageType,
									ctrlValue);

							if (webElement.isDisplayed()) {

								currentValue = webElement.getText();
								ctrlValue = currentValue;
								log.info(currentValue);

							}
						} catch (Exception e) {
							log.error(e.getMessage(), e);
							currentValue = "WebElement " + logicalName + " is not Present";
							ctrlValue = logicalName;
						}
					}

					break;

				case V_availabilityStatus: // Mrinmayee 15-11-2018
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {

						break;
					}

					if (webElement != null) {
						currentValue = "Available";
					} else {
						currentValue = "Not Available";
					}

					if (currentValue.equalsIgnoreCase(ctrlValue)) {
						currentValue = ctrlValue;
					} else if (!(ctrlValue.equalsIgnoreCase("available")
							|| ctrlValue.equalsIgnoreCase("not available"))) {

						ctrlValue = currentValue;
					}
					break;

				case VFCount: // Mrinmayee 15-11-2018
					if (webElement != null) {
						ctrlValue = "Expected Form Count = " + ctrlValue;
						currentValue = "FORM COUNT EXCEEDS";
					} else {
						ctrlValue = "Form verification is successful, if all above individual forms' status is pass | Expected Form Count = "
								+ ctrlValue;
						currentValue = ctrlValue;
					}

					break;
				case VFormCount:
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")) {
						break;
					}
					if (!formList.isEmpty()) {
						// check duplicate form instances
						List<String> tempForms = new ArrayList<String>();
						List<String> duplicateForms = new ArrayList<String>();
						String duplicateFormStatus = "";
						// copy formList to tempForms
						for (int k = 0; k < formList.size(); k++) {
							tempForms.add(formList.get(k));
						}

						for (int j = 0; j < tempForms.size(); j++) {
							for (int l = 1; l < tempForms.size(); l++) {
								if (tempForms.get(j).equals(tempForms.get(l)) && (j != l)) {
									boolean presentInDuplicate = false;
									for (int k = 0; k <= duplicateForms.size(); k++) {
										if (duplicateForms.get(k).equals(tempForms.get(j))) {
											presentInDuplicate = true;
										}
									}
									if (presentInDuplicate == false) {
										duplicateForms.add(tempForms.get(j));
									}

								}
							}
						}
						if (duplicateForms.isEmpty()) {
							duplicateFormStatus = "No duplicate form present.";
						} else {
							duplicateFormStatus = "Multiple instances of " + duplicateForms.size()
									+ " Forms present on Forms Tab.";

							for (int k = 0; k < duplicateForms.size(); k++) {
								duplicateFormStatus = duplicateFormStatus + "\n" + duplicateForms.get(k);
							}
						}

						// compare expected and actual form count
						if (formList.size() == Integer.parseInt(ctrlValue)) {
							ctrlValue = "Expected and Actual Form Count Matches. Count = " + ctrlValue + ". ";
							ctrlValue = ctrlValue + "\n" + duplicateFormStatus;
							currentValue = ctrlValue;
						} else {
							currentValue = "Mismatch in Form count. Expected Count = " + ctrlValue;
							ctrlValue = "Actual Count = " + formList.size();
							ctrlValue = ctrlValue + "\n" + duplicateFormStatus;
						}

					} else {
						if (ctrlValue.equalsIgnoreCase("0")) {
							ctrlValue = "No form is displayed on screen. Expected Form Count = " + ctrlValue;
							currentValue = ctrlValue;
						} else {
							ctrlValue = "Expected Form Count = " + ctrlValue;
							currentValue = "No form is displayed on screen.";
						}
					}

					break;
				case VFPresence: // Mrinmayee 15-11-2018
					if (webElement != null) {
						currentValue = ctrlValue;
					} else {
						currentValue = "Form is not present";
					}

					break;

				case VFormPresence:
					if (!ctrlValue.equals("")) {
						expectedForms.add(ctrlValue);
						if (!formList.isEmpty()) {
							boolean formFound = false;
							for (int k = 0; k < formList.size(); k++) {
								if (formList.get(k).equals(ctrlValue)) {
									formFound = true;
									break;
								}
							}
							if (formFound == true) {
								presentForms.add(ctrlValue);
								currentValue = ctrlValue;
							} else {
								missingForms.add(ctrlValue);
								currentValue = "Form is not present";
							}

						} else {
							ctrlValue = "";
							currentValue = "No form is displayed on screen.";
						}
					}
					break;

				case FormVerificationDetails:
					String temp = "", temp2 = "";
					boolean fail = false;

					// List missing forms
					if (!missingForms.isEmpty()) {
						fail = true;
						temp = missingForms.size() + " form(s) are missing.\n";
						for (int p = 0; p < missingForms.size(); p++) {
							temp = temp + missingForms.get(p) + "\n";
						}
					} else {
						temp = "No form is missing.";
					}

					// List additional forms
					List<String> tempForms = new ArrayList<String>();
					// copy formList to tempForms
					for (int k = 0; k < formList.size(); k++) {
						tempForms.add(formList.get(k));
					}
					tempForms.removeAll(expectedForms);
					// fill additionalForms
					for (int k = 0; k < tempForms.size(); k++) {
						additionalForms.add(tempForms.get(k));
					}

					if (!additionalForms.isEmpty()) {
						fail = true;
						temp2 = additionalForms.size() + " forms are additional.\n";
						for (int p = 0; p < additionalForms.size(); p++) {
							temp2 = temp2 + additionalForms.get(p) + "\n";
						}
					} else {
						temp2 = "No additional form present.";
					}

					if (fail == true) {
						ctrlValue = temp;
						currentValue = "\n" + temp2;
					} else {
						ctrlValue = temp + "\n" + temp2;
						currentValue = ctrlValue;
					}

					missingForms.removeAll(missingForms);
					additionalForms.removeAll(additionalForms);
					presentForms.removeAll(presentForms);
					formList.removeAll(formList);
					expectedForms.removeAll(expectedForms);

					break;
				case NCIF:

					// webElement = getElementByType(controlId,
					// controlName,control,imageType,ctrlValue);

					try {

						// if(!ctrlValue.equalsIgnoreCase("null"))

						if (webElement.isDisplayed()) {

							// if(webElement.isEnabled() == true)

							webElement.click();
							Thread.sleep(500);
						} else {
							log.info("Element is not displayed");
						}

					} catch (Exception e) {
						log.error(e.getMessage(), e);
						log.info("Element Not Found on Page.");
					}

					break;

				case Highlight:

					WebHelperUtil.fnHighlightMe(Automation.driver,
							Automation.driver.findElement(By.xpath(controlName)));

					// WebHelper.fnHighlightMe(Automation.driver,Automation.driver.findElement(controlName));

				}
				break;

			case MouseClick:

				try {

					if (webElement.isDisplayed()) {

						Actions builderClick = new Actions(Automation.driver);
						Action clickAction = builderClick.moveToElement(webElement).clickAndHold().release().build();
						clickAction.perform();

					}

				} catch (Exception e) {
					log.error(e.getMessage(), e);
					log.info("Element does not exist");
				}

				break;

			case C:

				try {

					if (!ctrlValue.equalsIgnoreCase("null"))

						currentValue = webElement.getText();

					if (currentValue.contentEquals(ctrlValue)) {

						System.out.print("Expected" + currentValue + "Matches with " + "Actual" + ctrlValue);
					}

					else {
						System.out.print("Expected" + currentValue + "Not Matches with " + "Actual" + ctrlValue);
					}
				} catch (Exception e) {
					log.error(e.getMessage(), e);

					log.info("Element Not Found");
				}

				break;

			case JSScript:
				((JavascriptExecutor) Automation.driver).executeScript(controlName, ctrlValue);

				break;

			// bhaskar
			// case IJSScript:
			// IJavascriptExecutor ijs = Automation.driver;
			// bhaskar

			case WaitToLoad:
				return WebHelper.doAction(null, null, null, null, controlType, controlId, controlName, ctrlValue, null,
						null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex, rowcount,
						rowNo, colNo, null, null, null);

			case WaitForPageToLoad:

				// Automation.driver.manage().timeouts().pageLoadTimeout(120, TimeUnit.SECONDS);
				Automation.driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(120));// Sel4

				WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, (Duration.ofSeconds(500)));

				WaitForPageLoad.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

				break;

			case WaitForElementToVisible: // Mrinmayee - Wait until the
				// expected value is present on
				// screen

				String temp1 = Automation.driver.findElement(By.xpath(controlName)).getText();
				log.info(temp1);
				int x = 200;

				while (x > 0) {

					if (temp1.equals(ctrlValue)) {
						// log.info("Element is present");
						log.info("temp1 value is:-" + temp1);
						log.info("ctrlValue value is:-" + ctrlValue);
						break;
					} else {
						String temp2 = Automation.driver.findElement(By.xpath(controlName)).getText();
						temp1 = temp2;
						log.info("temp1 value is:-" + temp1);
						log.info("ctrlValue value is:-" + ctrlValue);
						Thread.sleep(2000);
					}
					x--;
				}

				break;

			case AttributeisPresent:

				WebHelper.wait.until(ExpectedConditions.attributeContains(By.xpath(controlName), "value", ctrlValue));
				break;
			case WaitUntilElementInvisible:

				try {

					WebHelper.wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(controlName)));

					log.info("Element is Invisible");
				} catch (Exception e) {
					log.error(e.getMessage(), e);
					log.info("Element is visible");
				}

				break;

			case filedownload:

				FirefoxProfile profile = new FirefoxProfile();

				// Accept SSL certificate errors
				profile.setAcceptUntrustedCertificates(true);
				profile.setAssumeUntrustedCertificateIssuer(true);

				// Download a file in FF browser
				profile.setPreference("browser.download.folderList", 2);
				profile.setPreference("browser.helperApps.alwaysAsk.force", false);

				profile.setPreference("browser.download.dir", "D:\\Downloads");
				profile.setPreference("browser.download.defaultFolder", "D:\\Downloads");
				profile.setPreference("browser.download.manager.showWhenStarting", false);
				// Set MIME types
				profile.setPreference("browser.helperApps.neverAsk.saveToDisk",
						"multipart/x-zip,application/zip,application/x-zip-compressed,application/x-compressed,application/msword,application/csv,text/csv,image/png ,image/jpeg, application/pdf, text/html,text/plain,  application/excel, application/vnd.ms-excel, application/x-excel, application/x-msexcel, application/octet-stream");

				// FirefoxDriver driver = new FirefoxDriver(profile);

				Thread.sleep(2000);

				// driver.findElement(By.id("download")).click();

				break;

			case PageRefresh:

				switch (actionName) { // Mrinmaye 05-12-2018
				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}

					// Automation.driver.navigate().refresh();
					Automation.refreshPage(controlName);
					break;

				case NC:
					Automation.refreshPage(controlName);
					break;
				}
				break;

			case WaitForObjectPresent:

				WebHelperUtil.WebObjectPresent(webElement);

				break;

			case Wait:

				if (StringUtils.equalsIgnoreCase(Config.applyStaticWait, "false")) {
					log.info("Wait not applied");
				} else {
					Thread.sleep(Integer.parseInt(controlName) * 1000);

				}
				break;
			case Wait_DM: // Minaakshi : 05-02-2019

				if (controlName.equalsIgnoreCase("$value")) {
					if (ctrlValue.equalsIgnoreCase("HIGH")) {
						Thread.sleep(30000);
					} else if (ctrlValue.equalsIgnoreCase("VERYHIGH")) {
						Thread.sleep(60000);
					} else if (ctrlValue.equalsIgnoreCase("MEDIUM")) {
						Thread.sleep(10000);
					} else if (ctrlValue.equalsIgnoreCase("LOW")) {
						Thread.sleep(5000);
					} else if (ctrlValue.equalsIgnoreCase("VERYLOW")) {
						Thread.sleep(2000);
					} else {
						break;
					}
				} else {

					Thread.sleep(1000);// Minaakshi : 01-04-2019
					break;
				}

				break;

			case Wait_IfValue: // Mrinmayee : 15-11-2018
				switch (actionName) {
				case I:
					if (!ctrlValue.isEmpty()) {
						Thread.sleep(Integer.parseInt(controlName) * 1000);
					}
					break;
				}
				break;

			case D_Wait:
				switch (actionName) {
				case I:
					if (!ctrlValue.isEmpty()) {
						waitForVaadin();
						Thread.sleep(1500);
					}
					break;

				case NC:
					waitForVaadin();
					Thread.sleep(1500);
					break;
				}
				break;

			case RowNumbersToExecute:
				log.info(ctrlValue);
				break;

			case CheckBox:
				switch (actionName) {
				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					if (Config.projectName.equals("DistributionManagement")) { // Minaakshi
						// :
						// 18-12-2018
						webElement.click();
						break; // Minaakshi : 01-10-2019

					} else {
						int j = 4;
						while (j > 0) {

							if (ctrlValue.equalsIgnoreCase("Notdisplayed")) {
								break;
							}

							/*
							 * if(webElement==null){
							 * 
							 * break; }
							 */
							try {

								if (webElement.isDisplayed()) {

									webElement = getElementByType(controlId, controlName, WebHelper.control, imageType,
											ctrlValue);

									if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {

										if (!webElement.isSelected()) {
											Coordinates coordinate = ((Locatable) webElement).getCoordinates();
											coordinate.onPage();
											coordinate.inViewPort();
											JavascriptExecutor executor = (JavascriptExecutor) Automation.driver;
											executor.executeScript("arguments[0].click();", webElement);
											// Thread.sleep(5000);
											// webElement.click();
											if (webElement.isSelected()) {
												// Automation.driver.manage().timeouts().pageLoadTimeout(60,TimeUnit.SECONDS);
												break;
											}
										} else {

											break;
										}
									} else {
										Thread.sleep(500);
										if (!webElement.isSelected()) {

											break;
										} else {
											// webElement =
											// getElementByType(controlId,
											// controlName,control,imageType,ctrlValue);

											JavascriptExecutor executor = (JavascriptExecutor) Automation.driver;
											executor.executeScript("arguments[0].click();", webElement);

											break;
										}
									}
								}
							}

							catch (Exception e) {
								log.error(e.getMessage(), e);
								try {
									Thread.sleep(500);
									webElement = getElementByType(controlId, controlName, WebHelper.control, imageType,
											ctrlValue);

									// webElement.click();
									break;
								} catch (Exception e1) {
									log.error(e1.getMessage(), e1);

									Thread.sleep(2000);
									log.info("Checkbox not found");
									j--;
								}

							}

						}
					}

					break;

				case NC:
					webElement = getElementByType(controlId, controlName, WebHelper.control, imageType, ctrlValue);
					webElement.click();
					break;

				case SetCheckbox:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}

					// ctrlValue = getCellData(logicalName,headerValues,
					// valuesRowIndex, valuesHeader);
					int i = 4;

					while (i > 0) {

						if (ctrlValue.equalsIgnoreCase("Notdisplayed")) {
							break;
						}

						/*
						 * if(webElement==null){
						 * 
						 * break; }
						 */
						try {

							if (webElement.isDisplayed()) {

								webElement = getElementByType(controlId, controlName, WebHelper.control, imageType,
										ctrlValue);

								if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {

									if (!webElement.isSelected()) {
										Coordinates coordinate = ((Locatable) webElement).getCoordinates();
										coordinate.onPage();
										coordinate.inViewPort();
										JavascriptExecutor executor = (JavascriptExecutor) Automation.driver;
										executor.executeScript("arguments[0].click();", webElement);
										// Thread.sleep(5000);
										// webElement.click();
										if (webElement.isSelected()) {
											// Automation.driver.manage().timeouts().pageLoadTimeout(60,TimeUnit.SECONDS);
											break;
										}
									} else {

										break;
									}
								} else {
									Thread.sleep(500);
									if (!webElement.isSelected()) {

										break;
									} else {
										// webElement =
										// getElementByType(controlId,
										// controlName,control,imageType,ctrlValue);

										JavascriptExecutor executor = (JavascriptExecutor) Automation.driver;
										executor.executeScript("arguments[0].click();", webElement);

										break;
									}
								}
							}
						}

						catch (Exception e) {
							log.error(e.getMessage(), e);
							try {
								Thread.sleep(500);
								webElement = getElementByType(controlId, controlName, WebHelper.control, imageType,
										ctrlValue);

								// webElement.click();
								break;
							} catch (Exception e1) {

								Thread.sleep(2000);
								log.info("Checkbox not found");
								i--;
							}

						}

					}

				}
				break;
			case Radio:
				switch (actionName) {
				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}

					if (!ctrlValue.isEmpty()) {
						webElement = getElementByType(controlId, controlName, WebHelper.control, imageType, ctrlValue);
						int i = 4;
						while (i > 1) {
							try {
								if (webElement.isDisplayed()) {

									if (!webElement.isSelected()) {

										webElement.click();
										break;
									}
									// Thread.sleep(1000);
									if (webElement.isSelected()) {

										break;
									}
								}

							} catch (Exception e) {
								log.error(e.getMessage(), e);
								try {
									Thread.sleep(2000);
									webElement.click();
									break;
								} catch (Exception e1) {
									Thread.sleep(2000);
									System.out.print("WebElement Not Found");
									i--;
								}
							}
						}
					}

					break;

				case SelectRadioButton:
					if (!ctrlValue.isEmpty()) {
						webElement = getElementByType(controlId, controlName, WebHelper.control, imageType, ctrlValue);
						int i = 4;
						while (i > 1) {
							try {
								if (webElement.isDisplayed()) {

									webElement.click();
									break;

								}

							} catch (Exception e) {
								log.error(e.getMessage(), e);
								try {
									Thread.sleep(2000);
									webElement.click();
									break;
								} catch (Exception e1) {
									Thread.sleep(2000);
									System.out.print("WebElement Not Found");
									i--;
								}
							}
						}
					}

					break;

				case NC:

					if (!ctrlValue.isEmpty()) {
						webElement = getElementByType(controlId, controlName, WebHelper.control, imageType, ctrlValue);
						int i = 4;
						while (i > 1) {
							try {
								if (webElement.isDisplayed()) {
									webElement.click();
									break;
								}

							} catch (Exception e) {
								log.error(e.getMessage(), e);
								try {
									Thread.sleep(2000);
									webElement = getElementByType(controlId, controlName, WebHelper.control, imageType,
											ctrlValue);

									webElement.click();
									break;
								} catch (Exception e1) {
									Thread.sleep(2000);
									System.out.print("WebElement Not Found");
									i--;
								}
							}
						}

					}
					break;
				case V:
					if (ctrlValue == null || ctrlValue.trim().equals("") || ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					if (webElement.isSelected()) {
						currentValue = webElement.getAttribute(controlName.toString());
					}
					break;
				case F:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					if (webElement != null) {
						currentValue = "Y";
					}
					break;
				}
				break;

			case WebLink:
				if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
						|| !ctrlValue.equalsIgnoreCase(""))// Minaakshi

				// :
				// 03-10-2018
				{
					webElement.click();
				}
				break;
			case CloseWindow:// added this Case to bypass page loading
				// after
				// clicking the event
				switch (actionName) {
				case Read:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					uniqueNumber = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
					WebElementList = WebHelperUtil.getElementsByType(controlId, controlName, controlType, imageType,
							uniqueNumber);
					webElement = WebHelperUtil.GetControlByIndex("", WebElementList, controlId, controlName,
							controlType, uniqueNumber);

					webElement.click();
					break;
				case Write:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo,
							colNo);

					break;
				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					if (controlId.equalsIgnoreCase("LinkValue")) {
						webElement.click();
					} else {
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
							webElement.click();
						}
					}
					break;
				case NC:
					webElement = getElementByType(controlId, controlName, WebHelper.control, imageType, ctrlValue);
					// webElement.click();
					Actions MousebuilderClick = new Actions(Automation.driver);
					Action MouseclickAction = MousebuilderClick.moveToElement(webElement).clickAndHold().release()
							.build();

					MouseclickAction.perform();
					log.info("clicked");
					// MousebuilderClick
					// .keyDown(Keys.CONTROL).click(webElement).keyUp(Keys.CONTROL).build().perform();
					break;
				}
				break;

			case WaitForJS:
				waitForCondition();
				break;

			case ListBox:
			case ActionClick:
				switch (actionName) {
				case I:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					if (ctrlValue == null || ctrlValue.trim().equals("")) {

						break;
					}
					try {
						if (!ctrlValue.equalsIgnoreCase("")) {
							// sheetal 2-7-2019
							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
									webElement);

							Thread.sleep(500);

							Actions builderClick = new Actions(Automation.driver);
							highlightElement(webElement);
							Action clickAction = builderClick.moveToElement(webElement).clickAndHold().release()
									.build();

							clickAction.perform();

						}
					} catch (Exception e) {
						log.error(e.getMessage(), e);
						log.info("Element not exist");
					}
					break;

				case NC:
					Actions builderClick = new Actions(Automation.driver);
					highlightElement(webElement);
					Action clickAction = builderClick.moveToElement(webElement).clickAndHold().release().build();
					clickAction.perform();
					break;
				}
				break;
			case Dropdown:
				switch (actionName) {
				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					Select s1 = new Select(webElement);
					s1.selectByVisibleText(ctrlValue);
					break;
				}
				break;

			case WindowAlertOk:
				switch (actionName) {
				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					Automation.driver.switchTo().alert().accept();
					break;
				}
				break;

			case WebList:
				switch (actionName) {
				case Read:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					uniqueNumber = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
					new Select(webElement).selectByVisibleText(uniqueNumber);
					break;
				case Write:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo,
							colNo);

					break;
				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}

					ExpectedCondition<Boolean> isTextPresent = CommonExpectedConditions
							.textToBePresentInElement(webElement, ctrlValue);

					if (isTextPresent != null) {
						if (webElement != null) {
							new Select(webElement).selectByVisibleText(ctrlValue);
						}
					}
					break;
				case V:
					if (ctrlValue == null || ctrlValue.trim().equals("") || ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					if (!ctrlValue.contains(",")) {
						currentValue = new Select(webElement).getFirstSelectedOption().getText();
						if (currentValue.isEmpty()) {
							currentValue = new Select(webElement).getFirstSelectedOption().getAttribute("value");
						}

						break;
					} else {
						currentValue = new String();
						List<WebElement> currentValues = new ArrayList<WebElement>();
						currentValues = new Select(webElement).getOptions();

						for (int j = 0; j < currentValues.size(); j++) {
							if (j + 1 == currentValues.size())
								currentValue = currentValue.concat(currentValues.get(j).getText());
							else {
								currentValue = currentValue.concat(currentValues.get(j).getText() + ",");
							}
						}
						break;
					}
				}
				break;

			// New code for AJAX Dropdown with dojo
			case AjaxWebList:
				switch (actionName) {
				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					webElement.click();
					break;
				case VA:

					if (ctrlValue == null || ctrlValue.trim().equals("") || ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					Thread.sleep(20000);
					currentValue = new String();
					List<WebElement> currentValues = new ArrayList<WebElement>();
					currentValues = Automation.driver.findElements(By.xpath(controlName));

					for (int j = 0; j < currentValues.size(); j++) {
						if (j + 1 == currentValues.size())
							currentValue = currentValue.concat(currentValues.get(j).getText());
						else {
							currentValue = currentValue.concat(currentValues.get(j).getText() + ",");
						}
					}
					break;

				}
				break;

			case IFrame:
				switch (actionName) { // Mrinmayee 05-12-2018
				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					log.info("In method doaction debug4");
					Thread.sleep(5000);

					if (controlName.startsWith("//iframe")) {

						WebDriverWait wait1 = new WebDriverWait(Automation.driver, (Duration.ofSeconds(700)));

						wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

						Automation.driver.switchTo().frame(Automation.driver.findElement(By.xpath(controlName)));

					}

					else {

						Automation.driver.switchTo().frame(controlName);

					}

					log.info("In method doaction debug5");
					break;

				case NC:
					log.info("In method doaction debug4");
					Thread.sleep(5000);

					if (controlName.startsWith("//iframe")) {

						WebDriverWait wait1 = new WebDriverWait(Automation.driver, (Duration.ofSeconds(700)));

						wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

						Automation.driver.switchTo().frame(Automation.driver.findElement(By.xpath(controlName)));

					}

					else {

						Automation.driver.switchTo().frame(controlName);

					}

					log.info("In method doaction debug5");
					break;
				}
				break;

			case Browser:
			case URL:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case NewBrowser:
				// Thread.sleep(3000); //DS:Check if required
				String parentWindow = Automation.driver.getWindowHandle();
				Set<String> handles = Automation.driver.getWindowHandles();
				for (String windowHandle : handles) {
					if (!windowHandle.equals(parentWindow)) {
						Automation.driver.switchTo().window(windowHandle);
						// <!--Perform your operation here for new window-->
						// Automation.driver.close(); //closing child window
						// Automation.driver.switchTo().window(parentWindow);
						// //cntrl to parent window
					}
				}
				break;

			case CloseBrowser:

				Automation.driver.close();

				break;

			case CloseAndLaunchNewBrowser:

				if (ctrlValue == null || ctrlValue.trim().equals("")) {// Minaakshi
					// :
					// 01-04-2019
					break;
				}

				Automation.driver.quit();
				Automation.setUp();

				break;

			// case RowNumbers:
			// System.out.print("");

			case Menu:
				webElement.click();
				break;

			case Alert:
				switch (actionName) {
				case V:
					if (ctrlValue == null || ctrlValue.trim().equals("") || ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					Alert alert = Automation.driver.switchTo().alert();
					if (alert != null) {
						currentValue = alert.getText();
						log.info("Alert found on the web page");
						log.info(currentValue);
						alert.accept();

						Thread.sleep(5000);
					}
					break;
				case NC:
					// wait.until(ExpectedConditions.presenceOfElementLocated(alert()));
					Alert alert1 = Automation.driver.switchTo().alert();
					if (alert1 != null) {
						alert1.accept();
						Thread.sleep(2000);
					} else {

						System.out.print("Alert Not found");
					}
					break;

				case I: // Mrinmayee

					if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
						Alert alert2 = Automation.driver.switchTo().alert();
						if (alert2 != null) {
							alert2.accept();
							Thread.sleep(2000);
						} else {

							System.out.print("Alert Not found");
						}
					} else {
						break;
					}
					break;
				}
				break;

			case WebImage:
				webElement.sendKeys(Keys.TAB);
				webElement.click();
				Thread.sleep(5000);
				for (int Seconds = 0; Seconds <= Integer.parseInt(Config.timeOut); Seconds++) {
					if (!((Automation.driver.getWindowHandles().size()) > 1)) {
						webElement.click();
						Thread.sleep(5000);
					} else {
						// break;
					}
				}
				break;

			/*
			 * case ActionClick: Actions builderClick = new Actions(Automation.driver);
			 * Action clickAction = builderClick.
			 * moveToElement(webElement).clickAndHold().release().build();
			 * clickAction.perform(); break;
			 */

			case ActionDoubleClick:
				Actions builderdoubleClick = new Actions(Automation.driver);
				builderdoubleClick.doubleClick(webElement).build().perform();// TM-27/01/2015
				// :-
				// commented
				// following
				// code
				// and
				// used
				// this
				// code
				// for
				// simultaneous
				// clicks
				// Action doubleClickAction =
				// builderdoubleClick.moveToElement(webElement).click().build();
				// doubleClickAction.perform();
				// doubleClickAction.perform();
				break;

			case ActionClickandEsc:
				Actions clickandEsc = new Actions(Automation.driver);
				Action clickEscAction = clickandEsc.moveToElement(webElement).click().sendKeys(Keys.ENTER, Keys.ESCAPE)
						.build();

				clickEscAction.perform();
				break;

			case ActionMouseOver:
				Actions builderMouserOver = new Actions(Automation.driver);
				Action mouseOverAction = builderMouserOver.moveToElement(webElement).build();
				mouseOverAction.perform();
				break;

			case SwitchtoNewBrowser:

			case Calendar:
				// Thread.sleep(5000);
				Boolean isCalendarDisplayed = Automation.driver.switchTo().activeElement().isDisplayed();
				log.info(isCalendarDisplayed);
				if (isCalendarDisplayed == true) {
					String[] dtMthYr = ctrlValue.split("/");
					WebElement Year = WaitTool.waitForElement(Automation.driver, By.name("year"),
							Integer.parseInt(Config.timeOut));// Automation.driver.findElement(By.name("year"));

					while (!Year.getAttribute("value").equalsIgnoreCase(dtMthYr[2])) {
						if (Integer.parseInt(Year.getAttribute("value")) > Integer.parseInt(dtMthYr[2])) {
							WebElement yearButton = WaitTool.waitForElement(Automation.driver, By.id("button1"),
									Integer.parseInt(Config.timeOut));// Automation.driver.findElement(By.id("button1"));
							yearButton.click();
						} else if (Integer.parseInt(Year.getAttribute("value")) < Integer.parseInt(dtMthYr[2])) {
							WebElement yearButton = WaitTool.waitForElement(Automation.driver, By.id("Button5"),
									Integer.parseInt(Config.timeOut));// Automation.driver.findElement(By.id("Button5"));
							yearButton.click();
						}
					}
					Select date = new Select(WaitTool.waitForElement(Automation.driver, By.name("month"),
							Integer.parseInt(Config.timeOut)));

					WebHelper.month = CalendarSnippet.getMonthForInt(Integer.parseInt(dtMthYr[1]));
					date.selectByVisibleText(WebHelper.month);
					WebElement Day = WaitTool.waitForElement(Automation.driver, By.id("Button6"),
							Integer.parseInt(Config.timeOut));// Automation.driver.findElement(By.id("Button6"));

					int day = 6;
					while (Day.getAttribute("value") != null) {
						Day = WaitTool.waitForElement(Automation.driver, By.id("Button" + day),
								Integer.parseInt(Config.timeOut));// Automation.driver.findElement(By.id("Button"+day));

						if (Day.getAttribute("value").toString().equalsIgnoreCase(dtMthYr[0])) {
							Day.click();
							break;
						}
						day++;
					}
				} else {
					log.info("Calendar not Diplayed");
				}
				// Automation.selenium.click(controlName);
				break;

			case CalendarNew:
				isCalendarDisplayed = Automation.driver.switchTo().activeElement().isDisplayed();
				log.info(isCalendarDisplayed);
				if (isCalendarDisplayed == true) {

					String[] dtMthYr = ctrlValue.split("/");
					Thread.sleep(2000);
					// String[] CurrentDate =
					// dtFormat.format(frmDate).split("/");
					WebElement Monthyear = Automation.driver.findElement(By.xpath("//table/thead/tr/td[2]"));
					String Monthyear1 = Monthyear.getText();
					String[] Monthyear2 = Monthyear1.split(",");
					Monthyear2[1] = Monthyear2[1].trim();

					WebHelper.month = CalendarSnippet.getMonthForString(Monthyear2[0]);

					while (!Monthyear2[1].equalsIgnoreCase(dtMthYr[2])) {
						if (Integer.parseInt(Monthyear2[1]) > Integer.parseInt(dtMthYr[2])) {
							WebElement yearButton = Automation.driver.findElement(By.cssSelector("td:contains('�')"));
							yearButton.click();
							Monthyear2[1] = Integer.toString(Integer.parseInt(Monthyear2[1]) - 1);
						} else if (Integer.parseInt(Monthyear2[1]) < Integer.parseInt(dtMthYr[2])) {
							WebElement yearButton = Automation.driver.findElement(By.cssSelector("td:contains('�')"));
							yearButton.click();
							Monthyear2[1] = Integer.toString(Integer.parseInt(Monthyear2[1]) + 1);
						}
					}

					while (!WebHelper.month.equalsIgnoreCase(dtMthYr[1])) {
						if (Integer.parseInt(WebHelper.month) > Integer.parseInt(dtMthYr[1])) {
							WebElement monthButton = Automation.driver.findElement(By.cssSelector("td:contains('�')"));
							monthButton.click();
							if (Integer.parseInt(WebHelper.month) < 11) {
								WebHelper.month = "0" + Integer.toString(Integer.parseInt(WebHelper.month) - 1);
							} else {
								WebHelper.month = Integer.toString(Integer.parseInt(WebHelper.month) - 1);
							}

						} else if (Integer.parseInt(WebHelper.month) < Integer.parseInt(dtMthYr[1])) {
							WebElement monthButton = Automation.driver.findElement(By.cssSelector("td:contains('�')"));
							monthButton.click();
							if (Integer.parseInt(WebHelper.month) < 9) {
								WebHelper.month = "0" + Integer.toString(Integer.parseInt(WebHelper.month) + 1);
							} else {
								WebHelper.month = Integer.toString(Integer.parseInt(WebHelper.month) + 1);
							}
						}
					}

					WebElement dateButton = Automation.driver
							.findElement(By.cssSelector("td.day:contains('" + dtMthYr[0] + "')"));

					log.info(dateButton);
					dateButton.click();

				} else {
					log.info("Calendar not Diplayed");
				}
				break;

			case CalendarIPF:
				String[] dtMthYr = ctrlValue.split("/");
				Thread.sleep(2000);
				String year = dtMthYr[2];
				String monthNum = dtMthYr[1];
				String day = dtMthYr[0];

				// Xpath for Year, mMnth & Days
				String xpathYear = "//div[@class='datepicker datepicker-dropdown dropdown-menu datepicker-orient-left datepicker-orient-bottom']/div[@class='datepicker-years']";
				String xpathMonth = "//div[@class='datepicker datepicker-dropdown dropdown-menu datepicker-orient-left datepicker-orient-bottom']/div[@class='datepicker-months']";
				String xpathDay = "//div[@class='datepicker datepicker-dropdown dropdown-menu datepicker-orient-left datepicker-orient-bottom']/div[@class='datepicker-days']";

				// Selecting year in 3 steps
				Automation.driver.findElement(By.xpath(xpathDay + "/table/thead/tr[1]/th[2]")).click();
				Automation.driver.findElement(By.xpath(xpathMonth + "/table/thead/tr/th[2]")).click();
				Automation.driver
						.findElement(By.xpath(
								xpathYear + "/table/tbody/tr/td/span[@class='year'][contains(text()," + year + ")]"))

						.click();

				// Selecting month in 1 step
				Automation.driver.findElement(By.xpath(xpathMonth + "/table/tbody/tr/td/span[" + monthNum + "]"))
						.click();

				// Selecting day in 1 step

				Automation.driver
						.findElement(
								By.xpath(xpathDay + "/table/tbody/tr/td[@class='day'][contains(text()," + day + ")]"))
						.click();

				break;

			case Calendar_DM:// Minaakshi : 01-03-2018
				if (ctrlValue == null || ctrlValue.trim().equals("") || ctrlValue.equalsIgnoreCase("IGNORE")) {
					break;
				}

				if (action.equalsIgnoreCase("Read")) {// Minaakshi : 01-10-2019
					String strTemp = null;
					strTemp = DMProduct.ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
					ctrlValue = strTemp;
				}

				String[] dtMthYr_dm = ctrlValue.split("/");
				Thread.sleep(2000);
				String year_dm = dtMthYr_dm[2];
				String monthNum_dm = dtMthYr_dm[0];
				monthNum_dm = CalendarSnippet.getMonthForInt(Integer.parseInt(dtMthYr_dm[0])).substring(0, 3);
				String day_dm = dtMthYr_dm[1];

				// Xpath for Year, mMnth & Days
				String xpathYear_dm = "//div[@class='ui-datepicker ui-widget ui-widget-content ui-helper-clearfix ui-corner-all']/div[@class='ui-datepicker-header ui-widget-header ui-helper-clearfix ui-corner-all']/div";
				String xpathMonth_dm = "//div[@class='ui-datepicker ui-widget ui-widget-content ui-helper-clearfix ui-corner-all']/div[@class='ui-datepicker-header ui-widget-header ui-helper-clearfix ui-corner-all']/div";
				String xpathDay_dm = "//div[@class='ui-datepicker ui-widget ui-widget-content ui-helper-clearfix ui-corner-all']/table[@class='ui-datepicker-calendar']";

				// Selecting year in 3 steps
				Automation.driver.findElement(By.xpath(xpathYear_dm + "/select[@class='ui-datepicker-year']")).click();
				Thread.sleep(2000);
				Automation.driver.findElement(By.xpath(
						xpathYear_dm + "/select[@class='ui-datepicker-year']/option[contains(.," + year_dm + ")]"))
						.click();

				// Selecting month in 1 step
				Automation.driver.findElement(By.xpath(xpathMonth_dm + "/select[@class='ui-datepicker-month']"))
						.click();

				Thread.sleep(2000);
				Automation.driver.findElement(

						By.xpath(xpathMonth_dm + "/select[@class='ui-datepicker-month']/option[contains(.,'"
								+ monthNum_dm + "')]"))
						.click();

				// Selecting day in 1 step
				Automation.driver.findElement(By.xpath(xpathDay_dm + "/tbody/tr/td[contains(.," + day_dm + ")]"))
						.click();

				Thread.sleep(2000);

				// Clicking on Done button
				Automation.driver.findElement(By.xpath(".//*[@id='ui-datepicker-div']/div[3]/button[2]")).click();
				break;

			case CalendarEBP:
				String[] dtMthYrEBP = ctrlValue.split("/");
				Thread.sleep(2000);
				String yearEBP = dtMthYrEBP[2];
				String monthNumEBP = CalendarSnippet.getMonthForInt(Integer.parseInt(dtMthYrEBP[1])).substring(0, 3);
				String dayEBP = dtMthYrEBP[0];

				// common path used for most of the elements
				String pathToVisibleCalendar = "//div[@class='ajax__calendar'][contains(@style, 'visibility: visible;')]/div";

				// following is to click the title once to reach the year
				// page
				WebHelper.wait
						.until(ExpectedConditions.elementToBeClickable(
								By.xpath(pathToVisibleCalendar + "/div[@class='ajax__calendar_header']/div[3]/div")))
						.click();

				// check if 'Dec' is visibly clickable after refreshing
				WebHelper.wait.until(ExpectedConditions.elementToBeClickable(
						By.xpath(pathToVisibleCalendar + "/div/div/table/tbody/tr/td/div[contains(text(), 'Dec')]")));
				// following is to click the title once again to reach the
				// year page

				Automation.driver
						.findElement(
								By.xpath(pathToVisibleCalendar + "/div[@class='ajax__calendar_header']/div[3]/div"))
						.click();

				// common path used for most of the elements while selection
				// of year, month and date
				pathToVisibleCalendar = "//div[@class='ajax__calendar'][contains(@style, 'visibility: visible;')]/div/div/div/table/tbody/tr/td";

				// each of the following line selects the year, month and
				// date
				WebHelper.wait
						.until(ExpectedConditions.elementToBeClickable(
								By.xpath(pathToVisibleCalendar + "/div[contains(text()," + yearEBP + ")]")))

						.click();
				WebHelper.wait
						.until(ExpectedConditions.elementToBeClickable(By.xpath(pathToVisibleCalendar
								+ "/div[@class='ajax__calendar_month'][contains(text(),'" + monthNumEBP + "')]")))
						.click();
				WebHelper.wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
						pathToVisibleCalendar + "/div[@class='ajax__calendar_day'][contains(text()," + dayEBP + ")]")))
						.click();

				break;

			/** Code for window popups **/
			case Window:
				switch (actionName) {
				case O:
					// String parentHandle = Automation.driver.getWindowHandle();//Sel4
					for (String winHandle : Automation.driver.getWindowHandles()) {
						Automation.driver.switchTo().window(winHandle);
						/*
						 * if (Automation.driver.getTitle().equalsIgnoreCase ( controlName)) {
						 * Automation.driver.close(); }
						 */
					}
					// Automation.driver.switchTo().window(parentHandle);
					break;
				}
				break;

			case WebTable:
				switch (actionName) {
				case MultiTable:
					WebHelperUtil.handleWebTableAction(webElement);

				case Read:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
					break;
				case Write:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo,
							colNo);

					break;
				case NC:
					WebElement table = webElement;
					List<WebElement> tableRows = table.findElements(By.tagName("tr"));
					int tableRowIndex = 0;
					// int tableColumnIndex = 0;
					boolean matchFound = false;
					for (WebElement tableRow : tableRows) {
						tableRowIndex += 1;
						List<WebElement> tableColumns = tableRow.findElements(By.tagName("td"));
						if (tableColumns.size() > 0) {
							for (WebElement tableColumn : tableColumns)
								if (tableColumn.getText().equals(ctrlValue)) {
									matchFound = true;
									log.info(tableRowIndex);
									List<Object> elementProperties = getPropertiesOfWebElement(
											tableColumns.get(Integer.parseInt(colNo)), imageType);
									controlName = elementProperties.get(0).toString();
									if (controlName.equals("")) {
										controlName = elementProperties.get(1).toString();
									}
									controlType = elementProperties.get(2).toString();
									webElement = (WebElement) elementProperties.get(3);
									doAction(imageType, controlType, controlId, controlName, ctrlValue, logicalName,
											action, webElement, Results, strucSheet, valSheet, tableRowIndex, rowcount,
											rowNo, colNo);

									break;
								}
							if (matchFound) {
								break;
							}
						}

					}
					break;

				case NCTable:

					int i;
					int j;
					WebElement Table = webElement;
					List<WebElement> TableRows = Table.findElements(By.tagName("tr"));
					for (i = 1; i <= TableRows.size(); i++) {
						List<WebElement> TableColumns = TableRows.get(i).findElements(By.tagName("td"));

						for (j = 0; j < TableColumns.size(); j++) {

							if (TableColumns.get(j).getText().matches(ctrlValue)) {

								TableColumns.get(j).click();

							}

						}
					}
					break;
				case V:
					if (ctrlValue == null || ctrlValue.trim().equals("") || ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					WriteToDetailResults(ctrlValue, "", logicalName);
					try {
						Thread.sleep(2000);
					} catch (InterruptedException e) {
						log.error(e.getMessage(), e);
						e.printStackTrace();
					}
					break;
				case TableInput:
				case FIND:
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);
				}
				break;

			case Robot:

				if (ctrlValue != null && ctrlValue.trim().equals("IGNORE")) {// Minaakshi
					// :
					// 03-10-2018
					break;
				}

				if (controlName.equalsIgnoreCase("SetFilePath")) {
					// Automation.driver.switchTo().alert();
					// Robot robot=new Robot();
					StringSelection stringSelection = new StringSelection(ctrlValue);
					Toolkit.getDefaultToolkit().getSystemClipboard().setContents(stringSelection, null);
					WebHelper.robot.delay(1000);
					WebHelper.robot.keyPress(KeyEvent.VK_CONTROL);
					WebHelper.robot.keyPress(KeyEvent.VK_V);
					// download
					WebHelper.robot.keyRelease(KeyEvent.VK_V);
					WebHelper.robot.keyRelease(KeyEvent.VK_CONTROL);

				} else if (controlName.equalsIgnoreCase("TAB")) {
					// webElement.sendKeys(Keys.TAB);
					// webElement.sendKeys(Keys.ENTER);
					// webElement.sendKeys(Keys.TAB);
					WebHelper.robot.keyPress(KeyEvent.VK_TAB);
					WebHelper.robot.keyRelease(KeyEvent.VK_TAB);
					// JSONResponse();

					// JavascriptExecutor js1 =
					// (JavascriptExecutor)Automation.driver;
					// js.executeScript("return document.readyState").toString().equals("complete");
				}

				else if (controlName.equalsIgnoreCase("SPACE")) {
					WebHelper.robot.keyPress(KeyEvent.VK_SPACE);
					WebHelper.robot.keyRelease(KeyEvent.VK_SPACE);
				} else if (controlName.equalsIgnoreCase("ENTER")) {

					Thread.sleep(3000);

					WebHelper.robot.keyPress(KeyEvent.VK_ENTER);
					WebHelper.robot.keyRelease(KeyEvent.VK_ENTER);

					Thread.sleep(3000);

					// Minaakshi : 01-03-2019
				} else if (!controlName.equalsIgnoreCase("ENTER") && !controlName.equalsIgnoreCase("SPACE")
						&& !controlName.equalsIgnoreCase("TAB") && !controlName.equalsIgnoreCase("SetFilePath")) {
					WebElement webElementTAB = null;

					try {// Minaakshi : 01-04-2019

						if (controlId.equalsIgnoreCase("XPath")) {
							webElementTAB = Automation.driver.findElement(By.xpath(controlName));
						} else {
							webElementTAB = Automation.driver.findElement(By.name(controlName));
						}

						// getElementByType(controlId, controlName,
						// WebHelper.control, imageType, ctrlValue);
						if (webElementTAB.isEnabled()) {
							webElementTAB.sendKeys(Keys.TAB);
							Thread.sleep(1500);
						}
					} catch (Exception ex) {
						break;
					}
				}

				break;

			case DB:
				switch (actionName) {
				case Write:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					String policyNo = Automation.driver.findElement(By.xpath(controlName)).getText();
					ctrlValue = ctrlValue + "'" + policyNo + "'";
					ResultSet rs = null;
					Connection conn = JDBCConnection.establishDBConn();
					Statement st = conn.createStatement();
					rs = st.executeQuery(ctrlValue);

					rs.next();
					ctrlValue = String.valueOf(rs.getLong("COL_1"));
					rs.close();
					st.close();
					JDBCConnection.closeConnection(conn);
					WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo,
							colNo);

					break;
				case UpdateDBScript:
					log.info("logicalName----->" + logicalName);
					log.info("controlName----->" + controlName);

					Connection conn1 = JDBCConnection.establishPASDBConn();
					Statement st1 = conn1.createStatement();
					st1.execute(controlName);
					if (Config.databaseType.equalsIgnoreCase("ORACLE")) {
						st1.execute("commit");
					}
					st1.close();
					JDBCConnection.closeConnection(conn1);
					break;
				}
				break;

			case Pagination:// Minaakshi : 01-01-2020
				switch (actionName) {

				case I:

					log.info("In Pagination");
					log.info("In Pagination");
					if (ctrlValue == null || ctrlValue.trim().equals("") || ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}

					WebElement pgn = null;
					String tempAttribute = null;
					WebElement LastPage = null;
					String strLastPage = null;

					pgn = Automation.driver.findElement(By.xpath(controlName));
					// pgn =
					// dmgebtWait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
					log.info("Pagination object found.");
					log.info("Pagination object found.");

					Thread.sleep(5000);
					tempAttribute = pgn.getAttribute("class");

					if (tempAttribute.equalsIgnoreCase("last")) {
						log.info("Pagination Enabled");
						log.info("Pagination Enabled");
						strLastPage = controlName + "/a";

						LastPage = Automation.driver.findElement(By.xpath(strLastPage));
						// LastPage =
						// dmgebtWait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(strLastPage)));
						Thread.sleep(5000);
						LastPage.click();
						log.info("Last Page link has been clicked");
						log.info("Last Page link has been clicked");
						Thread.sleep(5000);
					} else {
						log.info("Pagination is not enabled");
						log.info("Pagination is not enabled");
					}
					break;
				}

				break;

			case Database:
		
				switch (actionName) {
				case CaptureDataFromDB:
					if (logicalName.equalsIgnoreCase("SQLQuery"))
						sqlQuery = ctrlValue;
					else if (logicalName.equalsIgnoreCase("ReadFromColName"))
						readFromColName = ctrlValue;
					else if (logicalName.equalsIgnoreCase("ColumnToBeFetchedFromDB"))
						toBeFetchedDBColName = ctrlValue;
					else if (logicalName.equalsIgnoreCase("WriteToColName")) {
						writeToColName = ctrlValue;
						String tempValue = "";
						String finalValue = "";
						String[] readfromArray;
						String[] tempReadFromArray;
						String[] tempMainReadFromArray;
						String[] writeToArray;
						String tempValue1 = "";
						// String tempValue2 = "";//Sel4
						int i;
						int j;
						// int k;//Sel4
						int l;
						// int m;//Sel4
						int n;
						if (!readFromColName.equalsIgnoreCase("")) {
							if ((readFromColName.contains(",")) && (readFromColName.contains("-"))
									&& (readFromColName.contains("*"))) {

								tempMainReadFromArray = readFromColName.split("\\*");
								l = 0;
								// m = tempMainReadFromArray.length - 1;//Sel4

								for (n = 0; n < tempMainReadFromArray.length; n++) {
									if (tempMainReadFromArray[n].contains(",")) {
										tempReadFromArray = tempMainReadFromArray[n].split(",");
										j = 0;
										// k = tempReadFromArray.length - 1;//Sel4
										for (i = 0; i < tempReadFromArray.length; i++) {
											readfromArray = tempReadFromArray[i].split("-");
											finalValue = DMProduct.ReadFromExcelUsingColumnName(readfromArray[0],
													readfromArray[1]);
											if (i == j) {
												tempValue1 = finalValue;
											}
											/*
											 * else if (i==k){ tempValue = tempValue + finalValue; }
											 */
											else {
												tempValue1 = tempValue1 + "," + finalValue;
											}
										}
									} else {

										readfromArray = tempMainReadFromArray[n].split("-");
										tempValue1 = DMProduct.ReadFromExcelUsingColumnName(readfromArray[0],
												readfromArray[1]);
									}

									if (n == l) {
										tempValue = tempValue1;
									} else {
										tempValue = tempValue + "*" + tempValue1;
									}
								}

							}

							// GRP_SC_Prod_EFFD_KPI-KPI_Code,GRP_SC_Prod_ISSD_KPI-KPI_Code,GRP_SC_Prem_KPI-KPI_Code
							else if ((readFromColName.contains(",")) && (readFromColName.contains("-"))
									&& !(readFromColName.contains("*"))) {// Minaakshi : 01-10-2020
								tempReadFromArray = readFromColName.split(",");
								j = 0;
								// k = tempReadFromArray.length - 1;//Sel4
								for (i = 0; i < tempReadFromArray.length; i++) {
									readfromArray = tempReadFromArray[i].split("-");
									finalValue = DMProduct.ReadFromExcelUsingColumnName(readfromArray[0],
											readfromArray[1]);
									if (i == j) {
										tempValue = finalValue;
									}
									/*
									 * else if (i==k){ tempValue = tempValue + finalValue; }
									 */
									else {
										tempValue = tempValue + "," + finalValue;
									}
								}
							} else if (readFromColName.contains("-")) {// Minaakshi
																		// :
																		// 01-04-2020
								readfromArray = readFromColName.split("-");
								tempValue = DMProduct.ReadFromExcelUsingColumnName(readfromArray[0], readfromArray[1]);
							} else {
								tempValue = DMProduct.ReadFromExcelUsingColumnName("", readFromColName);
							}

							sqlQuery = DMProduct.getDataFetchingSQLQuery(sqlQuery, tempValue);
						} // Minaakshi : 20-09-2018

						ResultSet rs2;
						Connection conn2 = JDBCConnection.establishPASDBConn();
						Statement st2 = conn2.createStatement();
						log.info("Updated dynamic sqlQuery : " + sqlQuery);
						log.info("Updated dynamic sqlQuery : " + sqlQuery);
						rs2 = st2.executeQuery(sqlQuery);
						rs2.next();

						ctrlValue = String.valueOf(rs2.getString(toBeFetchedDBColName));

						while (rs2.next()) {
							String tempctrlValue = String.valueOf(rs2.getString(toBeFetchedDBColName));
							ctrlValue = ctrlValue + "," + tempctrlValue;
						}

						rs2.close();
						st2.close();
						JDBCConnection.closeConnection(conn2);

						if (writeToColName.contains("-")) {
							writeToArray = writeToColName.split("-");

							DMProduct.writeDataToUniqueNumberSheet(ctrlValue, writeToArray[1], writeToArray[0]);
						} else if (writeToColName.contains("string*")) {
							String[] tempStr1;
							String tempStr2;
							tempStr1 = ctrlValue.split(",");
							tempStr2 = "'" + tempStr1[0] + "','" + tempStr1[1] + "','" + tempStr1[2] + "'";
							ctrlValue = tempStr2;

							String[] tempStr3;
							tempStr3 = writeToColName.split("\\*");

							DMProduct.writeDataToUniqueNumberSheet(ctrlValue, tempStr3[1], "");
						} else {
							DMProduct.writeDataToUniqueNumberSheet(ctrlValue, writeToColName, "");
						}
					}
					break;
				case UpdateDataFromDB:

					sqlQuery = ctrlValue;
					Connection conn1 = JDBCConnection.establishPASDBConn();
					Statement st1 = conn1.createStatement();
					st1.execute(sqlQuery);
					if (Config.databaseType.equalsIgnoreCase("ORACLE")) {
						st1.execute("commit");
					}
					st1.close();
					JDBCConnection.closeConnection(conn1);
					break;
				}
				break;
			case CreateDynamicData:
				log.info("CreateDynamicData");
				Jacob.main(ctrlValue, "!GenerateData_Click");
				// DynamicDataGeneration.xlsm!GenerateData_Click
				break;
			case FlatFile:// Minaakshi : 01-03-2019
				switch (actionName) {
				case I:

					DMProduct.CreateFlatFile(logicalName, ctrlValue);
					log.info("Testing 1 2 3");
					break;

				case Read:

					DMProduct.CreateFlatFile(logicalName, ctrlValue);

					break;

				case ReadTxtFile: // Chaitali : 06-17-2022

					if (logicalName.equalsIgnoreCase("ValidateTxtFile") && ctrlValue.equals("Yes")) {

						DMProduct.ReadTxtFile();

						break;
					}

					break;

				case CreateFile:

					DMProduct.CreateFlatFile(logicalName, ctrlValue);

					break;
				case CopyFilesToServer:
					if (logicalName.equalsIgnoreCase("FromLocalToServer") && ctrlValue.equals("Yes")) {
						/*
						 * source = Config.inputDataFilePath.toString() + Local_Path; destination =
						 * Remote_Path;
						 */
						DMProduct.CopyFilesToServer();
					}

					break;

				case CopyFilesToWindowsServer: // Minaakshi : 01-06-2020
					if (logicalName.equalsIgnoreCase("FromLocalToServer") && ctrlValue.equals("Yes")) {

						DMProduct.CopyFilesToWindowsServer();
						/*
						 * File src = new File(Config.inputDataFilePath.toString() + Local_Path + "\\" +
						 * FileNameToBeCreated); File dest = new
						 * File(Config.copyServerRemotePath.toString()+ "/" + FileNameToBeCreated);
						 * //DMProduct.CopyFilesToServer(source, destination, FileNameToBeCreated);
						 * 
						 * if (src.exists()) { FileUtils.copyFile(src, dest);
						 * 
						 * } else { webDriver.getReport().setMessage( "File has not been copied");
						 * log.error("File has not been copied"); throw new
						 * IOException("File has not been copied"); }
						 */
					}

					break;

				case CopyFilesToLocal:
					if (logicalName.equalsIgnoreCase("FromServerToLocal") && ctrlValue.equals("Yes")) {
						DMProduct.CopyFilesToLocal(); // Chaitali 06-17-2022
					}
					break;

				}

				break;

			case FileUpload_DM:// Minaakshi : 01-04-2019

				if (ctrlValue == null || ctrlValue.trim().equals("") || ctrlValue.equalsIgnoreCase("IGNORE")) {
					break;
				}

				String autoitFileDir_dm = Config.inputDataFilePath + "FileUpload";

				WebDriver driver = Automation.driver;
				WebElement element = driver.findElement(By.xpath(".//*[@class='btn btn-default btn-file']"));
				// browser button
				WebElement text = driver.findElement(By.xpath(".//*[@name='browseAndUpload']"));
				// text box of browse button
				String uploaded = "";
				int count = 0;
				while (uploaded.equals("")) {
					count++;
					try {
						uploaded = uploadFile(element, text, autoitFileDir_dm, ctrlValue);

						Runtime.getRuntime()
								.exec(autoitFileDir_dm + "\\FileUpload.exe " + autoitFileDir_dm + "\\" + ctrlValue);
						log.info("uploaded file name: " + uploaded);
					} catch (Exception e) {
						log.info("Trying to handle windows dialog box...");
						Runtime.getRuntime().exec(autoitFileDir_dm + ctrlValue);
						Thread.sleep(3000);
						try {
							text.getAttribute("value");
						} catch (Exception exp) {
							Runtime.getRuntime().exec(autoitFileDir_dm + ctrlValue);
							Thread.sleep(3000);
							break;
						}
					}

					if (count == 3) {
						break;
					}
				}
				log.info("File Upload Retry Counter is: " + count);

				try {
					String fileName = text.getAttribute("value");
					if (!fileName.equals(""))
						log.info("File uploaded successfully");
				} catch (Exception e) {
					log.info("File Upload Failed!!");
				}

				break;

			case CheckBatchStatus:// Minaakshi : 01-02-2020

				if (ctrlValue == null || ctrlValue.trim().equals("") || ctrlValue.equalsIgnoreCase("IGNORE")) {
					break;
				}

				String batchStatus = "";
				String SQLQuery = ctrlValue;

				ResultSet rs3;
				Connection conn3 = JDBCConnection.establishPASDBConn();

				Statement st3 = conn3.createStatement();

				// This loop is for waiting record to be arrived in // Minaakshi
				// : 01-01-2020
				startTime_batch = System.nanoTime();

				while ((System.nanoTime() - startTime_batch) < 100 * 60 * NANOSEC_PER_SEC) {
					rs3 = st3.executeQuery(SQLQuery);

					if (rs3 != null) {
						rs3.next();
					}

					if (rs3 != null) {
						log.info("wait not required in first loop");
						log.info("wait not required in first loop");
						rs3.close();
						st3.close();
						break;
					}
				}

				// This loop is for waiting status of record to be COMPLETED
				ResultSet rs4;
				Statement st4 = conn3.createStatement();

				startTime_batch = System.nanoTime();
				while ((System.nanoTime() - startTime_batch) < 50 * 60 * NANOSEC_PER_SEC) {

					log.info("in second wait loop : waiting for status of record to be COMPLETED/FAILED/TR");
					log.info("in second wait loop : waiting for status of record to be COMPLETED/FAILED/TR");

					rs4 = st4.executeQuery(SQLQuery);

					if (rs4 != null) {
						rs4.next();
						// if (rs4.next()){
						// while(rs4.next()){
						batchStatus = String.valueOf(rs4.getString("JOB_STATUS"));

						// if ((batchStatus.equalsIgnoreCase("COMPLETED")) ||
						// (batchStatus.equalsIgnoreCase("TR")) ||
						// (batchStatus.equalsIgnoreCase("FAILED"))) {
						if (controller.controllerTestCaseID.toString().contains("GRP_")) {
							if ((batchStatus.equalsIgnoreCase("COMPLETED"))) {
								log.info("JOB_STATUS : batchStatus");
								log.info("In Non Regression Batch Status Loop");
								log.info("JOB_STATUS : " + batchStatus);
								rs4.close();
								st4.close();
								JDBCConnection.closeConnection(conn3);
								break;
							}
						} else {
							if ((batchStatus.equalsIgnoreCase("COMPLETED")) || (batchStatus.equalsIgnoreCase("TR"))
									|| (batchStatus.equalsIgnoreCase("FAILED"))) {
								log.info("JOB_STATUS : batchStatus");
								log.info("In Regression Batch Status Loop");
								log.info("JOB_STATUS : " + batchStatus);
								rs4.close();
								st4.close();
								JDBCConnection.closeConnection(conn3);
								break;
							}
						}
						// }
					}
				}

				break;

			case WaitForEC:
				WebHelper.wait.until(CommonExpectedConditions.elementToBeClickable(webElement));
				break;
			// bhaskar
			// case SikuliRun:
			// App.open(controlName);
			case SikuliScreen:
				App.open(WebHelper.sikscreen);
				break;
			// bhaskar
			case SikuliType:
				log.info("in sikulitype");
				log.info("controlName is:" + controlName);
				Pattern image1 = new Pattern(controlName);
				sikuliScreen.type(image1, ctrlValue);
				break;

			case SikuliButton:
				log.info("in sikuliButton");
				log.info("controlName is:" + controlName);
				Pattern image2 = new Pattern(controlName);
				sikuliScreen.click(image2);
				log.info("Done");
				break;
			// bhaskar
			case Slider:
				WebElement slider = Automation.driver.findElement(By.xpath(controlName));
				Thread.sleep(3000);
				Actions moveSlider = new Actions(Automation.driver);
				Action actionslider = moveSlider.dragAndDropBy(slider, 30, 0).build();
				actionslider.perform();
				break;
			// bhaskar
			case ifExist:

				try {

					if (Automation.driver.findElement(By.xpath(controlName)).isDisplayed()) {

						webElement.clear();
						webElement.click();
						webElement.sendKeys(ctrlValue);

					}

				} catch (Exception e) {
					log.error(e.getMessage(), e);
					log.info("Field does not Exist");
				}

				break;

			case MaskedInputDate:
				if (!ctrlValue.equalsIgnoreCase("null")) // changed by asif
				{

					webElement.sendKeys(Keys.ENTER);

					js.executeScript("arguments[0].value = '';", webElement);

					webElement.sendKeys(ctrlValue);

				} else {
					log.info("Don't Perform any action");
				}
				break;

			case Date:
				Calendar cal = new GregorianCalendar();
				int i = cal.get(Calendar.DAY_OF_MONTH);
				if (i >= 31) {
					i = i - 10;
				}
				break;

			case FileUpload:
				webElement.sendKeys(ctrlValue);
				break;

			case ScrollTo:
				if (Config.projectName.equals("DistributionManagement")) {// Minaakshi
					switch (actionName) {
					case I:
						if (ctrlValue == null || ctrlValue.trim().equals("")) {
							break;
						}
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
									webElement);

						}
					}
					break;
				} else {
					try {
						if (webElement.isDisplayed()) {
							Coordinates coordinate = ((Locatable) webElement).getCoordinates();
							coordinate.onPage();
							coordinate.inViewPort();
						}
					} catch (Exception e) {
						log.error(e.getMessage(), e);
						log.info("Element does not exist");
					}
				}

				break;

			case ScrollToElement: // Sheetal-2-5-2019. Had to create new
				// due to
				// project specific code in 'ScrollTo'
				switch (actionName) {
				case I:
					if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
							|| !ctrlValue.equalsIgnoreCase("")) {
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								webElement);

					}
					break;
				case NC:
					((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
							webElement);

					break;

				case NCIF:
					try {
						if (webElement.isDisplayed() == true) {
							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
									webElement);

							break;
						}
					} catch (Exception e) {

						log.info("Element Not Found on Page.");
					}
					break;
				}
				break;

			case WebService:
				switch (actionName) {
				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					if (logicalName.equalsIgnoreCase("WSDL_URL"))
						WebHelper.wsdl_url = ctrlValue;
					else if (logicalName.equalsIgnoreCase("REQUEST_URL"))
						WebHelper.request_url = ctrlValue;
					else if (logicalName.equalsIgnoreCase("REQUEST_XML"))
						WebHelper.request_xml = ctrlValue;
					break;

				case CreateFile:
					DMProduct.CreateWebServiceRequest(logicalName, ctrlValue);
					break;

				case T:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					} else
						;

					// String username= Config.basicAuthUsername;
					// String password = Config.basicAuthPassword;
					// String wscycledate = null;
					String fileName = Config.inputDataFilePath.toString() + "WebService\\WebserviceFiles\\"
							+ WebHelper.request_xml + ".xml";
					// WebHelper.request_xml=fileName;

					WebService.callWebService("", WebHelper.wsdl_url, fileName, Config.basicAuthUsername,
							Config.basicAuthPassword);
					break;

				case V:
					if (ctrlValue == null || ctrlValue.trim().equals("") || ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					currentValue = WebService.getXMLTagValue(controlName);
					break;
				}
				break;

			// Added by Rajesh M to calling other framework jar from Policy
			// automation framework
			case ExecuteJar:
				switch (actionName) {
				case Billing:
					log.info("logicalName----->" + logicalName);
					log.info("controlName----->" + controlName);
					if (controlName != null) {
						WebHelperUtil.executeJar(controlName);
					} else {
						log.info("logicalName should not be blank or null");
					}
					break;
				}
				break;

			case SaveDocument:
				String[] temp = ctrlValue.split(";");

				String downloadedFileName2 = Config.actualPdfDownloadPath + "\\" + temp[0];
				File file1 = new File(downloadedFileName2);

				if (file1.exists()) {
					file1.delete();
					log.info("Old file DocumentPackaging.pdf deleted successfully");
				}

				// Below click downloads the document
				Actions MousebuilderClick1 = new Actions(Automation.driver);
				highlightElement(webElement);
				Action MouseclickAction1 = MousebuilderClick1.moveToElement(webElement).clickAndHold().release()
						.build();

				MouseclickAction1.perform();

				break;

			case DownloadDocument:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case WaitTillFileDownload:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case MoveDocument:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case RenameDocument:

				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case CopyFromServerLocation:

				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case IgnoreString:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case PDFComparisonMode:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case PDFDocumentCompare:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case OpenAPI:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, transactionType.toString());
			default:
				log.info("U r in Default");
				break;
			}

			// To handle processing icon in DM
			try {
				Thread.sleep(0);
				log.info("Waited for 0 second");
				for (int i = 1; i <= 340; i++) {
					List<WebElement> loader = Automation.driver
							.findElements(By.xpath("//body[contains(@class,'waiting')]"));

					int loaderSize = loader.size();
					// log.info("loaderSize is-->" + loaderSize);
					log.info("loaderSize is-->" + loaderSize);
					if (loaderSize >= 1)
						Thread.sleep(400);
					else
						break;
				}
			} catch (Exception e) {
			}

		} catch (WebDriverException we) {
			log.error(we.getMessage(), we);
			throw new Exception("Error Occurred from Do Action " + controlName + we.getMessage());
		} catch (Exception e) {
			String errorMessage = e.getMessage();
			String firstLine = errorMessage.split("\\r?\\n")[0];
			log.error("Error :" + firstLine);
			//log.error(e.getMessage(), e);//Sel4
			throw new Exception(e.getMessage());
		}

		// Mrinmayee - skip the step for which control value is blank end

		// TM-02/02/2015: Radio button found ("F") & AJAX control ("VA")
		if ((action.equalsIgnoreCase("V") || action.equalsIgnoreCase("F") || action.equalsIgnoreCase("PopUpCount")
				|| action.equalsIgnoreCase("VerifyPopUpElement") || action.equalsIgnoreCase("VA")
				|| action.toString().equalsIgnoreCase("VD") || action.toString().equalsIgnoreCase("V_Image")
				|| action.toString().equalsIgnoreCase("V_text")
				|| action.toString().equalsIgnoreCase("V_availabilityStatus")

				|| action.toString().equalsIgnoreCase("V_attributeValue")
				|| action.toString().equalsIgnoreCase("VFCount") || action.toString().equalsIgnoreCase("VFPresence")
				|| action.toString().equalsIgnoreCase("V_checkboxStatus")
				|| action.toString().equalsIgnoreCase("V_edit") || action.toString().equalsIgnoreCase("V_disableStatus")
				|| action.toString().equalsIgnoreCase("VFormPresence")
				|| action.toString().equalsIgnoreCase("VFormCount")) && !ctrlValue.equalsIgnoreCase("")
				&& !ctrlValue.equalsIgnoreCase("IGNORE") || action.equalsIgnoreCase("FormVerificationDetails")) // :
		{ // 03-10-2018
			if (Results == true) {
				webDriver.setReport(WriteToDetailResults(ctrlValue, currentValue, logicalName));
			}
		}

		return currentValue;
	}

	/** This Functions Waits for the HTMLPage To Load **/
	public static Boolean waitForCondition() throws IOException {
		ExpectedCondition<Boolean> expCondition = null;

		try {
			// wait = new
			// WebDriverWait(Automation.driver,Integer.parseInt(Automation.timeOut.toString()));//Integer.parseInt(Automation.timeOut.toString())
			expCondition = new ExpectedCondition<Boolean>() {

				@Override
				public Boolean apply(WebDriver driver) {
					return ((JavascriptExecutor) driver).executeScript("return document.readyState").equals("complete");
				}
			};
		} catch (WebDriverException e) {
			log.error(e.getMessage(), e);
			controller.pauseFun("From PageLoaded Function" + e.getMessage());
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			controller.pauseFun("Timed Out after waiting");
		}
		@SuppressWarnings("unused")
		WebDriverWait wait1 = new WebDriverWait(Automation.driver, (Duration.ofSeconds(30)));
		return WebHelper.wait.until(expCondition);
	}

	public static List<Object> getPropertiesOfWebElement(WebElement webElement, String imageType) {
		List<WebElement> elements = webElement.findElements(By.tagName(imageType));
		WebElement element = elements.get(0);
		List<Object> elementProperties = new ArrayList<Object>();
		String elementType = element.getAttribute("type");
		String elementTagName = element.getTagName();
		// String elementClassName = element.getClass().toString();
		String action = "";
		String controlType = "";
		String id = "";
		String name = "";
		String controlName = "";
		String controlID = "";
		if (elementType.equals("text") && elementTagName.equals("input")) {
			id = element.getAttribute("id");
			name = element.getAttribute("name");
			controlType = "WebEdit";
			controlID = "Id";
			// controlName="//"+elementTagName+"[@id="+id+"]";
			controlName = id;

			// action="I";

		} else if (elementType.contains("checkbox") && elementTagName.equals("input")) {
			id = element.getAttribute("id");
			name = element.getAttribute("name");
			controlType = "CheckBox";
		} else if (elementType.contains("listbox") && elementTagName.equals("select")) {
			id = element.getAttribute("id");
			name = element.getAttribute("name");
			controlType = "WebList";
		} else if (elementType.contains("radio") && elementTagName.equals("input")) {
			id = element.getAttribute("id");
			name = element.getAttribute("name");
			controlType = "Radio";
		}

		else if (elementType.contains("") && elementTagName.equals("a")) {
			id = element.getAttribute("id");
			name = element.getAttribute("name");
			controlType = "WebLink";
		} else if (elementType.contains("") && elementTagName.equals("")) {
			id = element.getAttribute("id");
			name = element.getAttribute("name");
			controlType = "WebEdit";
			controlID = "Id";
			controlName = id;

		}
		elementProperties.add(action);
		elementProperties.add(id);
		elementProperties.add(name);
		elementProperties.add(controlName);
		elementProperties.add(controlID);
		elementProperties.add(controlType);
		elementProperties.add((Object) element);

		for (int i = 0; i < 5; i++) {
			log.info(elementProperties.get(i).toString());
		}
		return elementProperties;
	}

	public static void saveScreenShot(String controlType) {
		if (ITAFWebDriver.isSuiteApplication()) {
			WebHelperUtil.saveScreenShot("");
		} else {
			if (!(Automation.driver instanceof TakesScreenshot)) {
				log.info(
						"Not able to take screenshot: Current WebDriver does not support TakesScreenshot interface.");
				return;
			}

			File scrFile;
			try {
				scrFile = ((TakesScreenshot) Automation.driver).getScreenshotAs(OutputType.FILE);
			} catch (Exception e) {
				log.error(e.getMessage(), e);
				log.info("Taking screenshot failed for: " + webDriver.getReport().getTestcaseId());
				return;
			}

			String date = null;

			if (StringUtils.isNotBlank(webDriver.getReport().getFromDate())
					&& !controlType.equalsIgnoreCase("Screenshot")) {
				date = webDriver.getReport().getFromDate().replaceAll("[-/:]", "");
				date = date.replace(" ", "_");
			} else {
				webDriver.getReport().setFromDate(Config.dtFormat.format(new Date()));
				date = webDriver.getReport().getFromDate().replaceAll("[-/:]", "");
				date = date.replace(" ", "_");
			}

			if (webDriver.getReport().getTrasactionType() == null || webDriver.getReport().getTrasactionType() == "") {
				webDriver.getReport().setTrasactionType(controller.controllerTransactionType.toString());
			}
			if (webDriver.getReport().getTestcaseId() == null || webDriver.getReport().getTestcaseId() == "") {
				webDriver.getReport().setTestcaseId(controller.controllerTestCaseID.toString());
			}
			String fileName = webDriver.getReport().getTestcaseId() + "_" + webDriver.getReport().getTrasactionType()
					+ "_" + date;

			String location = "";
			if (controlType.equalsIgnoreCase("Screenshot")) {
				try {
					String subfolder = Config.resultFilePath + "\\ScreenShots\\SelectiveScreenshots";
					File createDir = new File(subfolder);
					if (!createDir.exists()) {
						createDir.mkdirs();
					}
					location = subfolder + "\\" + fileName + ".jpeg";

					FileUtils.copyFile(scrFile, new File(location));
				} catch (IOException e) {
					log.error(e.getMessage(), e);
					e.printStackTrace();
					return;
				}
			} else {
				location = Config.resultFilePath + "\\ScreenShots\\" + fileName + ".jpeg";
				// controller.testDescription = location;
				webDriver.getReport().setScreenShot("file:\\\\" + location);
				try {
					FileUtils.copyFile(scrFile, new File(location));
					controller.FailScreen = new File(location).getAbsolutePath();
				} catch (IOException e) {
					log.error(e.getMessage(), e);
					e.printStackTrace();
					return;
				}
			}
		}
	}

	public static void JSONResponse() {

		try {

			for (int i = 0; i < 2; i++) {

				Thread.sleep(50);

				// long start = System.currentTimeMillis();
				// URL obj = new
				// URL(Automation.driver.getCurrentUrl().toString());
				// URL obj=new URL("https://127.0.0.1") ;
				URL obj = new URL("https://mic99.cover-all.com/mic/pctv2/pctentrypoint/");
				HttpsURLConnection conn = (HttpsURLConnection) obj.openConnection();

				// getFullURL(ServletRequest request);
				conn.setRequestMethod("POST");

				conn.connect();

				log.info(conn.getConnectTimeout());

				// int header = conn.getResponseCode();//Sel4

				switch (conn.getResponseCode()) {
				case HttpsURLConnection.HTTP_OK:
					log.info("200OK");

					// Thread.sleep(200);

					break; // fine, go on
				case HttpsURLConnection.HTTP_GATEWAY_TIMEOUT:
					System.out.print("TimOut");
					break;// retry
				case HttpsURLConnection.HTTP_UNAVAILABLE:
					System.out.print("Unavailable");
					break;// retry, server is unstable
				default:
					System.out.print("UnknownResposnecode");
					break; // abort
				}

			}

		} catch (Exception e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
		}

	}

	/**
	 * Auth : DW Date : 02-Jan-2018 This function is used to highlight element on
	 * the html screen
	 **/
	public static void highlightElement(WebElement element) {
		try {
			((JavascriptExecutor) Automation.driver).executeScript("arguments[0].setAttribute('style', arguments[1]);",
					element, "border: 3px groove lime;");
			Thread.sleep(300);
			((JavascriptExecutor) Automation.driver).executeScript("arguments[0].setAttribute('style', arguments[1]);",
					element, "");

		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}

	public static void waitForVaadin() {
		StringBuilder commandBuilder = new StringBuilder(500);
		commandBuilder.append("if (window.vaadin == null) { return 4; }");
		commandBuilder.append("var clients = window.vaadin.clients;");
		commandBuilder.append("if (!clients) { return 3; }");
		commandBuilder.append("for (var key in clients) {");
		commandBuilder.append(" var client = clients[key];");
		commandBuilder.append(" if (client.getElementsByPath == undefined) { return 2; }");
		commandBuilder.append(" else if (client.isActive()) { return 1; } }");
		commandBuilder.append("return 0;");
		String command = commandBuilder.toString();
		long startTime = System.currentTimeMillis();
		long maxTime = startTime + 120000;
		JavascriptExecutor js = (JavascriptExecutor) Automation.driver;
		long errorCode = -1;
		while (System.currentTimeMillis() < maxTime && errorCode != 0) {
			errorCode = (Long) js.executeScript(command);
		}
		if (errorCode != 0) {
			log.info("Timeout after waiting for Vaadin call to finish.");
		}
	}

	// Minaakshi : 01-03-2019
	private static String uploadFile(WebElement element, WebElement text, String autoitFileDir, String ctrlValue)
			throws InterruptedException, IOException {

		element.click();
		Thread.sleep(2000);
		Runtime.getRuntime().exec(autoitFileDir + "\\FileUpload.exe " + autoitFileDir + "\\" + ctrlValue);
		// Below Thread.sleep is required once windows browse modal populated
		// with file path should be opened that takes sometime.
		Thread.sleep(5000);
		log.info("Inside uploadFile() called");
		return text.getAttribute("value");
	}
}
