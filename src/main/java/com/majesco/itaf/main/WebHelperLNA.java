package com.majesco.itaf.main;

import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.PrintStream;
import java.io.UnsupportedEncodingException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Scanner;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import javax.net.ssl.HttpsURLConnection;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;
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
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.common.io.Files;
import com.majesco.itaf.util.CalendarSnippet;
import com.majesco.itaf.util.CommonExpectedConditions;
import com.majesco.itaf.util.Constants;
import com.majesco.itaf.util.DMProduct;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.util.JDBCConnection;
import com.majesco.itaf.util.Jacob;
import com.majesco.itaf.util.PDFComparisonUtil;
import com.majesco.itaf.util.WaitTool;
import com.majesco.itaf.verification.WebVerification;
import com.majesco.itaf.vo.Reporter;
import com.majesco.itaf.webservice.WebService;
import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;

public class WebHelperLNA {
	private final static Logger log = LogManager.getLogger(WebHelperLNA.class.getName());
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
	public static String sqlQuery = "";
	public static String readFromColName = "";
	public static String toBeFetchedDBColName = "";
	public static String expectedDBStatus = "";
	public static String writeToColName = "";
	public static WebElement groupid;
	public static Wait<WebDriver> waitForElementPresence; 
	public static Wait<WebDriver> dmgebtWait;
	public static String Local_Path = "";
	public static String Template_Path = "";
	public static String Remote_Path = "";
	public static String Remote_Path_Outbound = "";
	public static String EffectiveDate = "";
	public static String EffectiveDate1 = "";
	public static String RiskCommencementDate = "";
	public static String RiskCommencementDate1 = "";
	public static String PolicyEffectiveToDate = "";
	public static String ProposalApplDate = "";
	public static String ProposalSubmDate = "";
	public static String PolicyNumber; 
	public static String PolicyFileLoadNumber;
	public static String PremiumFileLoadNumber; 
	public static String EntityCode1 = "";
	public static String EntityCode2 = "";
	public static String EntityCode3 = "";
	public static String EntityCode4 = "";
	public static String EntityCode5 = "";
	public static String TemplateFileName = "";
	public static String FileNameToBeCreated = "";
	public static String CreateFile = "";
	public static String FromLocalToServer = "";
	public static String FromServerToLocal = "";
	public static String FileNameCopiedToLocal = "";
	public static File copySource;
	public static File copyDestination;
	public static String source;
	public static String destination;
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

	@SuppressWarnings("unused")
	public static void GetCellInfo(String FilePath, Row rowValues, int valuesRowIndex, int valuesRowCount)
			throws IOException
	{
		try {

			WebHelper.frmDate = new Date();
			WebHelper.isDynamicNumFound = true;
			List<WebElement> controlList = null;
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
			int rowCount = sheetStructure.getLastRowNum() + 1;
			Sheet headerValues = ExcelUtility.GetSheet(FilePath, "Values");
			String fromDate = Config.dtFormat.format(WebHelper.frmDate);
			webDriver.getReport().setFromDate(fromDate);
			WebHelper.structureHeader = WebHelperUtil.getValueFromHashMap(sheetStructure);
			WebHelper.columnName = null;
			int dynamicIndexNumber;// Added for Action Loop
			String imageType, indexVal, controlName, executeFlag, action, logicalName, controltype, controlID,
					dynamicIndex, newDynamicIndex, rowNo, colNo, columnName1, CompareText;// newly
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
					System.out.println("iTAF:" + logicalName + " " + rowIndex);
					log.info("iTAF:" + logicalName + " " + rowIndex);

					/*
					 * Below code has been written To Handle condition for multiple rows in excel
					 * sheet
					 */

					if (imageType.equalsIgnoreCase("START_SCRIPT")) {

						try {
							webElement = getElementByType(controlID, controlName, WebHelper.control, imageType,
									ctrlValue);
							if (webElement.isDisplayed()) {
								System.out.println("Proceed");
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
								|| (action.equals("V_edit") && !ctrlValue.isEmpty())
								|| (action.equals("V_disableStatus") && !ctrlValue.isEmpty())
								|| (action.equals("V_checkboxStatus") && !ctrlValue.isEmpty())
								|| (action.equals("VFPresence") && !ctrlValue.isEmpty())
								|| (action.equals("VFormPresence") && !ctrlValue.isEmpty())
								|| (action.equals("InputifExist") && !ctrlValue.isEmpty())
								|| (action.equals("VFormCount") && !ctrlValue.isEmpty())
								|| (!action.equals("I") && !action.equals("V") && !action.equals("V_availabilityStatus")
										&& !action.equals("V_edit") && !action.equals("V_text")
										&& !action.equals("V_disableStatus") && !action.equals("VFPresence")
										&& !action.equals("V_checkboxStatus") && !action.equals("VFormCount")
										&& !action.equals("VerifyPopUpElement") && !action.equals("ActionClick"))) {
							if (controltype.equalsIgnoreCase("PDFDocumentCompare")
									|| controltype.equalsIgnoreCase("IgnoreString")) {
								if (controltype.equalsIgnoreCase("IgnoreString")) {
									PDFComparisonUtil.setStringToIgnore(ctrlValue);
								}
								if (controltype.equalsIgnoreCase("PDFDocumentCompare")) {
									if (!ctrlValue.isEmpty()) {
										Object[] result = PDFComparisonUtil.PDFCompare(logicalName, controltype,
												ctrlValue);
										webDriver.report = WriteToDetailResults(result[1].toString(),
												result[2].toString(), logicalName);
									}
								}

							} else {
								if (!controltype.startsWith("Sikuli")) {
									if (!action.equalsIgnoreCase("LOOP") && !controltype.equalsIgnoreCase("Wait")
											&& !controltype.equalsIgnoreCase("Wait_DM")
											// Minaakshi : 05-02-2019
											&& !controltype.equalsIgnoreCase("OpenAPI")// Minaakshi
											// :
											// 01-06-2019
											&& !action.equalsIgnoreCase("END_LOOP")
											// added to skip webElment creation suhas 08/14/2019 ..Started
											&& !controltype.equalsIgnoreCase("Opt_WebElment")
											&& !controltype.equalsIgnoreCase("spWait")
											&& !controltype.equalsIgnoreCase("CurrentDate")
											&& !controltype.equalsIgnoreCase("CreateJSON_Claims")
											&& !controltype.equalsIgnoreCase("WebService_Claims")

											&& !controltype.equalsIgnoreCase("Edit_Enrollment")

											&& !controltype.equalsIgnoreCase("BatchStatus")

											&& !controltype.equalsIgnoreCase("logoutLogic")
											&& !controltype.equalsIgnoreCase("screenshot")/// added temp
											&& !controltype.equalsIgnoreCase("Back_Date")
											&& !controltype.equalsIgnoreCase("Opt_CurrentDate")
											&& !controltype.equalsIgnoreCase("Dental_ClaimStatus")
											&& !controltype.equalsIgnoreCase("RunExe")
											&& !action.equalsIgnoreCase("V_opt")
											&& !controltype.equalsIgnoreCase("TaskStatus_GUW")
											&& !controltype.equalsIgnoreCase("TaskStatus_LTD")
											&& !controltype.equalsIgnoreCase("TaskStatus_WOP")
											&& !controltype.equalsIgnoreCase("TaskStatus")
											&& !controltype.equalsIgnoreCase("Opt_scroll")
											&& !controltype.equalsIgnoreCase("TestCaseUP")
											// added to skip webElment creation suhas 08/14/2019 ..Finished
											&& !controltype.equalsIgnoreCase("Browser")
											&& !controltype.equalsIgnoreCase("CloseBrowser")
											&& !controltype.equalsIgnoreCase("AttributeisPresent")
											&& !controltype.equalsIgnoreCase("CloseAndLaunchNewBrowser")
											&& !controltype.equalsIgnoreCase("PageRefresh")
											&& !controltype.equalsIgnoreCase("WaitForPageToLoad")
											&& !controltype.equalsIgnoreCase("WaitUntilElementInvisible")
											&& !controltype.equalsIgnoreCase("filedownload")
											&& !controltype.equalsIgnoreCase("NewBrowser")
											&& !controltype.equalsIgnoreCase("Window")
											&& !controltype.equalsIgnoreCase("Alert")
											&& !controltype.equalsIgnoreCase("URL")
											&& !controltype.equalsIgnoreCase("WaitForJS")
											&& !controltype.contains("Robot")
											&& !controltype.equalsIgnoreCase("Calendar")
											&& !controltype.equalsIgnoreCase("CalendarNew")
											&& !controltype.equalsIgnoreCase("CalendarIPF")
											&& !controltype.equalsIgnoreCase("CalendarEBP")
											&& (!action.equalsIgnoreCase("Read") || ((action.equalsIgnoreCase("Read")
													&& !StringUtils.isEmpty(controlName)
													&& !ctrlValue.equalsIgnoreCase("IGNORE"))))// Minaakshi
											&& !controltype.equalsIgnoreCase("JSScript")
											&& !controltype.equalsIgnoreCase("DB")
											&& !controltype.equalsIgnoreCase("DB_UpdateXML") // Suhas : 08-19-2019
											&& !controltype.equalsIgnoreCase("Database")// Minaakshi
											&& !controltype.equalsIgnoreCase("CreateDynamicData")
											// Minaakshi : 03-10-2018
											&& !controltype.equalsIgnoreCase("FlatFile")
											&& !controltype.equalsIgnoreCase("FileUpload_DM")// Minaakshi
											&& !controltype.equalsIgnoreCase("CheckBatchStatus")// Minaakshi
											// :
											// 01-05-2019
											// :
											// 01-04-2019
											&& !controlID.equalsIgnoreCase("XML") && !controltype.startsWith("Process")
											&& !controltype.startsWith("Destroy")
											&& !controltype.startsWith("ReadSikuli")
											&& !controltype.equalsIgnoreCase("WebService")
											&& !controltype.equalsIgnoreCase("WebService_LNA")
											&& !action.equalsIgnoreCase("VA")
											&& (ctrlValue == null || !ctrlValue.equalsIgnoreCase("IGNORE"))// Minaakshi
											&& !controltype.equalsIgnoreCase("IFrame")
											&& !action.equalsIgnoreCase("LOOP")
											&& !controltype.equalsIgnoreCase("RowNumbersToExecute")
											&& !controltype.equalsIgnoreCase("WaitForElementToVisible")
											&& !controltype.equalsIgnoreCase("Wait_IfValue")
											&& !action.equalsIgnoreCase("Billing")
											&& !controltype.equalsIgnoreCase("D_Wait")
											&& !action.equalsIgnoreCase("UpdateDBScript")
											&& !action.equalsIgnoreCase("VFormPresence")
											&& !action.equalsIgnoreCase("FormVerificationDetails")
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
						}

						/*** Perform action on the identified control ***/
						// added by sheetal for new control type:
						// XPathValueMultiWithInput
						if (controlID.equalsIgnoreCase("XPathValueMultiWithInput")) {
							String tempcontrolValue = ctrlValue;
							String[] tempValues2 = tempcontrolValue.split(";;");
							int l = tempValues2.length;
							ctrlValue = tempValues2[l - 1];
						}

						if (!(controltype.equalsIgnoreCase("PDFDocumentCompare")
								|| controltype.equalsIgnoreCase("IgnoreString"))) {
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
					System.out.println("ExecuteFlag is N");
					log.info("ExecuteFlag is N");
				}
			}

			// Setting of reporting values after execution in case of no
			// exception
			Date toDate = new Date();
			webDriver.getReport().setFromDate(Config.dtFormat.format(WebHelper.frmDate));
			webDriver.getReport().setIteration(Config.cycleNumber);
			webDriver.getReport().setTestcaseId(controller.controllerTestCaseID.toString());
			// webDriver.getReport().setGroupName(controller.controllerGroupName.toString());
			webDriver.getReport().setTrasactionType(controller.controllerTransactionType.toString());
			webDriver.getReport().setTestDescription(controller.testDescription);
			webDriver.getReport().setToDate(Config.dtFormat.format(toDate));

			// Setting status for field verification failures
			if (WebHelper.fieldVerFailCount > 0) {
				webDriver.getReport().setMessage("Check Detailed Results");
				webDriver.getReport().setStatus("FAIL");
			}
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			controller.pauseFun(e.getMessage());
		} finally {
			WebHelper.structureHeader.clear();
			WebHelper.valuesHeader.clear();
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
		// Quote ID for PAS report requirement
		if (controller.controllerQuoteId != null) {
			report.setStrQuoteId(controller.controllerQuoteId.toString());
		} else {
			report.setStrQuoteId("");
		}

		if ((expectedValue.trim()).equalsIgnoreCase(actualValue.trim())) {
			report.setActualValue(actualValue);
			report.setExpectedValue(expectedValue);
			report.setStatus("PASS");
			report.setToDate(Config.dtFormat.format(WebHelper.frmDate));
			passCount = "1";
			failCount = "0";
			TotalpassCount += 1;
			System.out.println(TotalpassCount);
		} else {
			report.setActualValue("FAIL|" + actualValue + "|" + expectedValue);
			System.out.println("ab");
			report.setExpectedValue("Value");
			// report.setExpectedValue(expectedValue);
			report.setStatus("FAIL");
			report.setToDate(Config.dtFormat.format(WebHelper.frmDate));
			failCount = "1";
			passCount = "0";
			// DS:30-05-2014
			WebHelper.fieldVerFailCount += 1;
			TotalfailCount += 1;
			System.out.println(TotalfailCount);
		}

		WebHelper.print = new PrintStream(new FileOutputStream(WebHelper.file, true));

		int usedRows = WebHelperUtil.count(WebHelper.file);
		if (usedRows == 0) {
			WebHelper.print.print(
					"Iteration,TestCaseID,TransactionType,CurrentDate,RowType,Status,PassCount,FailCount, Compare Result,Quote Number");
			WebHelper.print.println();
		}
		usedRows = WebHelperUtil.count(WebHelper.file);

		WebHelper.print.print(ExcelUtility.myChar + Config.cycleNumber + ExcelUtility.myChar + "," + ExcelUtility.myChar
				+ report.getTestcaseId() + ExcelUtility.myChar + "," + ExcelUtility.myChar + report.getTrasactionType()
				+ ExcelUtility.myChar + "," + ExcelUtility.myChar + report.getToDate() + ExcelUtility.myChar + ","
				+ ExcelUtility.myChar + "Field: " + columnName + ExcelUtility.myChar + "," + ExcelUtility.myChar
				+ report.getStatus() + ExcelUtility.myChar + "," + ExcelUtility.myChar + passCount + ExcelUtility.myChar
				+ "," + ExcelUtility.myChar + failCount + ExcelUtility.myChar + "," + ExcelUtility.myChar
				+ report.getActualValue() + ExcelUtility.myChar + "," + ExcelUtility.myChar + report.getStrQuoteId()
				+ ExcelUtility.myChar);
		WebHelper.print.println();
		DResult++;
		return report;
	}

	/** Locating Web Element **/
	@SuppressWarnings("incomplete-switch")
	public static WebElement getElementByType(String controlId, String controlName, String controlType,
			String imageType, String controlValue) throws Exception {
		// String a=controlValue;

		// code added to check controlName has + by Suhas (L&A) 08/14/2019
		if (controlName.contains("+")) {

			controlName = controlName.replace("+", controlValue);

		}

		// WebDriverWait wait = new WebDriverWait(Automation.driver,
		// Integer.parseInt(Config.timeOut));//Sel4
		WebDriverWait wait = new WebDriverWait(Automation.driver, Duration.ofSeconds(Integer.parseInt(Config.timeOut)));
		if (Config.projectName.equals("DistributionManagement")) // Minaakshi :
		// 05-02-2019
		{
			WebHelperLNA.implementWait();
		}
		WebElement controlList = null;
		Constants.ControlIdEnum controlID = Constants.ControlIdEnum.valueOf(controlId);
		// WebDriverWait wait2 = new WebDriverWait(Automation.driver, 2);//Sel4
		WebDriverWait wait2 = new WebDriverWait(Automation.driver, Duration.ofSeconds(2));

		// System.out.println("bhaski controlID:"+controlID);
		try {
			switch (controlID) {
			case doNothing:
				break;
			case PageToLoad:

			case Id:

				// WebHelper.getText(Automation.driver,Automation.driver.findElement(By.xpath("")));
			case HTMLID:
				// System.out.println(Automation.driver.getCurrentUrl());

				controlList = wait.until(ExpectedConditions.elementToBeClickable(By.id(controlName)));
				break;

			case XPath_ctrvalue:

			case XPath:

				// JavascriptExecutor js =
				// (JavascriptExecutor)Automation.driver;
				// added another condition to check LNA product as it was gonig in else and
				// failing
				if (Config.projectName.equals("DistributionManagement")) {// Minaakshi
					// :
					// 05-02-2019
					try {

						Thread.sleep(1000);
						controlList = dmgebtWait
								.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

					} catch (Exception ex) {
						log.error(ex.getMessage(), ex);
						log.info("In Second Wait : XPATH : Element is not clickable : " + controlName);
						controlList = waitForElementPresence
								.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(controlName)));
					}
				}

				else {

					Thread.sleep(1000);

					controlList = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
				}

				break;
			case XPathValue: // Minaakshi : 10-09-2018
				String tempCtrlName = controlName;
				String tempReplaceString = "";

				if (controlValue.contains("|")) {// Minaakshi : 01-05-2019
					String[] temp = controlValue.split("\\|");
					tempReplaceString = tempCtrlName.replace("$value", temp[0].toString());
				} else {
					tempReplaceString = tempCtrlName.replace("$value", controlValue);
				}

				if (Config.projectName.equals("DistributionManagement")) {
					try {// Minaakshi : 05-02-2019
						controlList = dmgebtWait
								.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString)));
					} catch (Exception ex) {
						log.error(ex.getMessage(), ex);
						log.info("In Second Wait : XPATH : Element is not visible : " + controlName);
						controlList = waitForElementPresence
								.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(tempReplaceString)));
						// controlList =
						// waitForElementPresence.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(controlName)));
					}
				} else {
					controlList = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString)));
				}
				break;

			case XPathValue_dy: // Minaakshi : 01-03-2019

				String tempCtrlValue = "";

				if (controlValue.contains("|")) {
					String[] temp = controlValue.split("\\|");
					tempCtrlValue = DMProduct.ReadFromExcelUsingColumnName(temp[0].toString(), WebHelper.columnName);
				} else {
					tempCtrlValue = DMProduct.ReadFromExcelUsingColumnName(controlValue, WebHelper.columnName);
				}

				String tempCtrlName5 = controlName;
				String tempReplaceString5 = tempCtrlName5.replace("$value", tempCtrlValue);
				// controlList =
				// wait.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString)));

				try {// Minaakshi : 05-02-2019
					controlList = dmgebtWait
							.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString5)));

				} catch (Exception ex) {
					log.error(ex.getMessage(), ex);
					log.info("In Second Wait : XPATH : Element is not visible : " + controlName);
					controlList = waitForElementPresence
							.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(tempReplaceString5)));
					// controlList =
					// waitForElementPresence.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(controlName)));
				}
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
					// WebDriverWait wait2 = new WebDriverWait(Automation.driver
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

			case XPath_dy:// Minaakshi : Added this case while handling dyanamic
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
				// controlList =
				// wait.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString1)));
				try {// Minaakshi : 05-02-2019
					controlList = dmgebtWait
							.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString1)));

				} catch (Exception ex) {
					log.error(ex.getMessage(), ex);
					log.info("In Second Wait : XPATH : Element is not visible : " + controlName);
					controlList = waitForElementPresence
							.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(tempReplaceString1)));
					// controlList =
					// waitForElementPresence.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
				}

				break;
			// Added by sheetal
			case XPath_Wait:
				// WebDriverWait wait3 = new WebDriverWait(Automation.driver, 5);//Sel4
				WebDriverWait wait3 = new WebDriverWait(Automation.driver, Duration.ofSeconds(5));
				int p, max = 600;
				for (p = 1; p < max; p += 10) {
					try {
						controlList = wait3.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
						System.out.println("Found element after waiting for approx: " + p + " seconds.");
						break;
					} catch (Exception e) {
						Thread.sleep(10000);
					}
				}
				if (p >= max) {
					System.out.println("Time-out occured after wating for approx." + max + " seconds. Element: "
							+ controlName + " is still not loaded.");
					controlList = null;
				}
				break;
			case Name:
				if (Config.projectName.equals("DistributionManagement")) {
					try {// Minaakshi : 05-02-2019
						controlList = dmgebtWait.until(ExpectedConditions.elementToBeClickable(By.name(controlName)));

					} catch (Exception ex) {
						log.error(ex.getMessage(), ex);
						log.info("In Second Wait : NAME : Element is not visible : " + controlName);
						controlList = waitForElementPresence
								.until(ExpectedConditions.visibilityOfElementLocated(By.name(controlName)));
					}
				} else {
					controlList = wait.until(ExpectedConditions.elementToBeClickable(By.name(controlName)));
				}
				// controlList =
				// wait.until(ExpectedConditions.elementToBeClickable(By.name(controlName)));
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

				System.out.println("css sel is" + controlName);

				break;

			// GAIC AJAX controls - TM:02/02/2015
			case AjaxPath:
				// controlList =
				if (Config.projectName.equals("DistributionManagement")) {
					try {// Minaakshi : 05-02-2019
						controlList = dmgebtWait.until(ExpectedConditions.elementToBeClickable(
								By.xpath(controlName + "[contains(text(),'" + controlValue + "')]")));

					} catch (Exception ex) {
						log.error(ex.getMessage(), ex);
						log.info("In Second Wait : XPATH : Element is not visible : " + controlName);
						controlList = waitForElementPresence
								.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(controlName)));

					}
				} else {

					controlList = wait.until(ExpectedConditions
							.elementToBeClickable(By.xpath(controlName + "[contains(text(),'" + controlValue + "')]")));
					break;
				}

				break;

			case AjaxPath_ExactValue:// Minaakshi : 01-07-2019 : Added this case
				// for selecting Exact value from
				// Drop-Down
				// controlList =
				if (Config.projectName.equals("DistributionManagement")) {
					try {// Minaakshi : 05-02-2019
						controlList = dmgebtWait.until(ExpectedConditions
								.elementToBeClickable(By.xpath(controlName + "[text()='" + controlValue + "']")));

					} catch (Exception ex) {
						log.error(ex.getMessage(), ex);
						log.info("In Second Wait : XPATH : Element is not visible : " + controlName);
						controlList = waitForElementPresence.until(ExpectedConditions
								.visibilityOfElementLocated(By.xpath(controlName + "[text()='" + controlValue + "']")));

					}
				} else {

					controlList = wait.until(ExpectedConditions
							.elementToBeClickable(By.xpath(controlName + "[contains(text(),'" + controlValue + "')]")));
					break;
				}

				break;
			case AjaxPath_dy:// Minaakshi : Added this case while handling
				// dyanamic xpath which requires ctrlvalue from
				// unique number sheet
				String UNSheetValue2 = "";
				if (controlValue.equalsIgnoreCase("Yes")) {
					UNSheetValue2 = DMProduct.ReadFromExcelUsingColumnName("", WebHelper.columnName);
				} else {
					UNSheetValue2 = DMProduct.ReadFromExcelUsingColumnName(controlValue, WebHelper.columnName);
				}

				@SuppressWarnings("unused")
				String temp2 = controlName + "[contains(text(),'" + UNSheetValue2 + "')]";

				try {// Minaakshi : 05-02-2019
					controlList = dmgebtWait.until(ExpectedConditions.elementToBeClickable(
							By.xpath(controlName + "[contains(text(),'" + UNSheetValue2 + "')]")));

				} catch (Exception ex) {
					log.error(ex.getMessage(), ex);
					log.info("In Second Wait : XPATH : Element is not visible : " + controlName);
					controlList = waitForElementPresence.until(ExpectedConditions.visibilityOfElementLocated(
							By.xpath(controlName + "[contains(text(),'" + UNSheetValue2 + "')]")));

				}
				break;

			case Id_p:
			case HTMLID_p:
				controlList = wait.until(ExpectedConditions.presenceOfElementLocated(By.id(controlName)));
				break;

			case XPath_H:

				@SuppressWarnings("unused")
				boolean controlList1 = wait
						.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(controlName)));
				/*
				 * controlList = wait.until(ExpectedConditions
				 * .presenceOfElementLocated(By.xpath(controlName)));
				 */
				break;

			case XPath_p:
				if (Config.projectName.equals("DistributionManagement")) {
					try { // Minaakshi : 05-02-2019
						controlList = dmgebtWait
								.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
					} catch (Exception ex) {
						log.error(ex.getMessage(), ex);
						log.info("In second wait : XPATH : Element is not visible : " + controlName);
						controlList = waitForElementPresence
								.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(controlName)));
					}
				} else {
					controlList = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
				}
				break;

			case Xpath_ctrlvalue:
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

								System.out.println(element.getAttribute("value"));

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

									System.out.println("Element value doesn't match with control value");
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

			}
			return controlList;
		} catch (Exception e) {
			// System.out.println("bhaskar in catch block");
			log.error(e.getMessage(), e);
			throw new Exception(e.getMessage());

		}
	}

	// added code specific to Report Module for LNA Product by Deepti
	public static void pickXLSFile(String ctrlValue) throws IOException, InterruptedException {
		selectXLSFile(ctrlValue);
	}

	private static void DeleteFile() throws IOException {
		MoveFiles();
	}

	public static void connect() throws IOException, InterruptedException {
		ListFiles("Downloads");
	}

	public static void ListFiles(String DirectoryName) throws IOException, InterruptedException {

		String workingDirectory = System.getProperty("user.dir");
		System.out.println(workingDirectory);
		String dir = workingDirectory + File.separator + DirectoryName;
		System.out.println("User Directory : " + dir);

		File file = new File(dir);
		File[] files = file.listFiles();

		if (file.exists()) {
			System.out.println("Directory is found!");
			for (File eachFile : files) {
				System.out.println(eachFile.getName());
				// unzip(dir+File.separator+eachFile.getName(),"D:/eclipse-kepler2/ElixirWorkspace/SeleniumExcel/Downloads");
				unzip(dir + File.separator + eachFile.getName(), dir);
				// Thread.sleep(20000);
				// selectFile(dir); -- moving select file method into unzip method to check
			}
		} else {
			System.out.println("Directory is not exits");
		}

	}

	// private static void unzip(String zipFilePath, String destDir) {
	private static void unzip(String zipFilePath, String destDir) throws InterruptedException {

		if (destDir == null)
			destDir = "";
		else
			destDir += File.separator;
		System.out.println(destDir += File.separator);
		// 1.0 Create output directory
		File outputDirectory = new File(destDir);
		if (outputDirectory.exists())
			outputDirectory.delete();
		outputDirectory.mkdir();
		// 2.0 Unzip (create folders & copy files)
		try {
			// 2.1 Get zip input stream
			ZipInputStream zip = new ZipInputStream(new FileInputStream(zipFilePath));
			ZipEntry entry = null;
			int len;
			byte[] buffer = new byte[1024];
			// 2.2 Go over each entry "file/folder" in zip file
			while ((entry = zip.getNextEntry()) != null) {
				if (!entry.isDirectory()) {
					System.out.println(entry.getName());
					// create a new file
					File file = new File(destDir + entry.getName());
					// create file parent directory if does not exist
					if (!new File(file.getParent()).exists())
						new File(file.getParent()).mkdirs();
					// get new file output stream
					FileOutputStream fos = new FileOutputStream(file);
					// copy bytes
					while ((len = zip.read(buffer)) > 0) {
						fos.write(buffer, 0, len);
					}
					fos.close();
					selectFile(destDir);
				}
				zip.close();
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	private static void selectFile(String path) throws IOException, InterruptedException {

		System.out.println(path);
		File file = new File(path);
		File[] files = file.listFiles();

		if (file.exists()) {
			System.out.println("Directory is found!");
			for (File eachFile : files) {
				String fileExtensionName = eachFile.getName().substring(eachFile.getName().indexOf("."));
				System.out.println("Iam to pick the excel in xlsx format" + eachFile.getName());
				if (fileExtensionName.equals(".xlsx")) {
					String reqFile = eachFile.getName();
					System.out.println(path + File.separator + reqFile);
					System.out.println(reqFile);
					String filename = reqFile.split("\\.")[0];
					// String extension = reqFile.split("\\.")[1];
					System.out.println(filename);
					// Thread.sleep(20000);
					takeToMacro(filename);
					// Thread.sleep(20000);
					// ChangeExcelExtension(reqFile,".xls");
				} else {
					System.out.println("not xlsx-- im in selectFile method");
				}

			}
		}
	}

	public static void takeToMacro(String reqFile) {

		try {

			System.out.println("i am in take to macro method");
			Thread.sleep(5000);
			File fil1 = new File(
					"D://iTAF_Version//iTAF-1.19.1.0-dist//Macro_Externsion_Change//Macro_Extension_Change.xls");

			FileInputStream fin = new FileInputStream(fil1);

			@SuppressWarnings("resource")
			HSSFWorkbook wbook = new HSSFWorkbook(fin);

			HSSFSheet sheet = wbook.getSheet("DataSheet");

			// HSSFCell TCid= MainController.controllerTestCaseID;

			// String tc=TCid.getStringCellValue();
			System.out.println("Take to Macro:" + reqFile);
			sheet.getRow(0).getCell(1).setCellValue(reqFile);

			FileOutputStream fout = new FileOutputStream(fil1);

			wbook.write(fout);
			fout.close();

		} catch (Exception e) {
			System.out.println("in ex");
		}

		finally {
			System.out.println("in finally");

		}

	}

	@SuppressWarnings("unused")
	private static void selectXLSFile(String ctrlValue) throws IOException, InterruptedException {

		System.out.println("im in select xls file method");
		String workingDirectory = System.getProperty("user.dir");
		System.out.println(workingDirectory);
		String dir = workingDirectory + File.separator + "Downloads";
		System.out.println("User Directory : " + dir);
		File file = new File(dir);
		File[] files = file.listFiles();
		for (int i = 0; i < files.length; i++)// length is the property of array
			System.out.println("files after changign the extension to xls:" + files[i]);
		try {
			if (file.exists()) {
				System.out.println("Directory is found!- im in select xls file method");
				for (File eachFile : files) {
					String fileExtensionName = eachFile.getName().substring(eachFile.getName().indexOf("."));
					System.out.println("Iam to pick the excel in xls format" + eachFile.getName());
					if (fileExtensionName.equals(".xls")) {
						String reqFile = eachFile.getName();
						System.out.println(dir + File.separator + reqFile);
						String sourcepath = "path+ File.separator + reqFile";
						System.out.println("Select XLS file :" + reqFile);
						Thread.sleep(5000);
						System.out.println("Im entering in to Compare Two Excel");
						CompareTwoExcel(dir + File.separator + reqFile, ctrlValue);
						// ChangeExcelExtension(reqFile,".xls");
					} else {
						System.out.println("not xls--im in selectxlsfile method");
					}

				}
			}
		} catch (StringIndexOutOfBoundsException e) {
			System.out.println("StringIndexOutOfBoundsException!!");
		}
	}

	// -----------------------------------------

	
	public static void CompareTwoExcel(String reqFile, String ctrlValue) throws IOException, InterruptedException {

		Thread.sleep(5000);
		System.out.println("TestcaseID:" + controller.controllerTestCaseID.toString());
		String TCID = controller.controllerTestCaseID.toString();
		System.out.println("Im in CompareTwoExcel Method");
		HSSFSheet worksheet1;
		FileInputStream fileInputStream1 = new FileInputStream(reqFile);

		@SuppressWarnings("resource")
		HSSFWorkbook workbook1 = new HSSFWorkbook(fileInputStream1);

		if (reqFile.contains("AH1_Samplereport") || reqFile.contains("Appeal_Claim1_Inventory"))

			worksheet1 = workbook1.getSheetAt(1);
		else
			worksheet1 = workbook1.getSheetAt(0);
		int rowCount1 = worksheet1.getPhysicalNumberOfRows();

		HSSFRow row1;
		HSSFRow row2;
		HSSFSheet worksheet2;
		FileInputStream fileInputStream2 = new FileInputStream(ctrlValue); // removed ctrl value

		@SuppressWarnings("resource")
		HSSFWorkbook workbook2 = new HSSFWorkbook(fileInputStream2);
		if ((ctrlValue.equals("D://Downloads//Claims - Dental DHMO Report Examples.xls")
				&& TCID.equals("TSC_Buyup_Port_01_15")) || TCID.equals("TSC_Buyup_Port_01_63")) {
			worksheet2 = workbook2.getSheetAt(4);
		}
		if (ctrlValue.equals("D://Downloads//AH1_Samplereport_4221_6246_6246.xls")
				|| ctrlValue.equals("D://Downloads//Appeal_Claim1_Inventory_4225_6250_6250.xls")) {
			worksheet2 = workbook2.getSheetAt(1);
		} else if (ctrlValue.equals("D://Downloads//Claims - Dental DHMO Report Examples.xls")
				&& TCID.equals("TSC_Buyup_Port_01_56")) {
			worksheet2 = workbook2.getSheetAt(6);
		} else if ((ctrlValue.equals("D://Downloads//Claims - Dental DHMO Report Examples.xls")
				&& TCID.equals("TSC_Buyup_Port_01_57")) || TCID.equals("TSC_Buyup_Port_01_10")) {
			worksheet2 = workbook2.getSheetAt(0);
		} else if ((ctrlValue.equals("D://Downloads//Claims - Dental DHMO Report Examples.xls")
				&& TCID.equals("TSC_Buyup_Port_01_58")) || TCID.equals("TSC_Buyup_Port_01_1")) {
			worksheet2 = workbook2.getSheetAt(1);
		} else if (ctrlValue.equals("D://Downloads//Claims - Dental DHMO Report Examples.xls")
				&& TCID.equals("TSC_Buyup_Port_01_59")) {
			worksheet2 = workbook2.getSheetAt(2);
		} else if ((ctrlValue.equals("D://Downloads//Claims - Dental DHMO Report Examples.xls")
				&& TCID.equals("TSC_Buyup_Port_01_60")) || TCID.equals("TSC_Buyup_Port_01_9")) {
			worksheet2 = workbook2.getSheetAt(3);
		} else if ((ctrlValue.equals("D://Downloads//Claims - Dental DHMO Report Examples.xls")
				&& TCID.equals("TSC_Buyup_Port_01_61")) || TCID.equals("TSC_Buyup_Port_01_64")) {
			worksheet2 = workbook2.getSheetAt(5);
		} else if (ctrlValue
				.equals("D://Downloads//LTD DISWEE5110_ SBC Bridging (Report View) 2019-02-18T08_23_59.856Z.xls")
				&& TCID.equals("TSC_Buyup_Port_01_21")) {
			worksheet2 = workbook2.getSheetAt(1);
		} else if (ctrlValue
				.equals("D://Downloads//LTD DISWEE5110_ SBC Bridging (Report View) 2019-02-18T08_23_59.856Z.xls")
				&& TCID.equals("TSC_Buyup_Port_01_22")) {
			worksheet2 = workbook2.getSheetAt(2);
		} else if (ctrlValue
				.equals("D://Downloads//LTD DISWEE5110_ SBC Bridging (Report View) 2019-02-18T08_23_59.856Z.xls")
				&& TCID.equals("TSC_Buyup_Port_01_23")) {
			worksheet2 = workbook2.getSheetAt(3);
		} else {
			worksheet2 = workbook2.getSheetAt(0);
		}
		int rowCount2 = worksheet2.getPhysicalNumberOfRows();
		System.out.println("Row Count 1 : " + rowCount1 + " Row Count 2 : " + rowCount2);
		int rowCount = rowCount1;
		if (rowCount2 > rowCount1)
			rowCount = rowCount2;
		if (rowCount1 == rowCount2 || rowCount1 != rowCount2) {
			Thread.sleep(5000);
			for (int i = 0; i < rowCount; i++) {
				if (ctrlValue.equals("D://Downloads//NH_Claims_Analysis_April_2019.xls")) {
					row1 = worksheet1.getRow(i);

					row2 = worksheet2.getRow(1);
				} else if (ctrlValue.equals(
						"D://Downloads//Claims - LTD Pending and Approved Claims Open Callup-LTD DenialTerminated Weekly Report.xls")) {
					row1 = worksheet1.getRow(i);

					row2 = worksheet2.getRow(2);

				} else if (ctrlValue.equals("D://Downloads//NH_DEnied Claims Field Descriptors.xls")
						&& TCID.equals("TSC_Buyup_Port_01_30")) {
					row1 = worksheet1.getRow(i);

					row2 = worksheet2.getRow(5);
				} else {
					row1 = worksheet1.getRow(i);
					row2 = worksheet2.getRow(i);
				}
				// ------------------------------ comparing Header --------------------------

				@SuppressWarnings("rawtypes")
				Iterator rowIterator;
				if (rowCount1 > rowCount2)
					rowIterator = worksheet1.rowIterator();
				else
					rowIterator = worksheet2.rowIterator();
				/**
				 * Escape the header row *
				 */

				if (rowIterator.hasNext()) {
					Row headerRow = (Row) rowIterator.next();
					// get the number of cells in the header row
					int numberOfCells = headerRow.getPhysicalNumberOfCells();
					System.out.println(numberOfCells);

					for (int j = 0; j < numberOfCells; j++) {
						HSSFCell id2;
						String idstr1 = "";
						if (i < rowCount1) {
							HSSFCell id1 = row1.getCell(j);
							if (id1 != null) {
								id1.setCellType(CellType.STRING);
								idstr1 = id1.getStringCellValue();
							}
						}

						// for null values to compare correct
						// else
						// idstr1 = "XXXXXX";

						String idstr2 = "";
						if (i < rowCount2) {
							if (ctrlValue.equals("D://Downloads//NH_Claims_Analysis_April_2019.xls")) {
								id2 = row2.getCell(j + 1);
							} else {
								id2 = row2.getCell(j);
							}
							if (id2 != null) {
								id2.setCellType(CellType.STRING);
								idstr2 = id2.getStringCellValue();
							}
						}
						// for null values to compare correct
						// else
						// idstr2 = "XXXXXX";

						@SuppressWarnings("unused")
						String logicalname = headerRow.getCell(j).toString();

						if (!idstr1.trim().equals(idstr2.trim())) {
							System.out.println("[ERROR] :" + "Diference for id (book1) " + idstr1 + "| Book 1 id = "
									+ idstr1 + " | Book 2 id = " + idstr2);
							// test.log(LogStatus.ERROR,"Diference for id (book1) " + idstr1 + "| Book 1 id
							// = " + idstr1+ "| Book 2 id = " + idstr2);
							// commented below method to print correct null comparison-- need to work in
							// detail
							// WebHelperLNA.WriteToDetailResults_distinctexcel(idstr1, idstr2, logicalname);
							WebHelperLNA.WriteToDetailResults(idstr1, idstr2, idstr1);
						} else {
							System.out.println("[Processing] :" + "ID " + idstr1 + "=> Book 1 id = " + idstr1
									+ " Book 2 id = " + idstr2);
							// commented below method to print correct null comparison-- need to work in
							// detail
							// WebHelperLNA.WriteToDetailResults_distinctexcel(idstr1, idstr2, logicalname);
							WebHelperLNA.WriteToDetailResults(idstr1, idstr2, idstr1);
						}
					}
				}
			}
		}
	}

	// -----added code specific to Report Module for LNA Product by Deepti----
	public static void MoveFiles() throws IOException {
		Date date = new Date();
		SimpleDateFormat formatter = new SimpleDateFormat("MM_dd_yyyy_HH_mm_ss");
		String strDate = formatter.format(date);

		String path = "D://iTAF_Version//iTAF-1.19.1.0-dist//Downloads";
		File file1 = new File(path);
		File[] files = file1.listFiles();
		if (file1.exists()) {
			System.out.println("Directory is found!");
			for (File eachFile : files) {

				String reqFile = eachFile.getName();
				String reqFilearray = reqFile.split("\\.")[0];
				String reqFilearray1 = reqFile.split("\\.")[1];
				reqFile = reqFilearray + '_' + strDate + '.' + reqFilearray1;
				File destDir = new File("D://iTAF_Version//iTAF-1.19.1.0-dist//Reporting-ComparedFiles//" + reqFile);
				try {
					FileUtils.moveFile(eachFile, destDir);
					System.out.println("Files moved succesfully");
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	/*
	 * public static String getToken() throws IOException {
	 * 
	 * String token = ""; try { String urlParameters =
	 * "username=Aaron.Austin&password=Austin"; byte[] postData =
	 * urlParameters.getBytes();// getBytes( // StandardCharsets.F_8 // ); int
	 * postDataLength = postData.length;
	 * 
	 * String request = Config.Token_URL; URL url = new URL(request);
	 * HttpURLConnection conn = (HttpURLConnection) url.openConnection();
	 * conn.setDoOutput(true); conn.setInstanceFollowRedirects(false);
	 * conn.setRequestMethod("POST"); conn.setRequestProperty("Content-Type",
	 * "application/x-www-form-urlencoded"); conn.setRequestProperty("charset",
	 * "utf-8"); conn.setRequestProperty("Content-Length",
	 * Integer.toString(postDataLength)); conn.setRequestProperty("X-PAS-CLIENT",
	 * "SWAGGER"); conn.setUseCaches(false);
	 * 
	 * OutputStream os = conn.getOutputStream(); os.write(postData); os.flush();
	 * os.close();
	 * 
	 * int responseCode = conn.getResponseCode();
	 * System.out.println("POST Response Code :: " + responseCode);
	 * 
	 * if (responseCode == HttpURLConnection.HTTP_OK) { // success BufferedReader in
	 * = new BufferedReader(new InputStreamReader(conn.getInputStream())); String
	 * inputLine; StringBuffer response = new StringBuffer();
	 * 
	 * while ((inputLine = in.readLine()) != null) { response.append(inputLine); }
	 * in.close();
	 * 
	 * System.out.println(response.toString()); JSONParser parser = new
	 * JSONParser(); JSONObject json = (JSONObject)
	 * parser.parse(response.toString()); String Part1; String Part2; Part1 =
	 * (String) (json.get("token_type")); Part2 = (String)
	 * (json.get("access_token"));
	 * 
	 * token = Part1.concat(" " + Part2);
	 * 
	 * } else { System.out.println("POST request not worked"); BufferedReader in =
	 * new BufferedReader(new InputStreamReader(conn.getErrorStream())); String
	 * inputLine; StringBuffer response = new StringBuffer();
	 * 
	 * while ((inputLine = in.readLine()) != null) { response.append(inputLine); }
	 * in.close();
	 * 
	 * // print result System.out.println(response.toString()); } } catch (Exception
	 * e) { e.printStackTrace(); }
	 * 
	 * return token;
	 * 
	 * }
	 */
	// Sel4

	/*
	 * public static String getToken() throws IOException { String token = ""; try {
	 * String urlParameters = "username=Aaron.Austin&password=Austin"; byte[]
	 * postData = urlParameters.getBytes(); int postDataLength = postData.length;
	 * 
	 * String request = Config.Token_URL; URL url = new URL(request);
	 * HttpURLConnection conn = (HttpURLConnection) url.openConnection();
	 * conn.setDoOutput(true); conn.setInstanceFollowRedirects(false);
	 * conn.setRequestMethod("POST"); conn.setRequestProperty("Content-Type",
	 * "application/x-www-form-urlencoded"); conn.setRequestProperty("charset",
	 * "utf-8"); conn.setRequestProperty("Content-Length",
	 * Integer.toString(postDataLength)); conn.setRequestProperty("X-PAS-CLIENT",
	 * "SWAGGER"); conn.setUseCaches(false);
	 * 
	 * OutputStream os = conn.getOutputStream(); os.write(postData); os.flush();
	 * os.close();
	 * 
	 * int responseCode = conn.getResponseCode();
	 * System.out.println("POST Response Code :: " + responseCode);
	 * 
	 * if (responseCode == HttpURLConnection.HTTP_OK) { BufferedReader in = new
	 * BufferedReader(new InputStreamReader(conn.getInputStream())); JsonReader
	 * reader = Json.createReader(in); JsonObject json = reader.readObject();
	 * reader.close();
	 * 
	 * String Part1 = json.getString("token_type"); String Part2 =
	 * json.getString("access_token");
	 * 
	 * token = Part1.concat(" " + Part2); } else {
	 * System.out.println("POST request not worked"); BufferedReader in = new
	 * BufferedReader(new InputStreamReader(conn.getErrorStream())); String
	 * inputLine; StringBuffer response = new StringBuffer();
	 * 
	 * while ((inputLine = in.readLine()) != null) { response.append(inputLine); }
	 * in.close(); System.out.println(response.toString()); } } catch (Exception e)
	 * { e.printStackTrace(); }
	 * 
	 * return token; }
	 */
	
	//Sel4
	public static String getToken() throws IOException {
	    String token = "";
	    try {
	        String urlParameters = "username=Aaron.Austin&password=Austin";
	        byte[] postData = urlParameters.getBytes("UTF-8");

	        String request = Config.Token_URL;
	        URL url = new URL(request);
	        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
	        conn.setDoOutput(true);
	        conn.setRequestMethod("POST");
	        conn.setRequestProperty("Content-Type", "application/x-www-form-urlencoded");

	        try (OutputStream os = conn.getOutputStream()) {
	            os.write(postData);
	        }

	        int responseCode = conn.getResponseCode();
	        System.out.println("POST Response Code :: " + responseCode);

	        if (responseCode == HttpURLConnection.HTTP_OK) {
	            try (BufferedReader in = new BufferedReader(new InputStreamReader(conn.getInputStream()))) {
	                ObjectMapper objectMapper = new ObjectMapper();
	                JsonNode jsonNode = objectMapper.readTree(in);
	                String part1 = jsonNode.get("token_type").asText();
	                String part2 = jsonNode.get("access_token").asText();
	                token = part1 + " " + part2;
	            }
	        } else {
	            System.out.println("POST request not worked");
	            try (BufferedReader in = new BufferedReader(new InputStreamReader(conn.getErrorStream()))) {
	                String inputLine;
	                StringBuffer response = new StringBuffer();

	                while ((inputLine = in.readLine()) != null) {
	                    response.append(inputLine);
	                }
	                System.out.println(response.toString());
	            }
	        }
	    } catch (IOException e) {
	        e.printStackTrace();
	    }
	    return token;
	}
	
	

	// adding call to webserviceJSON - deepthi
	public static void callWebServiceJSON(String wsdlUr, String requestXML, String user, String password)
			throws Exception {
		// OutputStreamWriter requestWriter = null;//Sel4
		// BufferedReader responseReader = null;//Sel4
		Scanner scanner = null;

		try {

			URL wsdlUrl = new URL(WebHelper.wsdl_url);
			HttpURLConnection con = (HttpURLConnection) wsdlUrl.openConnection();
			// String
			// basicAuth="eyJhbGciOiJIUzI1NiJ9.eyJleHAiOjE1NjkzMjQzOTgsInN1YiI6IlRva2VuIiwiaXNzIjoiUEFTIiwiYXVkIjoiQ2xpZW50IiwidXNlcm5hbWUiOiJ0ZXN0dXNlciJ9.VhSVu5dLQezXG9kLPDIrSahy5XyW2Y26tmAClxvJZrg";

			String basicAuth = getToken();
			con.setRequestProperty("Authorization", basicAuth);
			/** Proxy settings ONLY if required **/
			/*
			 * System.setProperty("http.proxyHost", "pneprxy.majesco.com");
			 * System.setProperty("http.proxyPort", "8080"); con.usingProxy();
			 */

			con.setRequestMethod("POST");
			con.setRequestProperty("Content-Language", "en-US");
			con.setRequestProperty("Accept", "application/json");
			con.setRequestProperty("Content-Type", "application/json");
			con.setRequestProperty("X-PAS-CLIENT", "PORTAL");
			con.setRequestProperty("X-PAS-OPERATOR", "Aaron.Austin");
			con.setDoOutput(true);
			con.setDoInput(true);

			/** Reading data from Request XML **/
			String reqXml = System.getProperty("user.dir") + "\\LNAProduct\\Resources\\Input\\WebService\\"
					+ WebHelper.request_xml + ".json";
			scanner = new Scanner(new File(reqXml));
			String soapMessage = scanner.useDelimiter("\\A").next();

			OutputStream os = con.getOutputStream();
			os.write(soapMessage.getBytes());
			os.flush();

			BufferedReader br = new BufferedReader(new InputStreamReader((con.getInputStream())));

			String output;
			System.out.println("Output from Server .... \n");
			while ((output = br.readLine()) != null) {
				System.out.println(output);
			}

			con.disconnect();

		} catch (Exception e) {
			log.error(e.getMessage(), e);
			throw new Exception("Error while triggering web service: " + e.getMessage());
		} finally {
			/** Closing input and output stream buffers **/

			scanner.close();
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
		// System.out.println(sikscreen);
		// bhaskar
		if (controlType.contains("Robot") && !WebHelper.isIntialized) {
			System.out.println("In method doaction debug1");
			WebHelper.robot = new Robot();
			WebHelper.isIntialized = true;
		}

		log.info("ctrlValue doAction:" + ctrlValue);
		// Mrinmayee - skip the step for which control value is blank : start

		JavascriptExecutor js = (JavascriptExecutor) Automation.driver;
		try {
			// getHTMLResponse();
			// System.out.println("In method doaction debug3");
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
							Automation.driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(50));
							// Thread.sleep(2000);

							Thread.sleep(2000);
						} else {
							System.out.println("Element Not Found");
						}
						;

						// Search_Box.sendKeys(Keys.ENTER);

					} catch (StaleElementReferenceException e) {
						log.error(e.getMessage(), e);
						System.out.println("Trying to recover from a stale element :" + e.getMessage());

					}

					break;

				case Read:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					if (Config.projectName.equals("DistributionManagement")) {
						uniqueNumber = DMProduct.ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
					} else {
						uniqueNumber = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
					}
					// System.out.println("!!!!!!!!!!!!!!!!");
					// System.out.println("uniqueNumber:"+uniqueNumber);
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
				case Input: // Replace fields InputifExist and I by Asif
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}

					try {
						if (webElement.isEnabled()) {

							if (!ctrlValue.equalsIgnoreCase("null")) {
								{

									if (Config.projectName.equals("LNAProduct")) {
										// webElement.clear();
										webElement.click();
										webElement.sendKeys(ctrlValue);
										break;
									} else

										// System.out.println("!@#$%^&*!@#$%^&*ctrlValue:"+ctrlValue);
										for (int i = 0; i < 5; i++) {
											webElement.sendKeys(Keys.ENTER);

											webElement.clear();

											Thread.sleep(500);
											webElement.sendKeys(ctrlValue);
											// webElement.click();
											System.out.println(webElement.getAttribute("value"));
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
						System.out.println("Element not exist");
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
						System.out.println(currentValue);
						if (currentValue.equalsIgnoreCase(ctrlValue)) {

							System.out.println("PASSED");
						} else {
							System.out.println("FAILED");
						}
					} catch (Exception e) {
						log.error(e.getMessage(), e);
						currentValue = "null";
						if (currentValue.equalsIgnoreCase(ctrlValue)) {

							System.out.println("PASSED");
						} else {
							System.out.println("FAILED");
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
						System.out.println("Element does not exist");

					}
					break;
				// js1.executeScript("arguments[0].value = '';",
				// webElement);

				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					// Thread.sleep(500);
					if (ctrlValue.contains("|")) {// Minaakshi : 01-05-2019
						String[] temp = ctrlValue.split("\\|");
						ctrlValue = temp[1].toString();
					}
					if (Config.projectName.equals("DistributionManagement")) {// Minaakshi
						if (!ctrlValue.equalsIgnoreCase("null")) {
							System.out.println("!@#$%^&*!@#$%^&*ctrlValue:" + ctrlValue);
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
									WebHelperUtil.InputValue(webElement, ctrlValue, controlId, controlName, imageType,
											ctrlValue);
								}

							} catch (Exception e) {
								log.error(e.getMessage(), e);
								System.out.println("Element not displayed");
							}
						}
					}
					break;// Minaakshi : 17-12-2018

				}
				break;

			case BatchStatus:
				switch (actionName) {

				case I:

					System.out.println("In BatchStatus");
					Long starttime = System.currentTimeMillis();
					String CurStus;
					if (controlName.equalsIgnoreCase("")) {
						break;
					}

					if (controlName.contains("+")) {
						controlName = controlName.replace("+", ctrlValue);
					}
					boolean flag = false;
					do {
						try {

							Long endtime = System.currentTimeMillis();
							if (((endtime - starttime) / 1000) > 260) {
								System.out.println("Task has not been created within 260 seconds");
								break;

							}

							CurStus = Automation.driver.findElement(By.xpath(controlName)).getText();

							// added for clicking correct group
							/*
							 * webElement = Automation.driver.findElement(By.xpath(controlName));
							 * 
							 * webElement.click();
							 */

							System.out.println("Task Created and its status or name is : " + CurStus);

							long ttime = endtime - starttime;

							int time = (int) (ttime / 1000);

							double time1 = time / 60;
							System.out.println("Activity created in " + time1 + "Minutes");

							flag = true;
						} catch (WebDriverException e) {
							(Automation.driver.findElement(By.xpath("(//*[@value='Refresh'])[1]"))).click();
							System.out.println("Refresh clicked");
							Thread.sleep(5000);

						}
					} while (flag == false);
				}
				break;

			// Added by Suhas (L&A) 08/13/2019 for Claim task finder
			case TaskStatus:
				switch (actionName) {

				case I:

					System.out.println("In TaskStatus");
					Long starttime = System.currentTimeMillis();
					String CurStus;
					if (ctrlValue.equalsIgnoreCase("")) {
						break;
					}

					if (controlName.contains("+")) {
						controlName = controlName.replace("+", ctrlValue);
					}
					boolean flag = false;
					do {
						try {

							Long endtime = System.currentTimeMillis();
							if (((endtime - starttime) / 1000) > 260) {
								System.out.println("Task has not been created within 260 seconds");
								break;

							}

							CurStus = Automation.driver.findElement(By.xpath(controlName)).getText();

							// added for clicking correct group
							/*
							 * webElement = Automation.driver.findElement(By.xpath(controlName));
							 * 
							 * webElement.click();
							 */

							System.out.println("Task Created and its status or name is : " + CurStus);

							long ttime = endtime - starttime;

							int time = (int) (ttime / 1000);

							double time1 = time / 60;
							System.out.println("Activity created in " + time1 + "Minutes");

							flag = true;
						} catch (WebDriverException e) {
							(Automation.driver.findElement(By.xpath("//*[contains(@class,'icon-refresh')]"))).click();
							System.out.println("Refresh clicked");
							Thread.sleep(5000);

						}
					} while (flag == false);
				}
				break;

			case HolidayBatch:

				break;

			// Added by Suhas (L&A) 08/13/2019 for search and collapse
			case Dental_ClaimStatus:
				System.out.println("in Dental_ClaimStatus");
				switch (actionName) {

				case NC:
					String CurStus;
					Long starttime = System.currentTimeMillis();

					if (controlName.contains("+")) {

						controlName = controlName.replace("+", ctrlValue);
					}

					boolean flag = false;
					do {

						try {
							Thread.sleep(1000);

							Long endtime = System.currentTimeMillis();
							if (((endtime - starttime) / 1000) > 960) {
								System.out.println("Task has not been created within 960 seconds");
								break;

							}
							CurStus = Automation.driver.findElement(By.xpath(controlName)).getText();
							System.out.println("Task Created and its status or name is : " + CurStus);

							flag = true;

						} catch (WebDriverException e) {

							/*
							 * new WebDriverWait(Automation.driver, 60).until(ExpectedConditions
							 * .presenceOfElementLocated(By.xpath("//*[contains(@class,'chevron-down')]")));
							 */// Sel4
							WebDriverWait wait = new WebDriverWait(Automation.driver, Duration.ofSeconds(60));
							wait.until(ExpectedConditions
									.presenceOfElementLocated(By.xpath("//*[contains(@class,'chevron-down')]")));
							/*
							 * new WebDriverWait(Automation.driver, 60).until(ExpectedConditions
							 * .presenceOfElementLocated(By.xpath("(//span[contains(text(),'Search')])[2]"))
							 * );
							 */// Sel4
							WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(60));
							wait1.until(ExpectedConditions
									.presenceOfElementLocated(By.xpath("(//span[contains(text(),'Search')])[2]")));

							(Automation.driver.findElement(By.xpath("//*[contains(@class,'chevron-down')]"))).click();
							(Automation.driver.findElement(By.xpath("//button//span[text()='Search']"))).click();
							System.out.println("Colapsar Clicked");

							Thread.sleep(3000);
						}

					} while (flag == false);
				}

				break;

			case Opt_CurrentDate:

				switch (actionName) {

				case I:

					int counter = 1;
					boolean flag = false;
					do {
						try {

							if (ctrlValue.equalsIgnoreCase("Yes") || ctrlValue.length() > 0) {
								// (//td[contains(@class, 'weekend active start-date active end-date available')
								// or contains(@class, 'today active start-date active end-date available')])
								WebElement newele = Automation.driver.findElement(By.xpath(
										"(//td[contains(@class, 'weekend active start-date active end-date available') or contains(@class, 'today active start-date')])"
												+ "[" + counter + "]"));

								boolean status = newele.isDisplayed();
								System.out.println("Element status is : " + status);

								String js2 = "arguments[0].style.height='auto'; arguments[0].style.visibility='visible';";

								// Execute the Java Script for the element which we find out
								((JavascriptExecutor) Automation.driver).executeScript(js2, newele);
								newele.click();

								System.out.println(newele);

								flag = true;
							}

							break;
						} catch (Exception e) {
							System.out.println("xpath not matched in " + counter + "  attempt");

							counter++;
							if (counter >= 1000) {
								/*
								 * System.out.println("IN enter"); robot.keyPress(KeyEvent.VK_ENTER);
								 * robot.keyRelease(KeyEvent.VK_ENTER);
								 * 
								 * robot.keyPress(KeyEvent.VK_ENTER); robot.keyRelease(KeyEvent.VK_ENTER);
								 * System.out.println("out enter");
								 */
								// throw new WebDriverException();

							}

						}

					} while (flag == false);

				}

				break;

			case TaskStatus_LTD:
				switch (actionName) {

				case I:

					System.out.println("In TaskStatus");

					if (ctrlValue.equalsIgnoreCase("")) {
						break;
					}

					Long starttime = System.currentTimeMillis();
					String CurStus;
					if (controlName.contains("+")) {
						controlName = controlName.replace("+", ctrlValue);
					}
					boolean flag = false;
					do {
						try {
							Long endtime = System.currentTimeMillis();
							if (((endtime - starttime) / 1000) > 300) {
								System.out.println("Task has not been created within 300 seconds");
								break;
							}
							CurStus = Automation.driver.findElement(By.xpath(controlName)).getText();

							System.out.println("Task Created and its status or name is : " + CurStus);

							long ttime = endtime - starttime;

							int time = (int) (ttime / 1000);

							double time1 = time / 60;
							System.out.println("Activity created in " + time1 + "Minutes");
							flag = true;
						} catch (WebDriverException e) {

							(Automation.driver
									.findElement(By.xpath("(//button//span[contains(@class,'fa fa-refresh')])[2]")))
									.click();
							System.out.println("Refresh clicked");
							Thread.sleep(3000);
							Automation.driver.findElement(By.xpath("//label[text()='My Cases']")).click();
							Thread.sleep(3000);
							Automation.driver
									.findElement(By.xpath(
											"//label[text()='My Team Cases']/following::label[text()='Offered']"))
									.click();
							Thread.sleep(3000);
						}
					} while (flag == false);
				}
				break;

			case TaskStatus_WOP:
				switch (actionName) {

				case I:

					System.out.println("In TaskStatus");

					if (ctrlValue.equalsIgnoreCase("")) {
						break;
					}

					Long starttime = System.currentTimeMillis();
					String CurStus;
					if (controlName.contains("+")) {
						controlName = controlName.replace("+", ctrlValue);
					}
					boolean flag = false;
					do {
						try {

							Long endtime = System.currentTimeMillis();
							if (((endtime - starttime) / 1000) > 300) {
								System.out.println("Task has not been created within 300 seconds");
								break;
							}
							CurStus = Automation.driver.findElement(By.xpath(controlName)).getText();
							System.out.println("Task Created and its status or name is : " + CurStus);
							long ttime = endtime - starttime;
							int time = (int) (ttime / 1000);
							double time1 = time / 60;
							System.out.println("Activity created in " + time1 + "Minutes");
							flag = true;
						} catch (WebDriverException e) {
							(Automation.driver
									.findElement(By.xpath("(//button//span[contains(@class,'fa fa-refresh')])[2]")))
									.click();
							System.out.println("Refresh clicked");
							Thread.sleep(3000);
							Automation.driver
									.findElement(By.xpath(
											"//label[text()='My Team Cases']/following::label[text()='Offered']"))
									.click();
							Thread.sleep(3000);
							Automation.driver.findElement(By.xpath("//label[text()='My Cases']")).click();
							Thread.sleep(3000);
						}
					} while (flag == false);
				}
				break;

			// Added by Suhas (L&A) 08/13/2019 for current date selection
			case CurrentDate:

				switch (actionName) {

				case I:

					int counter = 1;
					boolean flag = false;
					do {
						try {

							// (//td[contains(@class, 'weekend active start-date active end-date available')
							// or contains(@class, 'today active start-date active end-date available')])
							WebElement newele = Automation.driver.findElement(By.xpath(
									"(//td[contains(@class, 'weekend active start-date active end-date available') or contains(@class, 'today active start-date')])"
											+ "[" + counter + "]"));

							boolean status = newele.isDisplayed();
							System.out.println("Element status is : " + status);

							String js2 = "arguments[0].style.height='auto'; arguments[0].style.visibility='visible';";

							// Execute the Java Script for the element which we find out
							((JavascriptExecutor) Automation.driver).executeScript(js2, newele);
							newele.click();

							System.out.println(newele);

							flag = true;
							break;
						}

						catch (Exception e) {
							System.out.println("xpath not matched in " + counter + "  attempt");

							counter++;
							if (counter >= 500) {
								/*
								 * System.out.println("IN enter"); robot.keyPress(KeyEvent.VK_ENTER);
								 * robot.keyRelease(KeyEvent.VK_ENTER);
								 * 
								 * robot.keyPress(KeyEvent.VK_ENTER); robot.keyRelease(KeyEvent.VK_ENTER);
								 * System.out.println("out enter");
								 */
								// throw new WebDriverException();

							}

						}

					} while (flag == false);

				}

				break;

			case Opt_Back_Date:

				switch (actionName) {

				case I:

					System.out.println("In CalendarAll_Back ");
					int counter = 0;
					boolean flag = false;
					do {
						try {

							if (ctrlValue.equalsIgnoreCase("Yes") || ctrlValue.length() > 0) {
								// WebElement
								// newele=Automation.driver.findElement(By.xpath("(//input[@name='_start'])" +
								// "["+counter+"]" ));
								WebElement newele = Automation.driver
										.findElement(By.xpath(controlName + "[" + counter + "]"));

								boolean status = newele.isDisplayed();
								System.out.println("Element status is : " + status);

								String jsBD = "arguments[0].style.height='auto'; arguments[0].style.visibility='visible';";

								// Execute the Java Script for the element which we find out
								((JavascriptExecutor) Automation.driver).executeScript(jsBD, newele);

								System.out.println(newele);
								newele.click();
								newele.clear();
								newele.sendKeys(ctrlValue);
								Thread.sleep(1500);
								newele.sendKeys(Keys.ENTER);
								System.out.println(newele);

								// newele.click();

								flag = true;
							}
							break;
						} catch (Exception e) {
							System.out.println("xpath not matched in " + counter + "  attempt");
							counter++;

							if (counter >= 500) {
								throw new WebDriverException();

							}

						}

					} while (flag == false);

				}

				break;

			// Added by Suhas (L&A) 08/13/2019 for back/future date selection 08/20/2019
			case Back_Date:

				switch (actionName) {

				case I:

					System.out.println("In CalendarAll_Back ");
					int counter = 0;
					boolean flag = false;
					do {
						try {

							// WebElement
							// newele=Automation.driver.findElement(By.xpath("(//input[@name='_start'])" +
							// "["+counter+"]" ));
							WebElement newele = Automation.driver
									.findElement(By.xpath(controlName + "[" + counter + "]"));

							boolean status = newele.isDisplayed();
							System.out.println("Element status is : " + status);

							String jsBD = "arguments[0].style.height='auto'; arguments[0].style.visibility='visible';";

							// Execute the Java Script for the element which we find out
							((JavascriptExecutor) Automation.driver).executeScript(jsBD, newele);

							System.out.println(newele);
							newele.click();
							newele.clear();
							newele.sendKeys(ctrlValue);
							Thread.sleep(1500);
							newele.sendKeys(Keys.ENTER);
							System.out.println(newele);

							// newele.click();

							flag = true;
							break;
						} catch (Exception e) {
							System.out.println("xpath not matched in " + counter + "  attempt");
							counter++;

							if (counter >= 500) {
								throw new WebDriverException();

							}

						}

					} while (flag == false);

				}

				break;

			// Added cse WebEdit_Cal by Deepthi(L&A) 08/13/2019 for optional WebEdit_Cal
			case WebEdit_Cal:
				switch (actionName) {
				case I:
					webElement.click();
					Thread.sleep(1000);
					webElement.clear();
					// Thread.sleep(1000);
					webElement.sendKeys(ctrlValue);
					webElement.sendKeys(Keys.ENTER);
					break;
				}
				break;
			// Added cse WebEdit_C by Deepthi(L&A) 08/13/2019 for optional WebEdit_C
			case WebEdit_C:
				switch (actionName) {
				case I:
					webElement.click();
					Date date = new Date();
					SimpleDateFormat formatter = new SimpleDateFormat("MM/dd/yyyy");
					String strDate = formatter.format(date);
					ctrlValue = strDate;
					System.out.println("Dte passed is:" + ctrlValue);
					webElement.sendKeys(ctrlValue);

					break;
				}
				break;

			// Added cse Opt_WebElment by Suhas (L&A) 08/13/2019 for optional webelement

			case Opt_WebElment:
				switch (actionName) {
				case I:

					try {

						if (controlName.contains("+"))
							controlName = controlName.replace("+", ctrlValue);

						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| ctrlValue.length() > 0) {

							WebElement optElement = Automation.driver.findElement(By.xpath(controlName));

							optElement.click();
							Thread.sleep(700);

						}
					} catch (Exception e) {
						break;
					}

					break;
				case NC:
					if (controlName.contains("+"))
						controlName = controlName.replace("+", ctrlValue);

					if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
							|| ctrlValue.length() > 0) {

						WebElement optElement = Automation.driver.findElement(By.xpath(controlName));

						optElement.clear();
						optElement.sendKeys(ctrlValue);

						break;

					}
				}
				break;

			case logoutLogic:

				switch (actionName) {
				case NC:

					@SuppressWarnings("unused")
					Long starttime = System.currentTimeMillis();

					boolean flag = false;
					do {
						try {

							WebElement logout = Automation.driver.findElement(By.xpath("//a[text()='Log Out']"));
							logout.click();
							System.out.println("Firsttime");
							Thread.sleep(20000);
							if (logout.isDisplayed())

								logout.click();
							System.out.println("Clicked 2nd time");
							flag = true;
							// long ttime = endtime - starttime;

							// int time = (int) (ttime/1000);

							// double time1 = time/60;
						} catch (WebDriverException e) {

							System.out.println("in exception");
							WebElement logout2 = Automation.driver
									.findElement(By.xpath(".//input[@name='j_username']"));
							if (logout2.isDisplayed())

								flag = true;

						}

					} while (flag == false);

				}
				break;

			// Added case spWait by Suhas (L&A) 08/13/2019 for waiting last element is
			// loaded on page then click next element

			case spWait:
				if (controlName.contains("+")) {

					controlName = controlName.replace("+", ctrlValue);
				}

				switch (actionName) {
				/*
				 * case NC:
				 * 
				 * new WebDriverWait(Automation.driver, 300)
				 * .until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
				 * new WebDriverWait(Automation.driver, 300)
				 * .until(ExpectedConditions.elementToBeClickable(By.xpath(controlName))); new
				 * WebDriverWait(Automation.driver, 300)
				 * .until(ExpectedConditions.visibilityOfElementLocated(By.xpath(controlName)));
				 * // System.out.println("Out of 60 seconds");
				 * 
				 * break;
				 */
				// Sel4

				case NC:
					WebDriverWait wait = new WebDriverWait(Automation.driver, Duration.ofSeconds(300));
					wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
					wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(controlName)));
					break;

				/*
				 * case I:
				 * 
				 * if (ctrlValue.equalsIgnoreCase("Yes") || ctrlValue.equalsIgnoreCase("Y")) {
				 * 
				 * new WebDriverWait(Automation.driver, 300)
				 * .until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
				 * new WebDriverWait(Automation.driver, 300)
				 * .until(ExpectedConditions.elementToBeClickable(By.xpath(controlName))); //
				 * System.out.println("Out of 60 seconds"); break;
				 * 
				 * } break;
				 * 
				 * } break;
				 */ // Sel4
				// added for Report Module specific for LNA Product by deepti

				case I:
					if (ctrlValue.equalsIgnoreCase("Yes") || ctrlValue.equalsIgnoreCase("Y")) {

						WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(300));
						wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
						wait1.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
						break;
					}
					break;
				}
				break;

			case connect:

				connect();
				break;

			case pickXLS:

				pickXLSFile(ctrlValue);
				break;

			case DeleteFiles:

				DeleteFile();
				break;

			// added DB_UpdateXML for LNA Product
			case DB_UpdateXML:
				switch (actionName) {

				case Read: {

					// uniqueNumber = ReadFromExcel(ctrlValue);
					// String Query =
					// Automation.driver.findElement(By.xpath(controlName)).getText();
					// ctrlValue = ctrlValue+"'"+policyNo+"'";
					ResultSet rs = JDBCConnection.establishDBConn_LnA("", controlName);

					rs.next();
					ctrlValue = String.valueOf(rs.getLong("COL_1"));
					rs.close();

				}

				case Write:

					try {

						// String GroupNumber =
						// Automation.driver.findElement(By.xpath(controlName)).getText();
						// ctrlValue = ctrlValue+"'"+policyNo+"'";
						String DB_Query = "";
						// uniqueNumber = ReadFromExcel(ctrlValue);
						uniqueNumber = DMProduct.ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);

						if (controlName.contains("uniqueNumber")) {
							String controlName1 = controlName.replace("uniqueNumber", uniqueNumber);
							DB_Query = controlName1;
							// String query = "select * from ContractGroup where ContractGroupNumber = " +
							// "'"+GroupNumber+"'" + "and ContractGroupTypeCT = 'Group' " ;

							// ctrlValue = query;

							// ctrlValue = policyNo;
							ResultSet rs = JDBCConnection.establishDBConn_LnA("", DB_Query);

							rs.next();
							// ctrlValue = String.valueOf(rs.getLong("COL_1"));
							DB_Query = String.valueOf(rs.getLong("ContractGroupPK"));
							rs.close();
						}

						String XML_file = (controller.controllerTransactionType.toString());

						if (XML_file.equalsIgnoreCase("DB_Query")) {

							String Filepath = System.getProperty("user.dir") + "\\" + Config.projectName
									+ "\\Resources\\Input\\WebService\\WebserviceFiles\\" + "\\"
									+ "NewTextDocument.xml";
							System.out.println("file path is " + Filepath);
							SAXReader reader = new SAXReader();
							Document document = null;
							try {
								document = (Document) reader.read(new FileInputStream(Filepath));
							} catch (Exception e) {
								e.printStackTrace();
							}

							@SuppressWarnings("unused")
							Element emp = null;

							List<Element> list = getAllChildNodes(document.getRootElement(), "GroupingSysKey");
							if (CollectionUtils.isNotEmpty(list)) {
								for (Element child : list) {
									System.out.println(child.asXML());
									Attribute systemCode = child.attribute("SystemCode");
									if (null != systemCode) {
										systemCode.setValue(DB_Query);
									}
									System.out.println(child.asXML());
									// document.getElementsByTagName("ns2:GroupingSysKey").item(1).setTextContent(ctrlValue);
								}
							}

							writeToFile(Filepath, document);
						}

						else {

							// String Filepath = Config.inputDataFilePath +
							// TransactionMapping.directoryPathFileUpload+"/"+"GroupActiveCI.xml";

							String Filepath = System.getProperty("user.dir") + "\\" + Config.projectName
									+ "\\Resources\\Input\\WebService\\WebserviceFiles\\" + "\\" + "GroupActiveCI.xml";

							// String Filepath =
							// "D:\\Auto_Selenium\\iTAFSeleniumWeb\\LNAProduct\\Resources\\Input\\WebService\\GroupActive.xml";
							// File xmlfile = new File (Filepath);
							SAXReader reader = new SAXReader();
							Document document = null;
							try {
								document = (Document) reader.read(new FileInputStream(Filepath));
							} catch (Exception e) {
								e.printStackTrace();
							}

							@SuppressWarnings("unused")
							Element emp = null;

							List<Element> list = getAllChildNodes(document.getRootElement(), "GroupingSysKey");
							if (CollectionUtils.isNotEmpty(list)) {
								for (Element child : list) {
									System.out.println(child.asXML());
									Attribute systemCode = child.attribute("SystemCode");
									if (null != systemCode) {
										systemCode.setValue(DB_Query);
									}
									System.out.println(child.asXML());
									// document.getElementsByTagName("ns2:GroupingSysKey").item(1).setTextContent(ctrlValue);
								}
							}

							writeToFile(Filepath, document);
						}
					}

					catch (Exception e) {
						// e.setStackTrace(null);
						e.printStackTrace();
						System.out.println("Error is:" + e);
					}

				}
				break;

			// Added case TaskStatus_GUW by Suhas (L&A) 08/13/2019
			case TaskStatus_GUW:

				switch (actionName) {

				case I:

					System.out.println("In TaskStatus_GW");
					if (ctrlValue.equalsIgnoreCase("")) {
						break;
					}

					Long starttime = System.currentTimeMillis();
					String CurStus;
					if (controlName.contains("+")) {
						controlName = controlName.replace("+", ctrlValue);
					}
					boolean flag = false;
					do {
						try {

							Thread.sleep(1000);
							Long endtime = System.currentTimeMillis();
							if (((endtime - starttime) / 1000) > 960) {
								System.out.println("Task has not been created within 960 seconds");
								break;

							}
							CurStus = Automation.driver.findElement(By.xpath(controlName)).getText();

							// added for clicking correct group
							/*
							 * webElement = Automation.driver.findElement(By.xpath(controlName));
							 * 
							 * webElement.click();
							 */

							System.out.println("Task Created and its status or name is : " + CurStus);

							long ttime = endtime - starttime;

							int time = (int) (ttime / 1000);

							double time1 = time / 60;
							System.out.println("Activity created in " + time1 + "Minutes");

							flag = true;
						} catch (WebDriverException e) {
							WebElement ele = Automation.driver
									.findElement(By.xpath("//*[text()='My Tasks' or text()='Recent']"));

							Actions builderMouserOver = new Actions(Automation.driver);
							builderMouserOver.moveToElement(ele).build().perform();

							(Automation.driver.findElement(By.xpath(
									"(//*[contains(@class,'tab-card-actions')])[1]/..//*[contains(@class,'-refresh')]")))
									.click();
							System.out.println("Refresh clicked");
							Thread.sleep(2000);

						}

					} while (flag == false);

				}
				break;

			case RunExe:

				switch (actionName) {
				case NC:

					if (ctrlValue.length() > 0) {

						System.out.println("IN Run Exe");
						// Runtime.getRuntime().exec("wscript Executor.vbs",);
						Runtime.getRuntime().exec("wscript.exe" + " " + Config.inputDataFilePath + Config.projectName
								+ "\\" + "Executor.vbs");
						// Runtime.getRuntime().exec("wscript.exe
						// D:\\iTAF_Version\\iTAF-1.19.1.0-dist\\Executor.vbs");

					} else {

						break;
					}

				}
				break;
			case RunMacro:

				switch (actionName) {
				case NC:

					if (ctrlValue.length() > 0) {

						System.out.println("IN Run Exe");
						// Runtime.getRuntime().exec("wscript Executor.vbs",);
						Runtime.getRuntime().exec("wscript.exe" + " " + Config.inputDataFilePath + Config.projectName
								+ "\\" + "Executorxls.vbs");
						// Runtime.getRuntime().exec("wscript.exe
						// D:\\iTAF_Version\\iTAF-1.19.1.0-dist\\Executor.vbs");

					} else {

						break;
					}

				}
				break;
			case TestCaseUP:
				switch (actionName) {
				case NC:

					try {

						if (ctrlValue.equalsIgnoreCase("Yes") || ctrlValue.length() > 0) {
							File fil = new File(Config.transactionInfo);
							// File fil = new
							// File("D:\\iTAF_Version\\iTAF-1.19.1.0-dist\\LNAProduct\\CommonResources\\UniqueNumber.xls");
							System.out.println(fil);
							FileInputStream fin = new FileInputStream(fil);

							@SuppressWarnings("resource")
							HSSFWorkbook wbook = new HSSFWorkbook(fin);
							HSSFSheet sheet = wbook.getSheet("DataSheet");
							HSSFCell TCid = (HSSFCell) controller.controllerTestCaseID;
							String tc = TCid.getStringCellValue();
							sheet.getRow(1).getCell(3).setCellValue(tc);
							FileOutputStream fout = new FileOutputStream(fil);

							wbook.write(fout);
							fout.close();

						} else {

							break;
						}
					}

					catch (Exception e) {
						System.out.println("in ex");
					}

					finally {
						System.out.println("in finally");

					}
				}

				break;

			case Opt_scroll:

				if (controlName.contains("+"))
					controlName = controlName.replace("+", ctrlValue);

				if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes") || ctrlValue.length() > 0) {
					WebElement optele = Automation.driver.findElement(By.xpath(controlName));
					((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();", optele);
				}

				break;

			case ReplaceDefault:
				switch (actionName) {
				case I:
					if (!ctrlValue.equalsIgnoreCase("null")) {

						try {
							int counter = 1;
							String temp = "";
							while (!temp.equals(ctrlValue)) {

								JavascriptExecutor js1 = (JavascriptExecutor) Automation.driver;
								js1.executeScript("arguments[0].value = '';", webElement);
								webElement.sendKeys(ctrlValue);
								Thread.sleep(2000);
								webElement.sendKeys(Keys.TAB);
								webElement.sendKeys(Keys.TAB);
								Thread.sleep(5000);
								temp = webElement.getAttribute("value");
								temp = temp.replace("$", "");
								temp = temp.replace(" ", "");
								temp = temp.replace(",", "");
								counter++;
								if (counter == 5) {
									System.out.println("Could not set correct value");
									break;
								}
							}

						} catch (StaleElementReferenceException e) {
							e.toString();
							System.out.println("Trying to recover from a stale element :" + e.getMessage());
						}
					}
					break;

				}
				break;

			case WaitForElementToBeFound:
				if (webElement == null) {
					System.out.println("Element not found.");
				} else {
					System.out.println("Element found.");
				}
				break;
			case WebButton:
				switch (actionName) {
				case FileUpload:

					// File upload making input element visible and passing the file path
					String autoitFileDir = Config.inputDataFilePath + TransactionMapping.directoryPathFileUpload;
					System.out.println("Census file path:-" + autoitFileDir);

					WebElement filename = Automation.driver.findElement(By.xpath(controlName));
					String filename1 = "arguments[0].style.height='auto'; arguments[0].style.visibility='visible';";

					((JavascriptExecutor) Automation.driver).executeScript(filename1, filename);

					Automation.driver.findElement(By.xpath(controlName)).sendKeys(autoitFileDir + "\\" + ctrlValue);
					// Automation.driver.findElement(By.xpath("//*[@id='fileName']/div/div[2]/div[2]/div[2]/input")).sendKeys("D:\\iTAF_Version\\iTAF-1.19.1.0-dist\\LNAProduct\\Resources\\Input\\LNAProduct\\"+ctrlValue);

					Thread.sleep(1000);
					// Runtime.getRuntime().exec(autoitFileDir
					// "\\FileUpload.exe "+ctrlValue);
					// Runtime.getRuntime().exec(autoitFileDir + "\\FileUpload.exe " + autoitFileDir
					// + "\\" + ctrlValue);
					// Thread.sleep(5000);
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

									Thread.sleep(500);
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
											// js.executeScript("arguments[0].focus(); arguments[0].blur(); return
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
									System.out.println("Element does not exist");
									// e1.printStackTrace();
								}
							}
						}
					} else {
						System.out.println("Don't Perform any action");
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
										System.out.println("catchclicked");
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
								// webElement = getElementByType(controlId,
								// controlName,control,imageType,ctrlValue);
								if (iCtr <= 2) {
									log.error(e1.getMessage(), e1);
								} else {
									log.error(e1.getMessage());
								}
								Thread.sleep(1000 * iCtr);

								iCtr--;
								System.out.println("Element does not exist");
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
					// Sheetal: 2/19/2019, commenting below line as its already
					// executed in flow
					// webElement = getElementByType(controlId, controlName,
					// WebHelper.control, imageType, ctrlValue);

					try {

						if (webElement.isDisplayed()) {

							// if(webElement.isEnabled() == true)

							Actions Click = new Actions(Automation.driver);
							highlightElement(webElement);
							Action Mouseclick = Click.moveToElement(webElement).clickAndHold().release().build();
							Mouseclick.perform();
							Thread.sleep(500);

							// Sheetal: 2/19/2019, commenting below line as it
							// should not be part of Webbutton->NCIF
						}

					} catch (Exception e) {
						log.error(e.getMessage(), e);
						System.out.println("Element Not Found on Page.");
						log.info("Element Not Found on Page in NCIF case");
					}

					break;
				}
				break;
			// break;

			case Edit_Enrollment:
				switch (actionName) {

				case NC:

					try {
						String GroupNumber = Automation.driver.findElement(By.xpath("//*[text()='Group Number']//pre"))
								.getText();

						System.out.println("GroupNumber:" + GroupNumber);

						DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
						Date date = new Date();
						String CurrentDate = dateFormat.format(date);

						String CSV_PATH = Config.inputDataFilePath + Config.projectName + "\\" + "E_"
								+ controller.controllerTestCaseID + ".csv";
						;

						File inputFile = new File(CSV_PATH);
						// Read existing file
						@SuppressWarnings("deprecation")
						CSVReader reader = new CSVReader(new FileReader(inputFile), ',');
						List<String[]> csvBody = reader.readAll();
						@SuppressWarnings("unused")
						String body = csvBody.toString();
						// Get CSV row column and replace with by using row and column
						for (int i = 0; i < csvBody.size(); i++) {
							String[] strArray = csvBody.get(i);
							for (int j = 0; j < strArray.length; j++) {

								if (strArray[j].equalsIgnoreCase("2500000"))

								{ // String to be replaced
									csvBody.get(i)[j] = GroupNumber; // Target replacement
								} else if (strArray[j].equalsIgnoreCase("08/15/2019")) { // String to be replaced
									csvBody.get(i)[j] = CurrentDate; // Target replacement

								}
							}
						}
						reader.close();

						// Write to CSV file which is open
						@SuppressWarnings("deprecation")
						CSVWriter writer = new CSVWriter(new FileWriter(inputFile), ',', CSVWriter.NO_QUOTE_CHARACTER);
						writer.writeAll(csvBody);
						writer.flush();
						writer.close();

					} catch (Exception e) {
						e.getMessage();
						log.error(e.getMessage());

					}

					break;
				// to update back the default values in enrollment file
				case I:

					try {
						String GroupNumber = Automation.driver.findElement(By.xpath("//*[text()='Group Number']//pre"))
								.getText();

						System.out.println("GroupNumber:" + GroupNumber);

						DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
						Date date = new Date();
						String CurrentDate = dateFormat.format(date);

						String CSV_PATH = Config.inputDataFilePath + Config.projectName + "\\" + "E_"
								+ controller.controllerTestCaseID + ".csv";
						;

						File inputFile = new File(CSV_PATH);
						// Read existing file
						@SuppressWarnings("deprecation")
						CSVReader reader = new CSVReader(new FileReader(inputFile), ',');
						List<String[]> csvBody = reader.readAll();
						@SuppressWarnings("unused")
						String body = csvBody.toString();

						for (int i = 0; i < csvBody.size(); i++) {
							String[] strArray = csvBody.get(i);
							for (int j = 0; j < strArray.length; j++) {

								if (strArray[j].equalsIgnoreCase(GroupNumber))

								{ // String to be replaced
									csvBody.get(i)[j] = GroupNumber; // Target replacement
								} else if (strArray[j].equalsIgnoreCase(CurrentDate)) { // String to be replaced
									csvBody.get(i)[j] = "08/15/2019"; // Target replacement

								}
							}
						}

						reader.close();

						// Write to CSV file which is open
						@SuppressWarnings("deprecation")
						CSVWriter writer = new CSVWriter(new FileWriter(inputFile), ',', CSVWriter.NO_QUOTE_CHARACTER);
						writer.writeAll(csvBody);
						writer.flush();
						writer.close();

					} catch (Exception e) {
						e.getMessage();
						log.error(e.getMessage());

					}

				}

				break;

			case screenshot:
				switch (actionName) {

				case NC:

					saveScreenShot();
					System.out.println("Screen shot taken");

				}

				break;

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

								String element = webElement.toString();
								System.out.println("Next Button elment is :" + element);

								System.out.println("clicked");
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

								System.out.println("catchclicked");
								log.info("WebElement:NC->catchclicked");
								break;

							} catch (Exception e1) {
								log.error(e1.getMessage(), e1);
								// webElement = getElementByType(controlId,
								// controlName,control,imageType,ctrlValue);
								Thread.sleep(3000);

								i--;
								System.out.println("Element does not exist");
								log.info("WebElement:NC->Element does not exist");
								// e1.printStackTrace();

							}
						}
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
					if (Config.projectName.equals("DistributionManagement")) {// Minaakshi
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

				case VD: // Minaakshi : Added for handling dynamic field
					// level verification
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}

					if (ctrlValue.equalsIgnoreCase(null)) {
						ctrlValue = DMProduct.ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
					} else {
						String strTemp = ctrlValue;
						if (ctrlValue.contains("$value")) {// Minaakshi :
							ctrlValue = DMProduct.ReadFromExcelUsingColumnName("", WebHelper.columnName);

						} else if (ctrlValue.contains("|*") || ctrlValue.contains("-*") || ctrlValue.contains("- *")) {// Minaakshi
							// :
							// 01-07-2019
							// Minaakshi : 04-01-2019
							String[] temp = ctrlValue.split("\\|");
							String tempCtrlValue = DMProduct.ReadFromExcelUsingColumnName(temp[0].toString(),
									WebHelper.columnName);
							ctrlValue = temp[1].toString().replace("*value", tempCtrlValue);

						} else if (ctrlValue.contains("|")) {// Minaakshi :
							// 01-06-2019
							// Minaakshi : 04-01-2019
							String[] temp = ctrlValue.split("\\|");
							ctrlValue = temp[1].toString();

						} else {
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

							System.out.println("PASS");
						} else {
							System.out.println("FAIL");

						}
					}
					break;
				// System.out.println(b);
				// System.out.println(a);

				case V:
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					if (WebVerification.isFromVerification == true) {
						currentValue = webElement.getText();
						System.out.println(currentValue);

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

				case V_text: // Sheetal - 2-5-2019
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}

					highlightElement(webElement);
					currentValue = webElement.getAttribute("value");

					if (StringUtils.equalsIgnoreCase(currentValue, null) || currentValue.equalsIgnoreCase("")) {
						currentValue = webElement.getText();
						if (StringUtils.equalsIgnoreCase(currentValue, null))

						{
							currentValue = "";
						}

					}

					if (ctrlValue.equalsIgnoreCase("pickFromUniqueNumbers")) {
						// ctrlValue = ReadFromExcel(ctrlValue);
						ctrlValue = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
					} else if (ctrlValue.equalsIgnoreCase("blank")) {
						if (currentValue.isEmpty()) {
							currentValue = ctrlValue;
						}
					}

					break;

				case V_checkboxStatus: // Sheetal - 2-5-2019
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					highlightElement(webElement);
					String checkedValue = webElement.getAttribute("checked");

					if (StringUtils.equalsIgnoreCase(checkedValue, "true")) {
						currentValue = "Checked";
					} else {
						currentValue = "Unchecked";
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

					if (StringUtils.equalsIgnoreCase(disableStatus, "true")) {
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
								System.out.println(currentValue);

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
							System.out.println("Element is not displayed");
						}

					} catch (Exception e) {
						log.error(e.getMessage(), e);
						System.out.println("Element Not Found on Page.");
					}

					break;

				case Highlight:

					WebHelperUtil.fnHighlightMe(Automation.driver,
							Automation.driver.findElement(By.xpath(controlName)));

					// WebHelper.fnHighlightMe(Automation.driver,Automation.driver.findElement(controlName));

					// Added cse V_opt by Suhas (L&A) 08/13/2019 for optional verification
				case V_opt:

					System.out.println("in v_opt");

					if (ctrlValue.length() < 0) {
						System.out.println("Ctrl value is null");
					}

					if (controlName.contains("+")) {
						controlName = controlName.replace("+", ctrlValue);
					}

				{
					if (ctrlValue.length() > 0) {

						@SuppressWarnings("unused")
						WebElement ele = null;

						WebDriverWait wait = new WebDriverWait(Automation.driver, Duration.ofSeconds(100));// Sel4
						wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath(controlName)));
						wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));

						String newele = Automation.driver.findElement(By.xpath(controlName)).getText();

						currentValue = newele;
						if (ctrlValue.equalsIgnoreCase(currentValue)) {

							System.out.println("Verification case passed");
						} else
							System.out.println("Verification case Failed");
					}
				}

					break;

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
					System.out.println("Element does not exist");
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

					System.out.println("Element Not Found");
				}

				break;

			case JSScript:
				((JavascriptExecutor) Automation.driver).executeScript(controlName, ctrlValue);

				break;

			// bhaskar
			// case IJSScript:
			// IJavascriptExecutor ijs = Automation.driver;
			// bhaskar
			case WaitForPageToLoad:

				// Automation.driver.manage().timeouts().pageLoadTimeout(120,
				// TimeUnit.SECONDS);// Sel4
				Automation.driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(120));

				WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(500));// Sel4

				WaitForPageLoad.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

				break;

			case WaitForElementToVisible: // Mrinmayee - Wait until the
				// expected value is present on
				// screen

				String temp1 = Automation.driver.findElement(By.xpath(controlName)).getText();
				System.out.println(temp1);
				int x = 200;

				while (x > 0) {

					if (temp1.equals(ctrlValue)) {
						// System.out.println("Element is present");
						System.out.println("temp1 value is:-" + temp1);
						System.out.println("ctrlValue value is:-" + ctrlValue);
						break;
					} else {
						String temp2 = Automation.driver.findElement(By.xpath(controlName)).getText();
						temp1 = temp2;
						System.out.println("temp1 value is:-" + temp1);
						System.out.println("ctrlValue value is:-" + ctrlValue);
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

					System.out.println("Element is Invisible");
				} catch (Exception e) {
					log.error(e.getMessage(), e);
					System.out.println("Element is visible");
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
				System.out.println(ctrlValue);
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
									System.out.println("Checkbox not found");
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
								System.out.println("Checkbox not found");
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
			case CloseWindow:// added this Case to bypass page loading after
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
					System.out.println("clicked");
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
						System.out.println("Element not exist");
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
					System.out.println("In method doaction debug4");
					Thread.sleep(5000);

					if (controlName.startsWith("//iframe")) {

						WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(700));// Sel4

						wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

						Automation.driver.switchTo().frame(Automation.driver.findElement(By.xpath(controlName)));

					}

					else {

						Automation.driver.switchTo().frame(controlName);

					}

					System.out.println("In method doaction debug5");
					break;

				case NC:
					System.out.println("In method doaction debug4");
					Thread.sleep(1000);

					if (controlName.startsWith("//iframe") || (controlName.startsWith("//frame"))) {

						if (Config.projectName.equalsIgnoreCase("LNAProduct")
								&& (!controlName.equalsIgnoreCase("//frame[@name='headerTwoBatch']"))
								&& (!controlName.equalsIgnoreCase("//iframe_SwitchBack_ToParent"))
								&& (!controlName.equalsIgnoreCase("//iframe_RefreshBatch&VerifyStatus")
										&& (!controlName.equalsIgnoreCase("//iframe_ParentFrame")
												&& (!controlName.equalsIgnoreCase("//frame[@name='header']"))))) {

							// WebDriverWait wait1 = new WebDriverWait(Automation.driver, 700);
							WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(700));// Sel4

							wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

							Automation.driver.switchTo().frame(0);
							Automation.driver.switchTo().frame(1);

							System.out.println("I am in main frame");
						}
						// this is to click on BatchStatus
						if (controlName.equalsIgnoreCase("//frame[@name='header']")) {
							System.out.println("in batch status frame");

							Automation.driver.switchTo().defaultContent();// this is must

							Automation.driver.switchTo().frame("flexapp_iframe");
							Automation.driver.switchTo().frame("header");

						}
						// Added for Claims Pay proc batch & Batch
						if (controlName.equalsIgnoreCase("//frame[@name='headerTwoBatch']")) {

							System.out.println("Claims Pay proc batch & Batch frame");
							Automation.driver.switchTo().frame("flexapp_iframe");
							Automation.driver.switchTo().frame("main"); // added for payout initiated batch

						}
						// add below to verify status and click on batch if needed
						if (controlName.equals("//iframe_RefreshBatch&VerifyStatus")) {

							switch (actionName) {
							case NC:

								if (ctrlValue.equalsIgnoreCase("Yes")) {

									String parwin = Automation.driver.getWindowHandle();
									Set<String> allwin = Automation.driver.getWindowHandles();

									@SuppressWarnings("unused")
									int count = allwin.size();
									for (String child : allwin) {
										if (!parwin.equalsIgnoreCase(child))
											Automation.driver.switchTo().window(child);
									}
									break;
								}
							}

							if (controlName.equals("//iframe_ParentFrame")) {

								String parwin = Automation.driver.getWindowHandle();

								String title = Automation.driver.getTitle();
								System.out.println("Title is" + title);

								Set<String> allwin = Automation.driver.getWindowHandles();

								int count = allwin.size();
								System.out.println("Total windows are " + count);

								for (String child : allwin) {

									if (!parwin.equalsIgnoreCase(child))

										Automation.driver.switchTo().window(child);

								}
								Automation.driver.switchTo().window(parwin);

								// Automation.driver.switchTo().defaultContent();
							} // System.out.println("Switched to defaultCotent");

						}

						if (controlName.equals("//iframe_SwitchBack_ToParent")) {
							// this is for swiching back to parent after checking batch status
							switch (actionName) {
							case NC:

								// this is for swiching back to parent after checking batch status
								if (ctrlValue.equalsIgnoreCase("Yes") || (ctrlValue.length() > 0)) {

									@SuppressWarnings("unused")
									String title = Automation.driver.getTitle();

									String parwin = Automation.driver.getWindowHandle();
									@SuppressWarnings("unused")
									String curtitle = Automation.driver.getTitle();

									Set<String> allwin = Automation.driver.getWindowHandles();

									@SuppressWarnings("unused")
									int count = allwin.size();

									for (String child : allwin) {

										if (!parwin.equalsIgnoreCase(child))

											Automation.driver.switchTo().window(child);

									}

								}

							}
							break;

						}

					}

					if (!Config.projectName.equalsIgnoreCase("LNAProduct")) {
						// WebDriverWait wait1 = new WebDriverWait(Automation.driver, 700);
						WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(700));// Sel4

						wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

						Automation.driver.switchTo().frame(Automation.driver.findElement(By.xpath(controlName)));

					}
					System.out.println("In method doaction debug5");
					// Automation.driver.switchTo().frame(controlName);
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
						System.out.println("Alert found on the web page");
						System.out.println(currentValue);
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
							Thread.sleep(1000);
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
				System.out.println(isCalendarDisplayed);
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
					System.out.println("Calendar not Diplayed");
				}
				// Automation.selenium.click(controlName);
				break;

			case CalendarNew:
				isCalendarDisplayed = Automation.driver.switchTo().activeElement().isDisplayed();
				System.out.println(isCalendarDisplayed);
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
							WebElement yearButton = Automation.driver.findElement(By.cssSelector("td:contains('?)"));
							yearButton.click();
							Monthyear2[1] = Integer.toString(Integer.parseInt(Monthyear2[1]) - 1);
						} else if (Integer.parseInt(Monthyear2[1]) < Integer.parseInt(dtMthYr[2])) {
							WebElement yearButton = Automation.driver.findElement(By.cssSelector("td:contains('?)"));
							yearButton.click();
							Monthyear2[1] = Integer.toString(Integer.parseInt(Monthyear2[1]) + 1);
						}
					}

					while (!WebHelper.month.equalsIgnoreCase(dtMthYr[1])) {
						if (Integer.parseInt(WebHelper.month) > Integer.parseInt(dtMthYr[1])) {
							WebElement monthButton = Automation.driver.findElement(By.cssSelector("td:contains('?)"));
							monthButton.click();
							if (Integer.parseInt(WebHelper.month) < 11) {
								WebHelper.month = "0" + Integer.toString(Integer.parseInt(WebHelper.month) - 1);
							} else {
								WebHelper.month = Integer.toString(Integer.parseInt(WebHelper.month) - 1);
							}

						} else if (Integer.parseInt(WebHelper.month) < Integer.parseInt(dtMthYr[1])) {
							WebElement monthButton = Automation.driver.findElement(By.cssSelector("td:contains('?)"));
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
					System.out.println(dateButton);
					dateButton.click();

				} else {
					System.out.println("Calendar not Diplayed");
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
				Automation.driver
						.findElement(By.xpath(xpathMonth_dm
								+ "/select[@class='ui-datepicker-month']/option[contains(.,'" + monthNum_dm + "')]"))
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
				WebHelper.wait.until(ExpectedConditions.elementToBeClickable(
						By.xpath(pathToVisibleCalendar + "/div[contains(text()," + yearEBP + ")]"))).click();
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
					@SuppressWarnings("unused")
					String parentHandle = Automation.driver.getWindowHandle();
					for (String winHandle : Automation.driver.getWindowHandles()) {
						Automation.driver.switchTo().window(winHandle);
						/*
						 * if (Automation.driver.getTitle().equalsIgnoreCase( controlName)) {
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
									System.out.println(tableRowIndex);
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
					System.out.println("logicalName----->" + logicalName);
					System.out.println("controlName----->" + controlName);

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

			case Database:// Minaakshi : Case added to fetch Unique values
				// from Database
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
						if (!readFromColName.equalsIgnoreCase("")) {
							// Minaakshi : 03-10-2018
							String tempValue = DMProduct.ReadFromExcelUsingColumnName("", readFromColName);
							sqlQuery = DMProduct.getDataFetchingSQLQuery(sqlQuery, tempValue);
						} // Minaakshi : 20-09-2018

						ResultSet rs2;
						Connection conn2 = JDBCConnection.establishPASDBConn();
						Statement st2 = conn2.createStatement();
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

						// Minaakshi : 01-03-2021
						DMProduct.writeDataToUniqueNumberSheet(ctrlValue, writeToColName, "");
					}

					break;
				}
				break;

			case CreateDynamicData:// Minaakshi : 03-10-2018
				Jacob.main(ctrlValue, "!GenerateData_Click");
				// DynamicDataGeneration.xlsm!GenerateData_Click
				break;

			case FlatFile:// Minaakshi : 01-03-2019
				switch (actionName) {
				case I:
					if (logicalName.equalsIgnoreCase("Local_Path") && !ctrlValue.equals(""))
						Local_Path = ctrlValue;
					else if (logicalName.equalsIgnoreCase("Template_Path") && !ctrlValue.equals(""))
						Template_Path = ctrlValue;
					else if (logicalName.equalsIgnoreCase("Remote_Path") && !ctrlValue.equals(""))
						Remote_Path = ctrlValue;
					else if (logicalName.equalsIgnoreCase("EffectiveDate") && !ctrlValue.equals(""))
						EffectiveDate = ctrlValue;
					else if (logicalName.equalsIgnoreCase("EffectiveDate1") && !ctrlValue.equals(""))
						EffectiveDate1 = ctrlValue;
					else if (logicalName.equalsIgnoreCase("RiskCommencementDate") && !ctrlValue.equals(""))
						RiskCommencementDate = ctrlValue;
					else if (logicalName.equalsIgnoreCase("RiskCommencementDate1") && !ctrlValue.equals(""))
						RiskCommencementDate1 = ctrlValue;
					else if (logicalName.equalsIgnoreCase("PolicyEffectiveToDate") && !ctrlValue.equals(""))
						PolicyEffectiveToDate = ctrlValue;
					else if (logicalName.equalsIgnoreCase("ProposalApplDate") && !ctrlValue.equals(""))
						ProposalApplDate = ctrlValue;
					else if (logicalName.equalsIgnoreCase("ProposalSubmDate") && !ctrlValue.equals(""))
						ProposalSubmDate = ctrlValue;
					else if (logicalName.equalsIgnoreCase("TemplateFileName") && !ctrlValue.equals(""))
						TemplateFileName = ctrlValue;
					else if (logicalName.equalsIgnoreCase("FileNameToBeCreated") && !ctrlValue.equals(""))
						FileNameToBeCreated = ctrlValue;
					else if (logicalName.equalsIgnoreCase("FileNameCopiedToLocal") && !ctrlValue.equals(""))
						FileNameCopiedToLocal = ctrlValue;
					// Minaakshi : 01-05-2019
					else if (logicalName.equalsIgnoreCase("PolicyNumber") && !ctrlValue.equals(""))
						PolicyNumber = ctrlValue;
					// Minaakshi : 01-05-2019
					else if (logicalName.equalsIgnoreCase("PolicyFileLoadNumber") && !ctrlValue.equals(""))
						PolicyFileLoadNumber = ctrlValue;
					// Minaakshi : 01-05-2019
					else if (logicalName.equalsIgnoreCase("PremiumFileLoadNumber") && !ctrlValue.equals(""))
						PremiumFileLoadNumber = ctrlValue;
					break;

				case Read:
					if (logicalName.equalsIgnoreCase("EntityCode1") && !ctrlValue.equals(""))
						EntityCode1 = DMProduct.ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
					else if (logicalName.equalsIgnoreCase("EntityCode2") && !ctrlValue.equals(""))
						EntityCode2 = DMProduct.ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
					else if (logicalName.equalsIgnoreCase("EntityCode3") && !ctrlValue.equals(""))
						EntityCode3 = DMProduct.ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
					else if (logicalName.equalsIgnoreCase("EntityCode4") && !ctrlValue.equals(""))
						EntityCode4 = DMProduct.ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
					else if (logicalName.equalsIgnoreCase("EntityCode5") && !ctrlValue.equals(""))
						EntityCode5 = DMProduct.ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);

					break;

				case CreateFile:

					if (logicalName.equalsIgnoreCase("CreateFile") && ctrlValue.equals("Yes")) {
						copySource = new File(
								Config.inputDataFilePath.toString() + Template_Path + "//" + TemplateFileName);
						copyDestination = new File(
								Config.inputDataFilePath.toString() + Local_Path + "//" + FileNameToBeCreated);
						FileUtils.copyFile(copySource, copyDestination);

						// Open created and replace Transaction Date and
						// EntityCodes

						String originalFileContent = "";
						BufferedReader reader = null;
						BufferedWriter writer = null;

						try {
							reader = new BufferedReader(new FileReader(copyDestination));
							String currentReadingLine = reader.readLine();
							while (currentReadingLine != null) {
								originalFileContent += currentReadingLine;
								originalFileContent += System.lineSeparator();
								currentReadingLine = reader.readLine();
							}

							String modifiedFileContent = originalFileContent.replaceFirst("[\n\r]+$", "");
							modifiedFileContent = modifiedFileContent.replaceAll("\\$DATE", EffectiveDate);

							if (!EffectiveDate1.equals(""))
								modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE", EffectiveDate1);

							if (!RiskCommencementDate1.equals(""))
								modifiedFileContent = modifiedFileContent.replaceAll("\\$RC_DATE1",
										RiskCommencementDate1);

							if (!RiskCommencementDate.equals(""))
								modifiedFileContent = modifiedFileContent.replaceAll("\\$RC_DATE",
										RiskCommencementDate);

							if (!PolicyEffectiveToDate.equals(""))
								modifiedFileContent = modifiedFileContent.replaceAll("\\$PE_DATE",
										PolicyEffectiveToDate);

							if (!ProposalApplDate.equals(""))
								modifiedFileContent = modifiedFileContent.replaceAll("\\$PA_DATE", ProposalApplDate);

							if (!ProposalSubmDate.equals(""))
								modifiedFileContent = modifiedFileContent.replaceAll("\\$PS_DATE", ProposalSubmDate);

							if (transactionType.toString().equalsIgnoreCase("PolicyInterface")) {// Minaakshi
								// :
								// 01-05-2019
								if (!StringUtils.equalsIgnoreCase(WebHelperLNA.PolicyFileLoadNumber, "")
										|| !StringUtils.equalsIgnoreCase(WebHelperLNA.PolicyFileLoadNumber, null)) {
									modifiedFileContent = modifiedFileContent.replaceAll("\\$POLFLN",
											PolicyFileLoadNumber);
								}
							}

							if (transactionType.toString().equalsIgnoreCase("PremiumInterface")) {// Minaakshi
								// :
								// 01-05-2019
								if (!StringUtils.equalsIgnoreCase(WebHelperLNA.PremiumFileLoadNumber, "")
										|| !StringUtils.equalsIgnoreCase(WebHelperLNA.PremiumFileLoadNumber, null)) {
									modifiedFileContent = modifiedFileContent.replaceAll("\\$PREFLN",
											PremiumFileLoadNumber);
								}
							}

							if (!StringUtils.equalsIgnoreCase(WebHelperLNA.PolicyNumber, "")
									|| !StringUtils.equalsIgnoreCase(WebHelperLNA.PolicyNumber, null))// Minaakshi
								// :
								// 01-05-2019
								modifiedFileContent = modifiedFileContent.replaceAll("\\$POL_NUMBER", PolicyNumber);
							if (!EntityCode1.equals(""))
								modifiedFileContent = modifiedFileContent.replaceAll("\\$ENTITYCODE1", EntityCode1);

							if (!EntityCode2.equals(""))
								modifiedFileContent = modifiedFileContent.replaceAll("\\$ENTITYCODE2", EntityCode2);

							if (!EntityCode3.equals(""))
								modifiedFileContent = modifiedFileContent.replaceAll("\\$ENTITYCODE3", EntityCode3);

							if (!EntityCode4.equals(""))
								modifiedFileContent = modifiedFileContent.replaceAll("\\$ENTITYCODE4", EntityCode4);

							if (!EntityCode5.equals(""))
								modifiedFileContent = modifiedFileContent.replaceAll("\\$ENTITYCODE5", EntityCode5);

							writer = new BufferedWriter(new FileWriter(copyDestination));

							writer.write(modifiedFileContent);
						} catch (IOException e) {
							log.error(e.getMessage(), e);
							e.printStackTrace();
						} finally {

							try {
								if (reader != null) {
									reader.close();
								}
								if (writer != null) {
									writer.close();
								}
							} catch (IOException e) {
								log.error(e.getMessage(), e);
								e.printStackTrace();
							}
						}
					}

					break;

				case CopyFilesToServer:
					if (logicalName.equalsIgnoreCase("FromLocalToServer") && ctrlValue.equals("Yes")) {
						source = Config.inputDataFilePath.toString() + Local_Path;
						destination = Remote_Path;
						DMProduct.CopyFilesToServer(source, destination, FileNameToBeCreated);
					}

					break;
				case CopyFilesToLocal:
					if (logicalName.equalsIgnoreCase("FromServerToLocal") && ctrlValue.equals("Yes")) {

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

			/*
			 * case CheckBatchStatus:// Minaakshi : 01-05-2019 String batchStatus = "";
			 * startTime_batch = System.nanoTime();
			 * 
			 * String SQLQuery = ctrlValue;
			 * 
			 * ResultSet rs3; Connection conn3 = JDBCConnection.establishPASDBConn();
			 * Statement st3 = conn3.createStatement(); rs3 = st3.executeQuery(SQLQuery);
			 * rs3.next();
			 * 
			 * // This loop is for waiting record to be arrived in JOB_SCHEDULE // table
			 * while (rs3.equals("") || rs3.equals(null)) { rs3 =
			 * st3.executeQuery(SQLQuery); rs3.next(); }
			 * 
			 * // This loop is for waiting status of record to be COMPLETED while
			 * ((System.nanoTime() - startTime_batch) < 25 * 60 * NANOSEC_PER_SEC) {
			 * 
			 * if (ctrlValue.contains("BATCH_END_REASON")) { toBeFetchedDBColName =
			 * "BATCH_END_REASON"; expectedDBStatus = "BATCH_COMPLETED"; } else {
			 * toBeFetchedDBColName = "JOB_STATUS"; expectedDBStatus = "COMPLETED"; }
			 * 
			 * rs3 = st3.executeQuery(SQLQuery); rs3.next();
			 * 
			 * batchStatus = String.valueOf(rs3.getString(toBeFetchedDBColName));
			 * 
			 * if (batchStatus.equalsIgnoreCase(expectedDBStatus)) { rs3.close();
			 * st3.close(); JDBCConnection.closeConnection(conn3); Thread.sleep(10000);
			 * break; }
			 * 
			 * }
			 * 
			 * break;
			 */

			case CheckBatchStatus:
				String batchStatus = "";
				startTime_batch = System.nanoTime();
				String SQLQuery = ctrlValue;
				ResultSet rs3 = null;
				Connection conn3 = null;
				Statement st3 = null;

				try {
					conn3 = JDBCConnection.establishPASDBConn();
					st3 = conn3.createStatement();
					rs3 = st3.executeQuery(SQLQuery);
					rs3.next();
					// This loop is for waiting record to be arrived in JOB_SCHEDULE
					// table
					/* while (rs3.equals("") || rs3.equals(null)) { */// Sel4
					while (!rs3.next()) {
						rs3 = st3.executeQuery(SQLQuery);
						rs3.next();
					}

					// This loop is for waiting status of record to be COMPLETED
					while ((System.nanoTime() - startTime_batch) < 25 * 60 * NANOSEC_PER_SEC) {
						if (ctrlValue.contains("BATCH_END_REASON")) {
							toBeFetchedDBColName = "BATCH_END_REASON";
							expectedDBStatus = "BATCH_COMPLETED";
						} else {
							toBeFetchedDBColName = "JOB_STATUS";
							expectedDBStatus = "COMPLETED";
						}
						rs3 = st3.executeQuery(SQLQuery);
						rs3.next();
						batchStatus = String.valueOf(rs3.getString(toBeFetchedDBColName));

						if (batchStatus.equalsIgnoreCase(expectedDBStatus)) {
							break;
						}
					}
				} catch (Exception e) {
					e.printStackTrace();
				} finally {
					// Close ResultSet, Statement, and Connection in finally block to ensure they
					// are closed
					if (rs3 != null) {
						try {
							rs3.close();
						} catch (SQLException e) {
							e.printStackTrace();
						}
					}
					if (st3 != null) {
						try {
							st3.close();
						} catch (SQLException e) {
							e.printStackTrace();
						}
					}
					if (conn3 != null) {
						JDBCConnection.closeConnection(conn3);
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
				System.out.println("in sikulitype");
				System.out.println("controlName is:" + controlName);
				Pattern image1 = new Pattern(controlName);
				sikuliScreen.type(image1, ctrlValue);
				break;

			case SikuliButton:
				System.out.println("in sikuliButton");
				System.out.println("controlName is:" + controlName);
				Pattern image2 = new Pattern(controlName);
				sikuliScreen.click(image2);
				System.out.println("Done");
				break;
			// bhaskar
			case Slider:
				if (controlName.contains("+"))
					controlName = controlName.replace("+", ctrlValue);

				if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes") || ctrlValue.length() > 0) {
					WebElement slider = Automation.driver.findElement(By.xpath(controlName));
					Thread.sleep(3000);
					Actions moveSlider = new Actions(Automation.driver);
					Action actionslider = moveSlider.dragAndDropBy(slider, 30, 0).build();
					actionslider.perform();
				}
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
					System.out.println("Field does not Exist");
				}

				break;

			case MaskedInputDate:
				if (!ctrlValue.equalsIgnoreCase("null")) // changed by asif
				{

					webElement.sendKeys(Keys.ENTER);

					js.executeScript("arguments[0].value = '';", webElement);

					webElement.sendKeys(ctrlValue);

				} else {
					System.out.println("Don't Perform any action");
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

			case ScrollToElement: // Sheetal-2-5-2019. Had to create new due to
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

						System.out.println("Element Not Found on Page.");
					}
					break;
				}
				break;

			case WebService_LNA:

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
				case T:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {

					}

					else
						;

					URL wsdlUrl1 = new URL(WebHelper.wsdl_url);
					String wsdlUr = wsdlUrl1.toString();
					String user = "PAS";
					String password = "PAS";
					String wscycledate = null;
					String requestXML = Config.inputDataFilePath + "WebService\\" + "WebserviceFiles\\"
							+ WebHelper.request_xml + ".xml";
					WebService.callWebService(wscycledate, wsdlUr, requestXML, user, password);
				}
				break;

			case CreateJSON_Claims:

				switch (actionName) {

				case Read: {

					ResultSet rs = JDBCConnection.establishDBConn_LnA("", controlName);

					rs.next();
					ctrlValue = String.valueOf(rs.getLong("COL_1"));
					rs.close();

				}

				case Write:

					try {

						uniqueNumber = DMProduct.ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);

						if (controlName.contains("uniqueNumber")) {
							String controlName1 = controlName.replace("uniqueNumber", uniqueNumber);
							ResultSet rs = JDBCConnection.establishDBConn_LnA("", controlName1);
							ResultSetMetaData md = rs.getMetaData();
							int columns = md.getColumnCount();
							Map<String, Object> row = new HashMap<String, Object>(columns);
							while (rs.next()) {

								for (int j = 1; j <= columns; ++j) {
									row.put(md.getColumnName(j), rs.getObject(j));
								}
								// rows.put(key, value);

							}
							rs.close();
							String Filepath = System.getProperty("user.dir") + "\\" + Config.projectName
									+ "\\Resources\\Input\\WebService\\" + "\\" + "ClaimStatusUpdateJSON.json";
							System.out.println("file path is " + Filepath);
							// Added by deepthi to format and create JSON // 22-09-2019
							ObjectMapper mapper = new ObjectMapper();
							try (FileWriter file = new FileWriter(Filepath)) {
								String json = mapper.writeValueAsString(row);
								json = "{\"claimDisbursementFeedback\": [" + json + "]}";
								System.out.println(json);
								file.write(json);
								file.flush();

							} catch (IOException e) {
								e.printStackTrace();
							}
						}
					} catch (Exception e) {
						e.printStackTrace();
						System.out.println("Error is:" + e);
					}
				}
				break;

			case WebService_Claims:

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
				case T:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}

					else
						;

					URL wsdlUrl1 = new URL(WebHelper.wsdl_url);
					String wsdlUr = wsdlUrl1.toString();

					InputStream input = new FileInputStream("./Database.properties");

					String path = input.toString();
					System.out.println("file path:-" + path);
					Properties p = new Properties();
					p.load(input);
					String user = p.getProperty("UserId");
					// user = user1.trim();
					String password = p.getProperty("Password");
					// String port = p.getProperty("Port");
					String requestXML = Config.inputDataFilePath + "WebService\\" + WebHelper.request_xml + ".json";
					// WebService.callWebServiceJSON(wsdlUr, requestXML, user, password);
					WebHelperLNA.callWebServiceJSON(wsdlUr, requestXML, user, password);
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

				case T:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					WebService.callWebService();

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
					System.out.println("Old file DocumentPackaging.pdf deleted successfully");
				}

				// Below click downloads the document
				Actions MousebuilderClick1 = new Actions(Automation.driver);
				highlightElement(webElement);
				Action MouseclickAction1 = MousebuilderClick1.moveToElement(webElement).clickAndHold().release()
						.build();
				MouseclickAction1.perform();

				break;

			case RenameDocument:

				String mainWindow = Automation.driver.getWindowHandle();

				String temp2[] = ctrlValue.split(";");

				String downloadedFileName = Config.actualPdfDownloadPath + "\\" + temp2[0];
				String renameFileNameTo = Config.actualPdfDownloadPath + "\\" + temp2[1];

				File oldFile = new File(downloadedFileName);
				File newFile = new File(renameFileNameTo);

				if (newFile.exists()) {
					newFile.delete();
					System.out.println("Old file " + temp2[1] + " deleted successfully");
				}

				oldFile.renameTo(newFile);

				String timeStampForFileBackup = new SimpleDateFormat("ddMMMMyyyy_HH_mm_ss").format(new Date());
				String backUpFileName = renameFileNameTo.replace(".pdf", "") + "_" + timeStampForFileBackup + ".pdf";

				File backUpFile = new File(backUpFileName);

				Files.copy(newFile, backUpFile);

				Set<String> handlers1 = null;
				handlers1 = Automation.driver.getWindowHandles();
				for (String handler2 : handlers1) {
					if (!mainWindow.equalsIgnoreCase(handler2)) {
						Automation.driver.switchTo().window(handler2);
						Automation.driver.close();
					}
				}
				Automation.driver.switchTo().window(mainWindow);
				break;

			case OpenAPI:// Mandar //Minaakshi : 01-06-2019
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, transactionType.toString());
			default:
				log.info("U r in Default");
				break;

			}
		} catch (WebDriverException we) {
			log.error(we.getMessage(), we);
			throw new Exception("Error Occurred from Do Action " + controlName + we.getMessage());
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			throw new Exception(e.getMessage());
		}

		// Mrinmayee - skip the step for which control value is blank end

		// TM-02/02/2015: Radio button found ("F") & AJAX control ("VA")
		if ((action.equalsIgnoreCase("V") || action.equalsIgnoreCase("F") || action.equalsIgnoreCase("PopUpCount")
				|| action.equalsIgnoreCase("VerifyPopUpElement") || action.equalsIgnoreCase("VA")
				|| action.toString().equalsIgnoreCase("VD") || action.toString().equalsIgnoreCase("V_text")
				|| action.toString().equalsIgnoreCase("V_availabilityStatus")
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
		WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(30));
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
			System.out.println(elementProperties.get(i).toString());
		}
		return elementProperties;
	}

	public static void saveScreenShot() {
		if (!(Automation.driver instanceof TakesScreenshot)) {

			System.out.println(
					"Not able to take screenshot: Current WebDriver does not support TakesScreenshot interface.");
			return;
		}

		File scrFile;
		try {

			scrFile = ((TakesScreenshot) Automation.driver).getScreenshotAs(OutputType.FILE);
			System.out.println("Screen shot path: " + scrFile);
			// Automation.driver.findElement(By.xpath("//body")).sendKeys(Keys.F5);
			Automation.driver.navigate().refresh();
			System.out.println("Page refreshed post test failure");

		} catch (Exception e) {
			log.error(e.getMessage(), e);
			System.out.println("Taking screenshot failed for: " + webDriver.getReport().getTestcaseId());
			// e.printStackTrace();
			return;
		}
		String date = null;

		if (StringUtils.isNotBlank(webDriver.getReport().getFromDate()))
			date = webDriver.getReport().getFromDate().replaceAll("[-/: ]", "");
		else
			webDriver.getReport().setFromDate(Config.dtFormat.format(new Date()));

		String fileName = webDriver.getReport().getTestcaseId() + "_" + webDriver.getReport().getTrasactionType() + "_"
				+ date;
		// TM:19/01/2015 - Changes made to save screenshots in jpeg format
		// rather that png since they are heavier
		// String location = System.getProperty("user.dir")
		// +"\\Resources\\Results\\ScreenShots\\"+ fileName+".jpeg";
		String location = Config.resultFilePath + "\\ScreenShots\\" + fileName + ".jpeg";
		// bhaskar
		controller.testDescription = location;
		// bhaskar
		webDriver.getReport().setScreenShot("file:\\\\" + location);

		try {

			FileUtils.copyFile(scrFile, new File(location));
			controller.FailScreen = new File(location).getAbsolutePath(); // Sheetal:
			// 2/20/2019

		} catch (IOException e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
			return;
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

				System.out.println(conn.getConnectTimeout());

				@SuppressWarnings("unused")
				int header = conn.getResponseCode();

				switch (conn.getResponseCode()) {
				case HttpsURLConnection.HTTP_OK:
					System.out.println("200OK");

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
			System.out.println("Timeout after waiting for Vaadin call to finish.");
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


	public static Element getChild(Element root, String elementName) {
		List<Element> elementList = new ArrayList<Element>();
		for (Iterator<Node> it = root.nodeIterator(); it.hasNext();) {
			Node node = it.next();
			if (node instanceof Element) {
				Element element = (Element) node;
				getAllChildren(element, elementName, elementList);
			}
		}
		return elementList.size() > 0 ? elementList.get(0) : null;
	}


	public static List<Element> getAllChildNodes(Element root, String elementName) {
		List<Element> elementList = new ArrayList<Element>();
		for (Iterator<Node> it = root.nodeIterator(); it.hasNext();) {
			Node node = it.next();
			if (node instanceof Element) {
				Element element = (Element) node;
				getAllChildren(element, elementName, elementList);
			}
		}
		return elementList;
	}

	/**
	 * 
	 * @param root
	 * @param elementName
	 * @param elementList
	 * @return
	 */

	public static List<Element> getAllChildren(Element root, String elementName, List<Element> elementList) {
		for (Iterator<Node> it = root.nodeIterator(); it.hasNext();) {
			Node node = it.next();
			if (node instanceof Element) {
				Element element = (Element) node;
				if (element.getName().equals(elementName)) {
					elementList.add(element);
					// break;
				}
				getAllChildren(element, elementName, elementList);
			}

		}

		return elementList;
	}

	/**
	 * 
	 * @param element
	 * @param elementName
	 * @return
	 */
	public static Element getImmediateChild(Element element, String elementName) {
		for (Object child : element.elements()) {
			Element childElement = (Element) child;
			if (childElement.getName().equals(elementName)) {
				return childElement;
			}
		}
		return null;
	}

	public static List<Element> getImmediateChildren(Element element, String elementName) {
		List<Element> elementList = new ArrayList<Element>();
		for (Object child : element.elements()) {
			Element childElement = (Element) child;
			if (childElement.getName().equals(elementName)) {
				elementList.add(childElement);
			}
		}
		return elementList;
	}

	private static void writeToFile(String filePath, Document document) {
		FileOutputStream fos;
		try {
			fos = new FileOutputStream(filePath);
			OutputFormat format = OutputFormat.createPrettyPrint();
			XMLWriter writer = new XMLWriter(fos, format);
			writer.write(document);
			writer.flush();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
