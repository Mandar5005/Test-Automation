package com.majesco.itaf.main;

import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.time.Duration;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Iterator;
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
import org.openqa.selenium.Dimension;
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
import com.google.common.base.Function;
import com.majesco.itaf.batch.RunBatch;
import com.majesco.itaf.util.CalendarSnippet;
import com.majesco.itaf.util.CommonExpectedConditions;
import com.majesco.itaf.util.Constants;
import com.majesco.itaf.util.DMProduct;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.util.JDBCConnection;
import com.majesco.itaf.util.Jacob;
import com.majesco.itaf.util.PDFComparisonUtil;
import com.majesco.itaf.util.WaitTool;
import com.majesco.itaf.util.XmlComparisonUtil;
import com.majesco.itaf.verification.WebVerification;
import com.majesco.itaf.vo.Reporter;
import com.majesco.itaf.webservice.WebService;

public class WebHelperPAS {

	private final static Logger log = LogManager.getLogger(WebHelperPAS.class.getName());
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
	public static Wait<WebDriver> waitForElementPresence;
	public static Wait<WebDriver> dmgebtWait;
	public static String Local_Path = "";
	public static String Template_Path = "";
	public static String Remote_Path = "";
	public static String Remote_Path_Outbound = "";
	public static String EffectiveDate = "";
	public static String RiskCommencementDate = "";
	public static String PolicyEffectiveToDate = "";
	public static String ProposalApplDate = "";
	public static String ProposalSubmDate = "";
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
	static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	static MainController controller = ObjectFactory.getMainController();

	static {
		if (!StringUtils.equalsIgnoreCase(Config.executionScope, "only_api")) {
			WebHelper.wait = new WebDriverWait(Automation.driver, Duration.ofSeconds(Integer.parseInt(Config.timeOut)));
		}
	}

	static void implementWait() {
		dmgebtWait = new FluentWait<>(Automation.driver)
				.withTimeout(Duration.ofSeconds(Integer.parseInt(Config.timeOut))).pollingEvery(Duration.ofSeconds(3))
				.ignoring(NoSuchElementException.class);

		waitForElementPresence = new FluentWait<>(Automation.driver).withTimeout(Duration.ofSeconds(30))
				.pollingEvery(Duration.ofSeconds(3)).ignoring(NoSuchElementException.class);

	}

	public static void GetCellInfo(String FilePath, Row rowValues, int valuesRowIndex, int valuesRowCount)
			throws IOException {
		String logicalName = "";
		boolean reportWrittenAfterException = false;
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
			log.info("TestCaseID: " + controller.controllerTestCaseID.toString() + ", TransactionName: "
					+ controller.controllerTransactionType.toString());
			int rowCount = sheetStructure.getLastRowNum() + 1;
			Sheet headerValues = ExcelUtility.GetSheet(FilePath, "Values");
			String fromDate = Config.dtFormat.format(WebHelper.frmDate);
			webDriver.getReport().setFromDate(fromDate);
			WebHelper.structureHeader = WebHelperUtil.getValueFromHashMap(sheetStructure);
			WebHelper.columnName = null;
			int dynamicIndexNumber;// Added for Action Loop
			String imageType, indexVal, controlName, executeFlag, action, controltype = "", controlID, dynamicIndex,
					newDynamicIndex, rowNo, colNo;// newly//columnName1, CompareText
			webDriver.getReport().setMessage("");
			webDriver.getReport().setStatus("PASS");
			for (int rowIndex = 1; rowIndex < rowCount && !controller.pauseExecution; rowIndex++) {
				ctrlValue = "";
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
					rowNo = WebHelperUtil.getCellData("RowNo", sheetStructure, rowIndex, WebHelper.structureHeader);
					colNo = WebHelperUtil.getCellData("ColumnNo", sheetStructure, rowIndex, WebHelper.structureHeader);
					dynamicIndex = WebHelperUtil.getCellData("DynamicIndex", sheetStructure, rowIndex,
							WebHelper.structureHeader);
					log.info("iTAF:" + logicalName + " " + rowIndex);//comment this line if you do not want the numbers to be printed in logs.
					log.info("Logical Name is "+  ": "+ logicalName);
					if (controlName.toLowerCase().contains("ignorecase")) {
						String[] tempArr = controlName.split(";;");
						controlName = tempArr[1].toLowerCase();
						controlName = controlName.replace("text()",
								"translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')");
					}
					// end

					// Below code has been written To Handle condition for multiple rows in excel
					// sheet
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
							//log.info("actualValue is Null");//Mandar
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
							WebHelper.transactionType = rowValues.getCell(
									Integer.parseInt(WebHelper.valuesHeader.get("TransactionType").toString()));

							if (WebHelper.testcaseID == null) {
								testCase = "";
							} else {
								testCase = WebHelper.testcaseID.toString();
							}
							transactionType = rowValues.getCell(WebHelper.valuesHeader.get("TransactionType"));
						}

						if ((action.equals("I") && !ctrlValue.isEmpty()) || (action.equals("V") && !ctrlValue.isEmpty())
								|| (action.equals("InputDate") && !ctrlValue.isEmpty())
								|| (action.equals("InputBusinessDate") && !ctrlValue.isEmpty())
								|| (action.equals("ActionClick") && !ctrlValue.isEmpty())
								|| (action.equals("V_availabilityStatus") && !ctrlValue.isEmpty())
								|| (action.equals("V_text") && !ctrlValue.isEmpty())
								|| (action.equals("V_format") && !ctrlValue.isEmpty())
								|| (action.equals("V_attributeValue") && !ctrlValue.isEmpty())
								|| (action.equals("V_edit") && !ctrlValue.isEmpty())
								|| (action.equals("V_date") && !ctrlValue.isEmpty())
								|| (action.equals("V_disableStatus") && !ctrlValue.isEmpty())
								|| (action.equals("V_checkboxStatus") && !ctrlValue.isEmpty())
								|| (action.equals("V_toggleDisableStatus") && !ctrlValue.isEmpty())
								|| (action.equals("V_toggleCheckboxStatus") && !ctrlValue.isEmpty())
								|| (action.equals("VFPresence") && !ctrlValue.isEmpty())
								|| (action.equals("VFormPresence") && !ctrlValue.isEmpty())
								|| (action.equals("V_toggleOptionStatus") && !ctrlValue.isEmpty())
								|| (action.equals("V_buttonOptionStatus") && !ctrlValue.isEmpty())
								|| (action.equals("V_backgroundStatus") && !ctrlValue.isEmpty())
								|| (action.equals("V_toggleOptionDisableStatus") && !ctrlValue.isEmpty())
								|| (action.equals("InputifExist") && !ctrlValue.isEmpty())
								|| (action.equals("SetCheckbox") && !ctrlValue.isEmpty())
								|| (action.equals("VFormCount") && !ctrlValue.isEmpty())
								|| (!action.equals("I") && !action.equals("InputDate")
										&& !action.equals("InputBusinessDate") && !action.equals("V")
										&& !action.equals("V_availabilityStatus") && !action.equals("V_edit")
										&& !action.equals("V_text") && !action.equals("V_date")
										&& !action.equals("V_toggleOptionStatus")
										&& !action.equals("V_buttonOptionStatus")
										&& !action.equals("V_backgroundStatus")
										&& !action.equals("V_toggleOptionDisableStatus")
										&& !action.equals("V_attributeValue") && !action.equals("V_format")
										&& !action.equals("V_disableStatus") && !action.equals("VFPresence")
										&& !action.equals("V_checkboxStatus") && !action.equals("V_toggleDisableStatus")
										&& !action.equals("V_toggleCheckboxStatus") && !action.equals("SetCheckbox")
										&& !action.equals("VFormCount") && !action.equals("VerifyPopUpElement")
										&& !action.equals("ActionClick"))) {

							if (!controltype.startsWith("Sikuli")) {
								if (!action.equalsIgnoreCase("LOOP") && !controltype.equalsIgnoreCase("Wait")
										&& !controltype.equalsIgnoreCase("Wait_DM")
										// Minaakshi : 05-02-2019
										&& !action.equalsIgnoreCase("END_LOOP")
										&& !controltype.equalsIgnoreCase("Browser")
										&& !controltype.equalsIgnoreCase("CloseBrowser")
										&& !controltype.equalsIgnoreCase("BrowserSwitchUsingURL")
										&& !controltype.equalsIgnoreCase("HandleWSFailure")
										&& !controltype.equalsIgnoreCase("AttributeisPresent")
										&& !controltype.equalsIgnoreCase("CloseAndLaunchNewBrowser")
										&& !controltype.equalsIgnoreCase("ClearUniqueNumberFileContent")
										&& !controltype.equalsIgnoreCase("PageRefresh")
										&& !controltype.equalsIgnoreCase("MaximizeBrowser")
										&& !controltype.equalsIgnoreCase("WaitForPageToLoad")
										&& !controltype.equalsIgnoreCase("WaitUntilElementInvisible")
										&& !controltype.equalsIgnoreCase("filedownload")
										&& !controltype.equalsIgnoreCase("NewBrowser")
										&& !controltype.equalsIgnoreCase("Window")
										&& !controltype.equalsIgnoreCase("Alert")
										&& !controltype.equalsIgnoreCase("URL")
										&& !controltype.equalsIgnoreCase("RunJBeamBatch")
										&& !controltype.equalsIgnoreCase("Screenshot")
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
										&& !controltype.equalsIgnoreCase("CreateDynamicData")
										// Minaakshi : 03-10-2018
										&& !controltype.equalsIgnoreCase("FlatFile")
										&& !controlID.equalsIgnoreCase("XML") && !controltype.startsWith("Process")
										&& !controltype.startsWith("Destroy") && !controltype.startsWith("ReadSikuli")
										&& !controltype.equalsIgnoreCase("WebService") && !action.equalsIgnoreCase("VA")
										&& (ctrlValue == null || !ctrlValue.equalsIgnoreCase("IGNORE"))// Minaakshi
										&& !controltype.equalsIgnoreCase("IFrame") && !action.equalsIgnoreCase("LOOP")
										&& !controltype.equalsIgnoreCase("RowNumbersToExecute")
										&& !controltype.equalsIgnoreCase("TextboxToTextbox")
										&& !controltype.equalsIgnoreCase("InputUsingRefElements")
										&& !controltype.equalsIgnoreCase("WaitForElementToVisible")
										&& !controltype.equalsIgnoreCase("Wait_IfValue")
										&& !controltype.equalsIgnoreCase("WaitToLoad")
										&& !action.equalsIgnoreCase("Billing")
										&& !controltype.equalsIgnoreCase("D_Wait")
										&& !action.equalsIgnoreCase("UpdateDBScript")
										&& !controltype.equalsIgnoreCase("SPNR_Wait")
										&& !action.equalsIgnoreCase("VFormPresence")
										&& !action.equalsIgnoreCase("FormVerificationDetails")
										&& !controltype.equalsIgnoreCase("WindowAlertOk")
										&& !controltype.equalsIgnoreCase("RenameDocument")
										&& !controltype.equalsIgnoreCase("CopyFromServerLocation")
										&& !controltype.equalsIgnoreCase("IgnoreString")
										&& !controltype.equalsIgnoreCase("PDFDocumentCompare")
										&& !controltype.equalsIgnoreCase("PDFComparisonMode")
										&& !controltype.equalsIgnoreCase("MoveDocument")
										&& !controltype.equalsIgnoreCase("WaitTillFileDownload")
										&& !controltype.equalsIgnoreCase("MoveXMLDocument")
										&& !controltype.equalsIgnoreCase("UnzipFolderAndRename")
										&& !controltype.equalsIgnoreCase("XMLIgnoreTags_AttributeLevel")
										&& !controltype.equalsIgnoreCase("OpenAPI")
										&& !controltype.equalsIgnoreCase("RenameFile")
										&& !controltype.equalsIgnoreCase("MoveFile")
										&& !action.equalsIgnoreCase("WriteEntityReference")
										&& !controltype.equalsIgnoreCase("XMLIgnoreTags_NodeLevel")
										&& !controltype.equalsIgnoreCase("XMLCompare")
										&& !controltype.equalsIgnoreCase("ExecuteExe")) {
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
						// Update $value from controlName
						// case 1
						if ((controlID.equalsIgnoreCase("XPathValueMultiWithInput")
								|| controlID.equalsIgnoreCase("XPathValueMultiWithInput_p"))
								&& !ctrlValue.equalsIgnoreCase("")) {

							String tempcontrolValue = ctrlValue;
							String[] tempValues2 = tempcontrolValue.split(";;");
							int l = tempValues2.length;
							ctrlValue = tempValues2[l - 1];

							String tempReplaceString = controlName;
							int k = 1;
							for (int j = 0; j < tempValues2.length - 1; j++) {

								if (tempValues2[j].trim().equals("pickFromUniqueNumbers")) {
									tempValues2[j] = WebHelperUtil.ReadFromExcel(tempValues2[j].trim(),
											WebHelper.columnName);
								}

								tempReplaceString = tempReplaceString.replace("$" + k + "value", tempValues2[j].trim());
								k++;
							}
							controlName = tempReplaceString;
						}

						// case 2
						if ((controlID.equalsIgnoreCase("XPathValueMulti")
								|| controlID.equalsIgnoreCase("XPathValueMulti_p"))
								&& !ctrlValue.equalsIgnoreCase("")) {
							String tempcontrolValue = ctrlValue;
							String[] tempValues = tempcontrolValue.split(";;");
							String tempReplaceString3 = controlName;
							int i = 1;
							for (String temp : tempValues) {
								if (temp.trim().equals("pickFromUniqueNumbers")) {
									temp = WebHelperUtil.ReadFromExcel(temp.trim(), WebHelper.columnName);
								}
								tempReplaceString3 = tempReplaceString3.replace("$" + i + "value", temp.trim());
								i++;
							}
							controlName = tempReplaceString3;
						}

						// case 3
						if ((controlID.equalsIgnoreCase("XPathValue_p") || controlID.equalsIgnoreCase("XPathValue")
								|| controlID.equalsIgnoreCase("Xpath_ctrlvalue")) && !ctrlValue.equalsIgnoreCase("")) {

							String temp = "", tempReplaceString3 = "";
							if (ctrlValue.trim().equals("pickFromUniqueNumbers")) {
								temp = WebHelperUtil.ReadFromExcel(ctrlValue.trim(), WebHelper.columnName);
							} else
								temp = ctrlValue;

							tempReplaceString3 = controlName.replace("$value", temp);
							controlName = tempReplaceString3;
						}

						if (!((action.equals("I") || action.equals("V_text")) && ctrlValue.equalsIgnoreCase(""))) {
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
					// System.out.println("ExecuteFlag is N");
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

			// Retrieving PDF comparison prerequisite/error status
			if (StringUtils.equalsIgnoreCase(controltype, "PDFDocumentCompare")) {
				webDriver.getReport().setMessage(PDFComparisonUtil.webDriverPDF.getReport().getMessage());
				webDriver.getReport().setStatus(PDFComparisonUtil.webDriverPDF.getReport().getStatus());
			}
			// Retrieving XML comparison prerequisite/error status
			if (StringUtils.equalsIgnoreCase(controltype, "XMLCompare")) {
				webDriver.getReport().setMessage(XmlComparisonUtil.webDriverXML.getReport().getMessage());
				webDriver.getReport().setStatus(XmlComparisonUtil.webDriverXML.getReport().getStatus());
			}

		} catch (Exception e) {
			// log.error(e.getMessage(), e);
			//String errorMessage = e.getMessage();
			//String firstLine = errorMessage.split("\\r?\\n")[0];
			
			String errorMessage = (e.getMessage() != null && !e.getMessage().isEmpty()) ? e.getMessage(): "No error message available";
			String firstLine = errorMessage.split("\\r?\\n", 2)[0];
			log.error("Error :" + firstLine);
			// log.error(e.getMessage(), e);
			controller.pauseFun("Error : LogicalName: " + logicalName + ".\n" + e.getMessage());
			// controller.pauseFun("FieldName Or LogicalName: " + logicalName + ".\n" +
			// e.getMessage());
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
		Reporter report = new Reporter();
		if (ITAFWebDriver.isSuiteApplication()) {
			report = WebHelperUtil.WriteToDetailResultsSuite(expectedValue, actualValue, columnName);
		} else {
			// below condition changed to take the value from config sheet
			// if (WebHelper.file.exists() == true && DResult == 1) {
			if (WebHelper.file.exists() == true
					&& StringUtils.equalsIgnoreCase(Config.appendVerificationResultPath, "false") && DResult == 1) {
				// print = new PrintStream(file);
				WebHelper.file.delete();
			}
			// File file= new
			// File(Automation.getConfigValue("VERIFICATIONRESULTSPATH").toString());
			report.setReport(report);
			report = report.getReport();
			String passCount = "";
			String failCount = "";
			report.setGroupName(controller.controllerGroupName.toString());// Sheetal - 26/07/2019
			report.setTestcaseId(controller.controllerTestCaseID.toString());
			report.setTrasactionType(controller.controllerTransactionType.toString());
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
				// System.out.println("ab");
				report.setExpectedValue(expectedValue);
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
						"GroupName,TestCaseID,TransactionType,CurrentDate,RowType,Status,PassCount,FailCount, Compare Result,Quote Number");
				WebHelper.print.println();
			}
			usedRows = WebHelperUtil.count(WebHelper.file);

			WebHelper.print.print(ExcelUtility.myChar + report.getGroupName() + ExcelUtility.myChar + ","
					+ ExcelUtility.myChar + report.getTestcaseId() + ExcelUtility.myChar + "," + ExcelUtility.myChar
					+ report.getTrasactionType() + ExcelUtility.myChar + "," + ExcelUtility.myChar + report.getToDate()
					+ ExcelUtility.myChar + "," + ExcelUtility.myChar + "Field: " + columnName + ExcelUtility.myChar
					+ "," + ExcelUtility.myChar + report.getStatus() + ExcelUtility.myChar + "," + ExcelUtility.myChar
					+ passCount + ExcelUtility.myChar + "," + ExcelUtility.myChar + failCount + ExcelUtility.myChar
					+ "," + ExcelUtility.myChar + report.getActualValue() + ExcelUtility.myChar + ","
					+ ExcelUtility.myChar + report.getStrQuoteId() + ExcelUtility.myChar);
			WebHelper.print.println();
			DResult++;
		}
		return report;
	}

	/** Locating Web Element **/
	@SuppressWarnings("incomplete-switch")
	public static WebElement getElementByType(String controlId, String controlName, String controlType,
			String imageType, String controlValue) throws Exception {
		// String a=controlValue;

		// WebDriverWait wait = new WebDriverWait(Automation.driver,
		// Integer.parseInt(Config.timeOut));//Sel4
		WebDriverWait wait = new WebDriverWait(Automation.driver, Duration.ofSeconds(Integer.parseInt(Config.timeOut)));

		if (Config.projectName.equals("DistributionManagement")) // Minaakshi :
		// 05-02-2019
		{
			WebHelperPAS.implementWait();
		}
		WebElement controlList = null;
		Constants.ControlIdEnum controlID = Constants.ControlIdEnum.valueOf(controlId);
		WebDriverWait wait2 = new WebDriverWait(Automation.driver, Duration.ofSeconds(2));// Sel4

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
				// Thread.sleep(1000);
				controlList = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
				break;
			case XPathValue: // Minaakshi : 10-09-2018
				String tempCtrlName = controlName;
				String tempReplaceString = tempCtrlName.replace("$value", controlValue);
				// controlList =
				// wait.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString)));
				controlList = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString)));
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
					if (!controlName.equals("")) {
						String tempCtrlName1 = controlName;
						String tempReplaceString1 = tempCtrlName1.replace("$value", controlValue);
						controlList = wait2
								.until(ExpectedConditions.presenceOfElementLocated(By.xpath(tempReplaceString1)));
					} else {
						controlList = null;
					}
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
					if (temp.trim().equals("pickFromUniqueNumbers")) {
						temp = WebHelperUtil.ReadFromExcel(temp.trim(), WebHelper.columnName);
					}
					tempReplaceString3 = tempReplaceString3.replace("$" + i + "value", temp.trim());
					i++;
				}

				// swapnil41444: added try catch block to controlList
				try {
					controlList = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString3)));
				} catch (Exception e) {
					log.info("Element is not visible/clickable : " + tempReplaceString3);
					controlList = null;
				}

				break;

			case XPathValueMultiWithInput:
				String tempcontrolValue2 = controlValue;
				String[] tempValues2 = tempcontrolValue2.split(";;");
				String tempReplaceString4 = controlName;
				int k = 1;
				/*
				 * for (int j = 0; j < tempValues2.length - 1; j++) { tempReplaceString4 =
				 * tempReplaceString4.replace("$" + k + "value", tempValues2[j].trim()); k++; }
				 */
				for (String temp : tempValues2) {
					if (temp.trim().equals("pickFromUniqueNumbers")) {
						temp = WebHelperUtil.ReadFromExcel(temp.trim(), WebHelper.columnName);
					}
					tempReplaceString4 = tempReplaceString4.replace("$" + k + "value", temp.trim());
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
					for (String temp : tempValues3) {
						if (temp.trim().equals("pickFromUniqueNumbers")) {
							temp = WebHelperUtil.ReadFromExcel(temp.trim(), WebHelper.columnName);
						}
						tempReplaceString6 = tempReplaceString6.replace("$" + s + "value", temp.trim());
						s++;
					}
					controlList = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString6)));
					/*
					 * for (int j = 0; j < tempValues3.length - 1; j++) {
					 * if(tempValues3[j].trim().equals("pickFromUniqueNumbers")) { tempValues3 =
					 * WebHelperUtil.ReadFromExcel(tempValues3[j].trim(), WebHelper.columnName); }
					 * tempReplaceString6 = tempReplaceString6.replace("$" + s + "value",
					 * tempValues3[j].trim()); s++; } controlList =
					 * wait2.until(ExpectedConditions.elementToBeClickable(By.xpath(
					 * tempReplaceString6)));
					 */
				} catch (Exception e) {
					controlList = null;
				}
				break;

			// Designed considering CA-Covered Auto Symbol screen inputs
			case XPathValueEval:
				try {
					String tempcontrolValue1 = controlValue;
					String[] tempValues1 = tempcontrolValue1.split(";;");

					String interControlName = controlName;
					String finalControlName = "";
					String interImageType = imageType;
					// Replace $*value
					int ctr = 1;
					for (String temp1 : tempValues1) {
						interControlName = interControlName.replace("$" + ctr + "value", temp1.trim());
						interImageType = interImageType.replace("$" + ctr + "value", temp1.trim());
						ctr++;
					}

					// Evaluate
					String tempXpath = interImageType + "//preceding-sibling::td";
					List<WebElement> preElements = Automation.driver.findElements(By.xpath(tempXpath));
					if (!preElements.isEmpty()) {
						String replaceEval = Integer.toString(preElements.size() + 1);
						finalControlName = interControlName.replace("$eval", replaceEval);

						controlList = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(finalControlName)));
					} else {
						throw new Exception("Element(s) with XPath " + tempXpath + " not found");
					}
				} catch (Exception e) {
					log.error("Issue in XPathValueEval. Cross check XPaths given in ControlName and ImageType.\n"
							+ e.getMessage());
					throw new Exception(
							"Issue in XPathValueEval. Cross check XPaths given in ControlName and ImageType.\n"
									+ e.getMessage());
				}
				break;

			case XPathValueMulti_p:
				String tempcontrolVal = controlValue;
				String[] tempVal = tempcontrolVal.split(";;");
				String tempReplaceStr = controlName;
				int l = 1;
				for (String temp : tempVal) {
					if (temp.trim().equals("pickFromUniqueNumbers")) {
						temp = WebHelperUtil.ReadFromExcel(temp.trim(), WebHelper.columnName);
					}
					tempReplaceStr = tempReplaceStr.replace("$" + l + "value", temp.trim());
					l++;
				}
				try {
					controlList = wait2.until(ExpectedConditions.presenceOfElementLocated(By.xpath(tempReplaceStr)));
				} catch (Exception e) {

					controlList = null;
				}
				break;

			case XPath_dy:// Minaakshi : Added this case while handling
				// dyanamic
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
				WebDriverWait wait3 = new WebDriverWait(Automation.driver, Duration.ofSeconds(5));// Sel4
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
						// controlList =
						// waitForElementPresence.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
					}
				} else {
					// controlList =
					// Automation.driver.findElement(By.xpath(controlName+"[contains(text(),'"+controlValue+"')]"));
					controlList = wait.until(ExpectedConditions
							.elementToBeClickable(By.xpath(controlName + "[contains(text(),'" + controlValue + "')]")));
					break;
				}
				// Automation.driver.findElement(By.xpath(controlName+"[contains(text(),'"+controlValue+"')]"));
				// controlList = wait
				// .until(ExpectedConditions.elementToBeClickable(By.xpath(controlName
				// + "[contains(text(),'" + controlValue + "')]")));
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
				// controlList =
				// Automation.driver.findElement(By.xpath(controlName+"[contains(text(),'"+controlValue+"')]"));
				// controlList = wait
				// .until(ExpectedConditions.elementToBeClickable(By.xpath(controlName
				// + "[contains(text(),'" + UNSheetValue2 + "')]")));
				try {// Minaakshi : 05-02-2019
					controlList = dmgebtWait.until(ExpectedConditions.elementToBeClickable(
							By.xpath(controlName + "[contains(text(),'" + UNSheetValue2 + "')]")));

				} catch (Exception ex) {
					log.error(ex.getMessage(), ex);
					log.info("In Second Wait : XPATH : Element is not visible : " + controlName);
					controlList = waitForElementPresence.until(ExpectedConditions.visibilityOfElementLocated(
							By.xpath(controlName + "[contains(text(),'" + UNSheetValue2 + "')]")));

					// controlList =
					// waitForElementPresence.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
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
			//String errorMessage = e.getMessage();
			//String firstLine = errorMessage.split("\\r?\\n")[0];
			String errorMessage = (e.getMessage() != null && !e.getMessage().isEmpty()) ? e.getMessage(): "No error message available";
			String firstLine = errorMessage.split("\\r?\\n", 2)[0];
			log.error("Error :" + firstLine);
			// log.error(e.getMessage(), e);
			// log.error(e.getMessage(), e);
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
		// System.out.println(sikscreen);
		// bhaskar
		if (controlType.contains("Robot") && !WebHelper.isIntialized) {
			System.out.println("In method doaction debug1");
			WebHelper.robot = new Robot();
			WebHelper.isIntialized = true;
		}

		
		if(!(ctrlValue == "")&& !(ctrlValue == null)){
			log.info("In method doAction, ctrlValue is : " + ctrlValue);	
		}
		// Mrinmayee - skip the step for which control value is blank : start

		JavascriptExecutor js = (JavascriptExecutor) Automation.driver;
		try {
			// getHTMLResponse();
			// System.out.println("In method doaction debug3");
			switch (controlTypeEnum) {
			case Screenshot:
				switch (actionName) {
				case NC:
					if (webElement != null) {
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								webElement);
						Thread.sleep(100);
					}
					if (ITAFWebDriver.isSuiteApplication())
						WebHelperUtil.saveScreenShot("Screenshot");
					else
						saveScreenShot(controlTypeEnum.toString());
					break;
				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					if (webElement != null) {
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								webElement);
						Thread.sleep(100);
					}
					if (ITAFWebDriver.isSuiteApplication())
						WebHelperUtil.saveScreenShot("Screenshot");
					else
						saveScreenShot(controlTypeEnum.toString());

					break;
				}
				break;

			case RunJBeamBatch:
				switch (actionName) {
				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					// RunBatch.RunBatch_WS();//PDC count considered
					RunBatch.RunBatch_WSFF(); // PDC count not considered
					if (webDriver.getReport().getStatus().equalsIgnoreCase("FAIL")
							|| webDriver.getReport().getStatus().equalsIgnoreCase("BATCH USER INTERRUPTED")) {
						throw new Exception("Error in JBEAM Batch run");
					}

					break;
				}
				break;

			case WebEdit:
				switch (actionName) {
				case NC:

					webElement = getElementByType(controlId, controlName, WebHelper.control, imageType, ctrlValue);

					try {
						if (webElement.isEnabled()) {
							Thread.sleep(500);
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
							Thread.sleep(1000);
							if (webElement.getAttribute("value").isEmpty()) {
								webElement.sendKeys(ctrlValue);

							}
							// Automation.driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
							Automation.driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));// Sel4
							// Automation.driver.manage().timeouts().pageLoadTimeout(50,
							// TimeUnit.SECONDS);//Sel4
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
					Thread.sleep(1000);
					webElement.sendKeys(uniqueNumber);
					// Change requested by Dhiraj
					Thread.sleep(500);
					webElement.sendKeys(Keys.TAB);
					break;

				case Read2:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					uniqueNumber = DMProduct.ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
					webElement.clear();
					Thread.sleep(1000);
					webElement.sendKeys(uniqueNumber);
					Thread.sleep(500);
					webElement.sendKeys(Keys.TAB);
					break;

				case Override: // Tanuja 05272020
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					uniqueNumber = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
					webElement.sendKeys(Keys.chord(Keys.CONTROL, "a"), uniqueNumber);
					break;

				case Write:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo,
							colNo);
					break;

				case WriteAttribute:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					String AttributeValue = "";
					AttributeValue = webElement.getAttribute("value");
					if (AttributeValue == null || AttributeValue.equalsIgnoreCase("")) {
						AttributeValue = webElement.getText();
						if (AttributeValue == null) {
							AttributeValue = "";
						}
					}
					WebHelperUtil.writeToExcel(AttributeValue, null, "", "writeCtrlValueAsIs", "", "", "");

					break;

				case ProcessAndWrite:
					if (ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					String newValue = "";
					newValue = webElement.getAttribute("value");
					if (newValue == null || newValue.equalsIgnoreCase("")) {
						newValue = webElement.getText();
						if (newValue == null) {
							newValue = "";
						}
					}
					if (!newValue.equals("")) {
						DateTimeFormatter dtFormatWithHyphen = DateTimeFormatter.ofPattern("yyyy-MM-dd");
						DateTimeFormatter dtFormatWithSlash = DateTimeFormatter.ofPattern("MM/dd/yyyy");
						LocalDate cycledate = LocalDate.parse(newValue, dtFormatWithSlash);
						newValue = dtFormatWithHyphen.format(cycledate);

					}
					WebHelperUtil.writeToExcel(newValue, null, "", "writeCtrlValueAsIs", "", "", "");

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
					else {
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
					}
					break;
				// js1.executeScript("arguments[0].value = '';",
				// webElement);

				case I:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}

					if (ctrlValue != null) {
						// waitForPageLoadingToComplete();
						// waitForAjaxLoad(Automation.driver);
						try {
							// sheetal 2-7-2019
							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
									webElement);
							Thread.sleep(100);

							if (webElement.isEnabled()) {
								if (ctrlValue.equalsIgnoreCase("HitTabKey")) {
									webElement.sendKeys(Keys.TAB);
								} else if (ctrlValue.equalsIgnoreCase("REMOVEVALUE")) {
									webElement.clear();
								} else {
									WebHelperUtil.InputValue(webElement, ctrlValue, controlId, controlName, imageType,
											ctrlValue);
								}
							}

						} catch (Exception e) {
							//log.error(e.getMessage(), e);
							log.info("Element not displayed");
							//String errorMessage = e.getMessage();//Print only first line - 
	        				//String firstLine = errorMessage.split("\\r?\\n")[0];
	        				String errorMessage = (e.getMessage() != null && !e.getMessage().isEmpty()) ? e.getMessage(): "No error message available";
	        				String firstLine = errorMessage.split("\\r?\\n", 2)[0];
	                        log.error("Element not displayed."+ " - " + firstLine);//---
							
						}
					}

					break;// Minaakshi : 17-12-2018

				case InputDate:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}
					try {

						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								webElement);
						if (ctrlValue.contains("=") || ctrlValue.contains("-") || ctrlValue.contains("+")) {

							String effDate = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
							DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/uuuu");
							LocalDate d = LocalDate.parse(effDate, formatter);
							LocalDate newEffDate;

							if (ctrlValue.contains("=")) {
								ctrlValue = effDate;
							} else if (ctrlValue.contains("-")) {
								newEffDate = d.minusDays(Integer.parseInt(ctrlValue.replace("-", "")));
								ctrlValue = newEffDate.format(formatter);
							} else if ((ctrlValue.contains("+"))) {
								newEffDate = d.plusDays(Integer.parseInt(ctrlValue.replace("+", "")));
								ctrlValue = newEffDate.format(formatter);
							}
						}

						if (webElement.isEnabled()) {
							// WebHelperUtil.InputValue(webElement, ctrlValue, controlId, controlName,
							// imageType, ctrlValue);
							WebHelperPAS.doAction(null, "ReplaceDefault", null, null, ctrlValue, null, "I", webElement,
									Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo);
						}

					} catch (Exception e) {
						log.error(e.getMessage(), e);
					}
					break;

				case InputBusinessDate:
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}

					ctrlValue = WebHelper.doAction(null, null, null, null, "DeriveBusinessDate", null, null, ctrlValue,
							null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
							rowcount, rowNo, colNo, null, null, null);

					if (webElement.isEnabled()) {
						WebHelperPAS.doAction(null, "ReplaceDefault", null, null, ctrlValue, null, "I", webElement,
								Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo);
					}

					break;
					
				case Input_popup: // Same as Input case, however removed code for TAB, so that Pop-up shouldn't go away
					if (ctrlValue == null || ctrlValue.trim().equals("")) {
						break;
					}

					if (webElement.isEnabled() && !ctrlValue.equalsIgnoreCase("null")) {
							for (int i = 0; i < 5; i++) {
								try {
									webElement.sendKeys(Keys.ENTER);
									webElement.clear();
									Thread.sleep(500);
									webElement.sendKeys(ctrlValue);
									System.out.println(webElement.getAttribute("value"));
									break;
								} catch (Exception e) {
									System.out.println("Element not exist in WebEdit - Input_popup case");
									log.error(e.getMessage(), e);
									i++;
								}
							}
					}
					break;
	
				}
				break;

			case TextboxToTextbox:
				if (!ctrlValue.isEmpty()) {
					String[] combiValue = ctrlValue.split("\\|");
					String refValue = combiValue[0];
					String actualValue = combiValue[1];
					try {
						if (Automation.driver.findElement(By.xpath(controlName)).isDisplayed()) {
							List<WebElement> refEle = Automation.driver.findElements(By.xpath(controlName));
							for (WebElement element : refEle) {
								if (element.getAttribute("value").equalsIgnoreCase(refValue)) {
									String Id = element.getAttribute("id");
									String elementTagName = element.getTagName();
									String controlNameNew = "//" + elementTagName + "[@id='" + Id + "'" + "]/"
											+ imageType;
									WebElement refTextBox = Automation.driver.findElement(By.xpath(controlNameNew));
									Thread.sleep(500);
									((JavascriptExecutor) Automation.driver)
											.executeScript("arguments[0].scrollIntoView();", refTextBox);
									Thread.sleep(500);
									if (actualValue.equalsIgnoreCase("REMOVEVALUE")) {
										refTextBox.clear();
									} else {
										WebHelperUtil.InputValue(refTextBox, actualValue, controlId, controlName,
												imageType, ctrlValue);
									}
									refTextBox = null;
									controlNameNew = "";
									break;
								}
							}
						} else {
							System.out.println("Element not displayed on screen");
						}
					} catch (Exception rwe) {
						System.out.println("Element Not Found on Page");
					}
				} else {
					System.out.print("Cannot perform any action!!!");
				}
				break;

			// Mrinmayee 15-10-2019 Code to enter the value in textbox using enabled or
			// disabled input values from the table
			case InputUsingRefElements:
				switch (actionName) {
				case I:
					if (!ctrlValue.isEmpty()) {
						String[] combiValue = ctrlValue.trim().split(";;");
						String inputVal = combiValue[combiValue.length - 1];
						try {
							if (Automation.driver.findElement(By.xpath(controlName)).isDisplayed()) {
								List<WebElement> refEle = Automation.driver.findElements(By.xpath(controlName));
								int ctr = 1;

								for (@SuppressWarnings("unused")
								WebElement element : refEle) {
									String[] found = new String[combiValue.length - 1];
									int foundCtr = 0;

									for (int j = 0; j < combiValue.length - 1; j++) {
										String[] tempTag = controlId.trim().split(";;");
										String tag = tempTag[j];
										String refValue = combiValue[j];
										String tempXpath = controlName + "[" + ctr + "]//" + tag;
										List<WebElement> refEleInside = Automation.driver
												.findElements(By.xpath(tempXpath));
										for (WebElement elementInside : refEleInside) {
											try {

												String val = elementInside.getAttribute("value").trim();
												if (StringUtils.equalsIgnoreCase(val, null)
														|| val.equalsIgnoreCase("")) {
													val = elementInside.getText().trim();
													if (StringUtils.equalsIgnoreCase(val, null)) {
														val = "";
													}

												}
												if (val.equalsIgnoreCase(refValue)) {
													foundCtr++;
													break;
												}

											} catch (Exception e) {
											}

										}

										if (foundCtr >= 1)
											found[j] = "found";
										else
											found[j] = "not";

										foundCtr = 0;
									}
									boolean foundCtrFinal = true;
									for (int k = 0; k < found.length; k++) {
										if (found[k].equalsIgnoreCase("not")) {
											foundCtrFinal = false;
											break;
										}
									}
									if (foundCtrFinal) {
										String controlNameNew = controlName + "[" + ctr + "]/" + imageType;
										WebElement refTextBox = Automation.driver.findElement(By.xpath(controlNameNew));
										Thread.sleep(200);
										((JavascriptExecutor) Automation.driver)
												.executeScript("arguments[0].scrollIntoView();", refTextBox);
										Thread.sleep(200);
										WebHelperUtil.InputValue(refTextBox, inputVal, controlId, controlName,
												imageType, ctrlValue);
										break;
									}
									ctr++;
								}
							} else {
								System.out.println("Element not displayed on screen");
							}
						} catch (Exception rwe) {
							System.out.println("Element Not Found on Page");
						}
					} else {
						System.out.print("Cannot perform any action!!!");
					}
					break;
				}
				break;

			case ReplaceDefault:
				switch (actionName) {
				case I:
					if (ctrlValue.equalsIgnoreCase("null") || ctrlValue.equalsIgnoreCase("")) {
						break;
					}
					if (!StringUtils.equalsIgnoreCase(Config.applicationPlatform, "digital1st")) {
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView(true);",
								webElement);
						Thread.sleep(500);
					}
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
								System.out.println("Could not set correct value");
								break;
							}
						}

					} catch (StaleElementReferenceException e) {
						e.toString();
						System.out.println("Trying to recover from a stale element :" + e.getMessage());
					}
					break;
				}
				break;

			case WebEdit2:
				switch (actionName) {
				case I:
					if (ctrlValue.equalsIgnoreCase("null") || ctrlValue.equalsIgnoreCase("")) {
						break;
					}
					if (!StringUtils.equalsIgnoreCase(Config.applicationPlatform, "digital1st")) {
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView(true);",
								webElement);
						Thread.sleep(500);
					}
					webElement.sendKeys(Keys.HOME, Keys.chord(Keys.SHIFT, Keys.END), ctrlValue);
					Thread.sleep(250);
					webElement.sendKeys(Keys.TAB);
					Thread.sleep(250);
					break;
				}
				break;

			case WebEdit3:
				switch (actionName) {
				case I:
					if (ctrlValue.equalsIgnoreCase("null") || ctrlValue.equalsIgnoreCase("")) {
						break;
					}
					if (!StringUtils.equalsIgnoreCase(Config.applicationPlatform, "digital1st")) {
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView(true);",
								webElement);
						Thread.sleep(500);
					}
					webElement.sendKeys(Keys.HOME, Keys.chord(Keys.SHIFT, Keys.END), ctrlValue);
					Thread.sleep(250);
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

			case HandleWSFailure:
				if (!ctrlValue.equalsIgnoreCase("")) {
					try {

						@SuppressWarnings("unused")
						WebElement errorPopUp = WebHelper.wait.until(ExpectedConditions.presenceOfElementLocated(By
								.xpath("//div[contains(text(),'Error in rating : No Loss Data returned. There was a failure in the Web Service to retrieve Loss Data. Please select reorder losses on experience rating tab when the Web Service is enabled.')]")));
						WebElement OKButton = WebHelper.wait
								.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[text()='OK']")));

						Actions Click = new Actions(Automation.driver);
						highlightElement(OKButton);
						Action Mouseclick = Click.moveToElement(OKButton).clickAndHold().release().build();
						Mouseclick.perform();
						WebHelperPAS.doAction(null, "Wait", null, "10", null, null, "NC", webElement, Results,
								strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo);

						switch (ctrlValue) {
						case "CP":
							Automation.driver.findElement(By.xpath("//div[text()='Commercial Package']")).click();
							break;
						case "BP":
							Automation.driver.findElement(By.xpath("//div[text()='Business Protector Policy']"))
									.click();
							break;
						case "TC":
							Automation.driver.findElement(By.xpath("//div[text()='Target Contractor Protector']"))
									.click();
							break;
						case "CA":
							Automation.driver.findElement(By.xpath("//div[text()='Commercial Auto']")).click();
							break;
						}
						WebHelperPAS.doAction(null, "Wait", null, "10", null, null, "NC", webElement, Results,
								strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo);
						if (ctrlValue.equalsIgnoreCase("CA")) {
							// Thread.sleep(10000);
							Automation.driver.findElement(By.xpath("//div[text()='Experience Rating']")).click();
							WebHelperPAS.doAction(null, "Wait", null, "10", null, null, "NC", webElement, Results,
									strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo);
							Automation.driver.findElement(By.xpath("//div[text()='Liability Experience Rating']"))
									.click();
							WebHelperPAS.doAction(null, "Wait", null, "10", null, null, "NC", webElement, Results,
									strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo);
						} else {
							// Automation.driver.findElement(By.xpath("//div[text()='Commercial
							// Package']")).click();
							// Thread.sleep(10000);
							Automation.driver.findElement(By.xpath("(//div[text()='General Liability'])[1]")).click();
							WebHelperPAS.doAction(null, "Wait", null, "10", null, null, "NC", webElement, Results,
									strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo);
							Automation.driver.findElement(By.xpath("//div[text()='Experience Rating']")).click();
							WebHelperPAS.doAction(null, "Wait", null, "10", null, null, "NC", webElement, Results,
									strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo);
						}

						WebElement reorderLosses = Automation.driver
								.findElement(By.xpath("//label[text()='Reorder Losses']"));
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								reorderLosses);
						reorderLosses.click();

						WebHelperPAS.doAction(null, "Wait", null, "20", null, null, "NC", webElement, Results,
								strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo);
						WebElement calcButton = Automation.driver
								.findElement(By.xpath("//span[text()='Calculate Premium']"));

						Actions builderClick = new Actions(Automation.driver);
						Action clickAction = builderClick.moveToElement(calcButton).clickAndHold().release().build();
						clickAction.perform();

						// Thread.sleep(Integer.parseInt(tempCtrlValue[1])*1000);
						WebHelperPAS.doAction(null, "Wait", null, "20", null, null, "NC", webElement, Results,
								strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo);
						System.out.println(
								"WebService failure occurred. \n Reorder losses checkbox selected on Experience Rating tab and Policy re-calculated.");
						log.info(
								"WebService failure occurred. \n Reorder losses checkbox selected on Experience Rating tab and Policy re-calculated.");
					} catch (Exception e) {
						try {
							if (ctrlValue.equalsIgnoreCase("CA")) {
								Automation.driver.switchTo().frame(Automation.driver.findElement(
										By.xpath("//iframe[contains(@src,'cover-all.com/mic/portal/pcteui')]")));
								@SuppressWarnings("unused")
								WebElement premSummary = WebHelper.wait
										.until(ExpectedConditions.presenceOfElementLocated(
												By.xpath("//td[text()='Commercial Auto Premium Summary']")));
								System.out.println("WebService call is successful.");
								log.info("WebService call is successful.");
							} else {
								Automation.driver.switchTo().frame(Automation.driver.findElement(
										By.xpath("//iframe[contains(@src,'cover-all.com/mic/portal/pcteui')]")));
								@SuppressWarnings("unused")
								WebElement premSummary = WebHelper.wait
										.until(ExpectedConditions.presenceOfElementLocated(By.xpath(
												"//td[text()='General Information']/following::td[text()='Premium Information']")));
								System.out.println("WebService call is successful.");
								log.info("WebService call is successful.");
							}

							Set<String> handlers = null;
							handlers = Automation.driver.getWindowHandles();
							for (String handler : handlers) {
								Automation.driver = Automation.driver.switchTo().window(handler);

								// TM-19/01/2015: Changed following comparison from equalsIgnoreCase to contains
								if (Automation.driver.getTitle().contains("Cover-All Policy")) {
									System.out.println("Focus on window with title: " + Automation.driver.getTitle());
									break;
								}
							}

						} catch (Exception e1) {
							System.out.println("No Webservice Error | Unsuccessful Rating | Kindly check the error.");
							log.info("No Webservice Error | Unsuccessful Rating | Kindly check the error.");
						}
					}
				}
				break;

			case WebButton:
				switch (actionName) {
				case FileUpload:
					// Added code to handle Fileupload window using
					// AutoIt
					String autoitFileDir = Config.inputDataFilePath + TransactionMapping.directoryPathFileUpload;

					// webElement.click(); // swapnil41444: This is not working. Hence using
					// JavascriptExecutor option as below
					// JavascriptExecutor executor = (JavascriptExecutor) Automation.driver;
					// executor.executeScript("arguments[0].click();", webElement);

					// 22/06/23 - Changing it to double click as both webElement.click() and
					// JavascriptExecutor are not working
					Actions builderClick1 = new Actions(Automation.driver);
					Action clickAction1 = builderClick1.moveToElement(webElement).clickAndHold().release().build();
					clickAction1.perform();

					Thread.sleep(5000);
					// Runtime.getRuntime().exec(autoitFileDir +
					// "\\FileUpload.exe "+ctrlValue);
					Runtime.getRuntime().exec(autoitFileDir + "\\FileUpload.exe " + autoitFileDir + "\\" + ctrlValue);
					Thread.sleep(5000);
					break;

				case FileUpload_DM:// Minaakshi : 01-03-2019

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

				//swapnil41444: Updated the WebButton - I case code.. 
				case I:
				    if (ctrlValue == null || ctrlValue.trim().isEmpty() || webElement == null) {
				        break;
				    }
				    if (!StringUtils.equalsIgnoreCase(ctrlValue, "No")) {

				        int attempts = 6;
				        while (attempts > 0) {
				            try {
				                System.out.println("WebButton Click Attempt: " + attempts + " with XPATH: " + controlName);

				                WebDriverWait wait = new WebDriverWait(Automation.driver, Duration.ofSeconds(attempts));
				                webElement = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));

				                if (webElement.isEnabled() || webElement.isSelected()) {
				                	((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView(true);", webElement);
				                	webElement.click();
									break;
				                } else {
				                    System.out.println("Element not enabled.");
				                    attempts--;
				                }

				                Thread.sleep(500);

				                boolean isSelected = false;
								List<WebElement> options;
								int options_size=-1;
				                try {
									isSelected = webElement.isSelected();
									options = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath(controlName)));
									options_size = options.size();
								} catch (Exception e) {}

				                if (isSelected || options_size > 0) {
				                    System.out.println("WebButton is selected through webElement.click.");
				                    break;
				                } else {
				                    System.out.println("WebButton is not selected.");
				                    Thread.sleep(1000);
				                    attempts--;
				                }
				            } catch (Exception e) {
				                System.out.println("Inside WebButton main catch. Try alternative option now.");
				                //log.error(e.getClass().getCanonicalName(), e);
				                
				                try {
				                    WebDriverWait wait = new WebDriverWait(Automation.driver, Duration.ofSeconds(attempts));
				                    webElement = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
				                    Thread.sleep(500);

									boolean isSelected = false;

				                    if (webElement.isEnabled() || webElement.isSelected()) {
				                        if (attempts > 1) {
				                        	((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView(true);", webElement);
						                	((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();", webElement);
				                            Thread.sleep(500);

											List<WebElement> options;
											int options_size=-1;
											try {
												isSelected = true;
												options = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath(controlName)));
												options_size = options.size();
											} catch (Exception ee) {}

				                            if (isSelected || options_size > 0) {
				                                System.out.println("WebButton is selected through JavascriptExecutor.");
				                                break;
				                            } else {
				                                System.out.println("WebButton is not selected.");
				                                attempts--;
				                            }
				                        } else {
				                            System.out.println("Last attempt using Sendkeys.");
				                            webElement.sendKeys(Keys.ENTER);
											attempts--;
				                            break;
				                        }
				                    }
				                    Thread.sleep(500);
				                } catch (Exception e1) {
				                    System.out.println("Last catch attempt to click WebButton using MouseclickAction.");
				                    try {
				                        Actions actions = new Actions(Automation.driver);
				                        Action mouseClickAction = actions.moveToElement(webElement).clickAndHold().release().build();
				                        mouseClickAction.perform();
				                        attempts--;
				                    } catch (Exception lastE) {
				                        System.out.println("Element does not exist.");
				                        //log.error(lastE.getMessage(), lastE);// ---
				                        //String errorMessage = (lastE.getMessage() != null) ? lastE.getMessage() : "No error message available";		                        
				                        //String errorMessage = lastE.getMessage();
				        				String errorMessage = (lastE.getMessage() != null && !lastE.getMessage().isEmpty()) ? lastE.getMessage(): "No error message available";
				        				String firstLine = errorMessage.split("\\r?\\n", 2)[0];
				        				//String firstLine = errorMessage.split("\\r?\\n")[0];
				                        log.error("Element does not exist."+ " - " + firstLine);//---
				                        //Thread.sleep(1000);
				                        attempts--;
				                        break;
				                    }
				                }
				            }
				        }
				    } else {
				        System.out.println("Don't perform any action.");
				    }
				    break;

					
//---------------					
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

							} else {
								iCtr--;
								System.out.println("Element is not Enabled.");

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
								// webElement =
								// getElementByType(controlId,
								// controlName,control,imageType,ctrlValue);
								if (iCtr <= 2) {
									//log.error(e1.getMessage(), e1);//Mandar
									//String errorMessage = e1.getMessage();
									//String firstLine = errorMessage.split("\\r?\\n")[0];
									String errorMessage = (e1.getMessage() != null && !e1.getMessage().isEmpty()) ? e1.getMessage(): "No error message available";
			        				String firstLine = errorMessage.split("\\r?\\n", 2)[0];
									log.error(firstLine);
									
								} else {
									//log.error(e1.getMessage());//	
									//String errorMessage = e1.getMessage();
									//String firstLine = errorMessage.split("\\r?\\n")[0];
									String errorMessage = (e1.getMessage() != null && !e1.getMessage().isEmpty()) ? e1.getMessage(): "No error message available";
			        				String firstLine = errorMessage.split("\\r?\\n", 2)[0];
									log.error(firstLine);
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

				//Mandar - For Suite Integration -12/10/2024
					
				case NCIF_Button:// Sel4-26/07/24

					List<String> browserList1 = Arrays
							.asList(new String[] { "InternetExplorer", "Chrome", "MsEdge", "Firefox", "Safari" });
					// for Button
					try {
						if (Automation.driver.findElement(By.xpath(controlName)).isDisplayed()) {
							List<WebElement> ls = Automation.driver.findElements(By.xpath(controlName));
							System.out.println(ls.size());
							if (ls.size() > 0) {
								if (browserList1.contains(Automation.browserType.toString())) {
									((JavascriptExecutor) Automation.driver).executeScript(
											"arguments[0].click();",
											Automation.driver.findElement(By.xpath(controlName)));
									log.info("controlName: " + controlName + "clicked");
								} else {
									((JavascriptExecutor) Automation.driver).executeScript(
											"arguments[0].click();",
											Automation.driver.findElement(By.xpath(controlName)));
									log.info("controlName: " + controlName + "clicked");
								}
							}
						} else {

							log.info("Element" + "controlName" + "does not exists");
							break;
						}
					} catch (Exception e) {
						// throw(e);//log.error(e.getMessage(), e);
						log.info("Element Not Found on Page in NCIF case");
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
							Thread.sleep(1000);

							// Sheetal: 2/19/2019, commenting below line
							// as it
							// should not be part of Webbutton->NCIF
						}

					} catch (Exception e) {
						// log.error(e.getMessage(), e);
						System.out.println("Element Not Found on Page.");
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
								// webElement =
								// getElementByType(controlId,
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

				case WriteEntityReference:
					String curURL = Automation.driver.getCurrentUrl();
					if (curURL.indexOf("entityReference=") != -1) {
						int startIndex = curURL.indexOf("entityReference=") + 16;
						int endIndex = curURL.indexOf("&", startIndex);
						try {
							ctrlValue = curURL.substring(startIndex, endIndex);
						} catch (Exception e) {
							log.info("Error while retrieving Entity Reference number from URL - " + e);
							ctrlValue = "";
							webDriver.getReport()
									.setMessage("Error while retrieving Entity Reference number from URL - " + e);
							webDriver.getReport().setStatus("FAIL");
						}
					} else {
						log.info("URL does not contain Entity Reference Number. Please navigate to correct URL.");
						ctrlValue = "";
						webDriver.getReport().setMessage(
								"URL does not contain Entity Reference Number. Please navigate to correct URL.");
						webDriver.getReport().setStatus("FAIL");
					}
					WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, "writeCtrlValueAsIs", controlName,
							rowNo, colNo);
					break;

				case Extract_Value:// Mrinmayee 23/05/2019
					if (webElement == null || ctrlValue == null) {
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
				case VD: // Minaakshi : Added for handling dynamic
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
						} else if (ctrlValue.contains("|")) {
							// Minaakshi : 04-01-2019
							String[] temp = ctrlValue.split("\\|");
							String tempCtrlValue = DMProduct.ReadFromExcelUsingColumnName(temp[0].toString(),
									WebHelper.columnName);
							ctrlValue = temp[1].toString().replace("*value", tempCtrlValue);
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
				case UTOV:
					ExcelUtility.fillValidationValueSheet();
					break;

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

					// highlightElement(webElement);
					try {
						currentValue = webElement.getAttribute("value");

						if (StringUtils.equalsIgnoreCase(currentValue, null) || currentValue.equalsIgnoreCase("")) {
							currentValue = webElement.getText();
							if (StringUtils.equalsIgnoreCase(currentValue, null)) {
								currentValue = "";
							}
						}

					} catch (NullPointerException e) {
						currentValue = "Element not found : " + controlName;
					}
					if (ctrlValue.equalsIgnoreCase("pickFromUniqueNumbers")) {
						// ctrlValue = ReadFromExcel(ctrlValue);
						ctrlValue = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
						if (ctrlValue.equalsIgnoreCase("")) {
							ctrlValue = "Value missing in UniqueNumber file";
						}
					} else if (ctrlValue.contains("pickFromUniqueNumbers") && ctrlValue.contains(";;")) {
						String[] combiValue = ctrlValue.split(";;");
						ctrlValue = DMProduct.ReadFromExcelUsingColumnName(combiValue[1].toString(),
								WebHelper.columnName);
						if (ctrlValue.equalsIgnoreCase("")) {
							ctrlValue = "Value missing in UniqueNumber file. Expected value : " + ctrlValue + ".";
						}
					} else if (ctrlValue.contains("pickFromUniqueNumbers")) {
						String uniqueValue = "";
						uniqueValue = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
						if (uniqueValue.equalsIgnoreCase(""))
							ctrlValue = "Value missing in UniqueNumber file. Expected value : " + ctrlValue + ".";
						else
							ctrlValue = ctrlValue.replace("pickFromUniqueNumbers", uniqueValue);
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

				case V_stop:
					if (ctrlValue == null || ctrlValue.trim().equals("") || ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					try {
						currentValue = webElement.getAttribute("value");
						if (StringUtils.equalsIgnoreCase(currentValue, null) || currentValue.equalsIgnoreCase("")) {
							currentValue = webElement.getText();
							if (StringUtils.equalsIgnoreCase(currentValue, null)) {
								currentValue = "";
							}
						}
					} catch (NullPointerException e) {
						currentValue = "Element not found : " + controlName;
					}
					if (ctrlValue.equalsIgnoreCase("pickFromUniqueNumbers")) {
						ctrlValue = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
						if (ctrlValue.equalsIgnoreCase("")) {
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
					if (!ctrlValue.equals(currentValue)) {
						throw new Exception("Hard Assert Failure. Expected Value : " + ctrlValue + " | Actual Value : "
								+ currentValue);
					}

					break;

				case V_date:
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					highlightElement(webElement);
					currentValue = webElement.getAttribute("value");

					if (StringUtils.equalsIgnoreCase(currentValue, null) || currentValue.equalsIgnoreCase("")) {
						currentValue = webElement.getText();
						if (StringUtils.equalsIgnoreCase(currentValue, null)) {
							currentValue = "";
						}
					}

					try {

						DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/uuuu");
						ArrayList<String> finalOutput = new ArrayList<String>();
						String[] evaludateDate = ctrlValue.split(" ");
						int k = 0;
						while (k < evaludateDate.length) {
							if (evaludateDate[k].equalsIgnoreCase("plus")
									|| evaludateDate[k].equalsIgnoreCase("minus")) {
								finalOutput.remove(finalOutput.size() - 1);
								String d = "";
								if (evaludateDate[k - 1].equalsIgnoreCase("pickFromUniqueNumbers")) {
									d = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
								} else
									d = evaludateDate[k - 1];

								LocalDate tempDate;
								switch (evaludateDate[k].toLowerCase()) {
								case "plus":
									tempDate = LocalDate.parse(d, formatter);
									finalOutput.add(tempDate.plusDays(Integer.parseInt(evaludateDate[k + 1]))
											.format(formatter));
									break;
								case "minus":
									tempDate = LocalDate.parse(d, formatter);
									finalOutput.add(tempDate.minusDays(Integer.parseInt(evaludateDate[k + 1]))
											.format(formatter));
									break;
								default:
									break;
								}
								k = k + 2;
							} else {
								if (evaludateDate[k].equalsIgnoreCase("pickFromUniqueNumbers")) {
									finalOutput.add(WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName));
								} else {
									finalOutput.add(evaludateDate[k]);
								}
								k++;
							}
						}

						Iterator<String> l = finalOutput.iterator();
						String expectedValue = "";
						while (l.hasNext()) {
							expectedValue = expectedValue + l.next() + " ";
						}

						ctrlValue = expectedValue.trim();
					} catch (Exception e) {
						log.info("Error while decoding expected date in case V_date.");
						log.info(e.getMessage());
					}
					break;

				case V_format:
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					String[] formatTemp = ctrlValue.split("\\|");
					if (formatTemp.length != 2) {
						currentValue = "Incorrect input. Expected format example: 'X|20-CPXXXXX-01', where first letter specifies dynamic value.";
						ctrlValue = "Input provided: " + ctrlValue;
						break;
					}
					String dynamic = formatTemp[0];
					ctrlValue = formatTemp[1];

					highlightElement(webElement);
					currentValue = webElement.getAttribute("value");
					if (StringUtils.equalsIgnoreCase(currentValue, null) || currentValue.equalsIgnoreCase("")) {
						currentValue = webElement.getText();
						if (StringUtils.equalsIgnoreCase(currentValue, null)) {
							currentValue = "";
						}
					}
					if (currentValue.length() == ctrlValue.length()) {
						Boolean matched = true;
						for (int j = 0; j < ctrlValue.length(); j++) {
							String curChar = Character.toString(currentValue.charAt(j));
							String expChar = Character.toString(ctrlValue.charAt(j));
							if (!expChar.equalsIgnoreCase(dynamic) && !expChar.equals(curChar)) {
								currentValue = "Format mismatch. Actual Value : " + currentValue;
								ctrlValue = " Expected Format : " + ctrlValue;
								matched = false;
								break;
							}
						}
						if (matched) {
							currentValue = "Format matched. Actual Value : " + currentValue + " | Expected Format : "
									+ ctrlValue;
							ctrlValue = currentValue;
						}
					} else {
						currentValue = "Character Count Mismatch | Actual Value : " + currentValue;
						ctrlValue = " Expected Format : " + ctrlValue;
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
								// ctrlValue= ctrlValue + " is present in attribute "+ attribute + " of given
								// tag. Actual attribute value is " + currentValue;
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

				case V_toggleCheckboxStatus:
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					try {
						highlightElement(webElement);
						String classValue = webElement.getAttribute("class").toLowerCase();
						if (StringUtils.contains(classValue, "checked")) {
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

				case V_toggleOptionStatus:
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					try {
						highlightElement(webElement);
						String classValue = webElement.getAttribute("class").toLowerCase();
						if (StringUtils.contains(classValue, "toggle-option-selected")) {
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

				case V_buttonOptionStatus:
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					try {
						highlightElement(webElement);
						String attrValue = webElement.getAttribute("aria-checked").toLowerCase();
						if (StringUtils.contains(attrValue, "true")) {
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

				case V_backgroundStatus:
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					try {
						String attrValue = webElement.getAttribute("style").toLowerCase();
						if (StringUtils.contains(attrValue, "background-color: rgba")) {
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

				case V_toggleOptionDisableStatus:
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					try {
						highlightElement(webElement);
						String classValue = webElement.getAttribute("class").toLowerCase();

						if (StringUtils.contains(classValue, "disable-mouse-cursor")) {
							currentValue = "Disabled";
						} else {
							currentValue = "Enabled";
						}
					} catch (NullPointerException e) {
						currentValue = "Element not found : " + controlName;
					}
					if (currentValue.equalsIgnoreCase(ctrlValue)) {
						currentValue = ctrlValue;
					}
					break;

				case V_toggleDisableStatus:
					if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
							|| ctrlValue.equalsIgnoreCase("IGNORE")) {
						break;
					}
					try {
						highlightElement(webElement);
						String classValue = webElement.getAttribute("class").toLowerCase();

						if (StringUtils.contains(classValue, "disabled")) {
							currentValue = "Disabled";
						} else {
							currentValue = "Enabled";
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
					try {
						highlightElement(webElement);
						String disableStatus = webElement.getAttribute("disabled");

						if (StringUtils.equalsIgnoreCase(disableStatus, "true")) {
							currentValue = "Disabled";
						} else {
							currentValue = "Enabled";
						}
					} catch (NullPointerException e) {
						currentValue = "Element not found : " + controlName;
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

				case V_availabilityStatus:
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

				case VFCount:
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
									for (int k = 0; k < duplicateForms.size(); k++) {
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

						// replace value from Unique number sheet if required

						if (ctrlValue.trim().equals("pickFromUniqueNumbers")) {
							ctrlValue = WebHelperUtil.ReadFromExcel(ctrlValue.trim(), WebHelper.columnName);
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
				case VFPresence:
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

				}
				break;

			case CompleteFormVariable:
				System.out.println("webElement is: " + controlName);
				List<WebElement> list = Automation.driver.findElements(By.xpath(controlName));

				System.out.println("incomplete form count is: " + list.size());

				for (@SuppressWarnings("unused")
				WebElement form : list) {

					// form.click();
					Thread.sleep(5000);
					Automation.driver
							.findElement(By.xpath("//label[contains(text(),'Incomplete')]//following::td[1]/*"))
							.click();
					Thread.sleep(10000);
					// WebDriverWait waitToload = new WebDriverWait(Automation.driver, 500);
					// WebHelper.wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@value='Ok']")));
					// WebHelper.wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@value='Ok']")));
					// WebHelper.wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@value='Ok']")));

					List<WebElement> eleList = Automation.driver.findElements(
							By.xpath("//*[contains(@id,'MFV_VAR_VALUE_1')]//preceding::*[@value='Tax_percentage']"));
					List<WebElement> manuScriptList = Automation.driver.findElements(
							By.xpath("//*[contains(@id,'MFV_VAR_VALUE_1')]//preceding::*[@value='Manuscript Text']"));
					List<WebElement> enterDateList = Automation.driver.findElements(
							By.xpath("//*[contains(@id,'MFV_VAR_VALUE_1')]//preceding::*[@value='Enter Date']"));
					List<WebElement> maxHourlyWageRateList = Automation.driver.findElements(By.xpath(
							"//*[contains(@id,'MFV_VAR_VALUE_1')]//preceding::*[@value='Maximum hourly wage rate']"));
					if (eleList.size() > 0) {
						Automation.driver.findElement(By.xpath("//*[contains(@id,'MFV_VAR_VALUE_1')]")).clear();
						Automation.driver.findElement(By.xpath("//*[contains(@id,'MFV_VAR_VALUE_1')]")).sendKeys("10");
					} else if (manuScriptList.size() > 0) {

						Automation.driver.switchTo()
								.frame(Automation.driver.findElement(By.xpath("//*[@id='wysiwygMFV_VAR_VALUE_1']")));
						// Automation.driver.findElement(By.xpath("(//*[contains(@id,'MFV_VAR_VALUE_1')])[11]")).clear();
						// Automation.driver.findElement(By.xpath("(//*[contains(@id,'MFV_VAR_VALUE_1')])[11]")).sendKeys("X");
						Automation.driver.findElement(By.xpath("/html/body[@contenteditable=\"true\"]")).sendKeys("X");
						Automation.driver.switchTo().parentFrame();

					} else if (enterDateList.size() > 0) {
						Automation.driver.findElement(By.xpath("//*[contains(@id,'MFV_VAR_VALUE_1')]"))
								.sendKeys("10/02/2019");
					} else if (maxHourlyWageRateList.size() > 0) {
						Automation.driver.findElement(By.xpath("//*[contains(@id,'MFV_VAR_VALUE_1')]")).sendKeys("10");
					} else {
						Automation.driver.findElement(By.xpath("//*[contains(@id,'MFV_VAR_VALUE_1')]")).clear();
						Automation.driver.findElement(By.xpath("//*[contains(@id,'MFV_VAR_VALUE_1')]")).sendKeys("X");
					}
					List<WebElement> eleList2 = Automation.driver
							.findElements(By.xpath("//*[contains(@id,'MFV_VAR_VALUE_2')]"));
					if (eleList2.size() > 0) {
						List<WebElement> eachEvent = Automation.driver.findElements(
								By.xpath("//*[contains(@id,'MFV_VAR_VALUE_2')]//preceding::*[@value='Each Event']"));
						List<WebElement> leasedWorker = Automation.driver.findElements(By.xpath(
								"//*[contains(@id,'MFV_VAR_VALUE_2')]//preceding::*[@value='%of total cost of the contract for the \"leased workers\"']"));
						if (eachEvent.size() == 0 && leasedWorker.size() == 0) {
							Automation.driver.findElement(By.xpath("//*[contains(@id,'MFV_VAR_VALUE_2')]")).clear();
							Automation.driver.findElement(By.xpath("//*[contains(@id,'MFV_VAR_VALUE_2')]"))
									.sendKeys("X");
						}
					}
					Thread.sleep(2000);
					Automation.driver.findElement(By.xpath("//input[@value='Ok']")).click();
					Thread.sleep(2000);
				}

				System.out.println("Number of incomplete forms are: " + list.size());
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

			// case IJSScript:
			// IJavascriptExecutor ijs = Automation.driver;

			// ***Added (WaitToLoad:) by Sheetal for PAS team - 26/07/2019
			case WaitToLoad:
				return WebHelper.doAction(null, null, null, null, controlType, controlId, controlName, ctrlValue, null,
						null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex, rowcount,
						rowNo, colNo, null, null, null);

			case WaitForPageToLoad:
				// swapnil41444 06/30/2021: Adding case block to make option more customizable
				switch (actionName) {
				case I:
					if (!ctrlValue.isEmpty()) {
						// Automation.driver.manage().timeouts().pageLoadTimeout(120,
						// TimeUnit.SECONDS);//Sel4
						Automation.driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(120));
						WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(500));// Sel4
						WaitForPageLoad.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
					}

					break;
				case NC:
					try {
						// Automation.driver.manage().timeouts().pageLoadTimeout(120,
						// TimeUnit.SECONDS);//Sel4
						Automation.driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(120));
						WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(100));// Sel4
						WaitForPageLoad.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
						break;
					} catch (Exception e) {
						System.out.println("WebElement not found: " + controlName);
						Thread.sleep(1000);
					}

				}
				break;

			case WaitForElementToVisible: // Mrinmayee - Wait until the expected value is present on screen Added by COE
											// -26/08/2019//
				// swapnil41444 06/30/2021: Added try-catch block
				if (ctrlValue.equals("") || ctrlValue == null)
					break;

				int x = 500;
				while (x > 0) {
					try {
						String temp1 = Automation.driver.findElement(By.xpath(controlName)).getText();
						log.info("Inside WaitForElementToVisible | Current value:-" + temp1 + " | Expected value:-"
								+ ctrlValue);

						if (temp1.equals(ctrlValue)) {
							break;
						} else {
							Thread.sleep(2000);
						}
						x--;
					} catch (Exception e) {
						System.out.println("WebElement not found: " + controlName);
						Thread.sleep(1000);
						x--;
					}
				}

				if (x == 0)
					System.out.println("WebElement not found: " + controlName);

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

			// Tasleem - NYPIUA changes
			case ExecuteExe:
				switch (actionName) {
				case I:
					if (!ctrlValue.isEmpty()) {
						Runtime.getRuntime().exec(ctrlValue);
					}
					break;
				}
				// end of ExecuteExe

			case WaitForObjectPresent:
				WebHelperUtil.WebObjectPresent(webElement);
				break;

			case Wait:
				if (StringUtils.equalsIgnoreCase(Config.applyStaticWait, "false")) {
					log.info("Wait not applied");
				} else {
					// Thread.sleep(Integer.parseInt(controlName) * 1000);
					Thread.sleep((long) (Double.parseDouble(controlName) * 1000));
				}
				break;

			case Wait_DM: // // Minaakshi : 05-02-2019
				// Thread.sleep(5000);

				/*
				 * Wait<WebDriver> dmWait = new
				 * FluentWait<WebDriver>(Automation.driver).withTimeout(30, TimeUnit.SECONDS)
				 * .pollingEvery(1, TimeUnit.SECONDS).ignoring(NoSuchElementException.class);
				 */// Sel4

				Wait<WebDriver> dmWait = new FluentWait<>(Automation.driver).withTimeout(Duration.ofSeconds(30))
						.pollingEvery(Duration.ofSeconds(1)).ignoring(NoSuchElementException.class);

				if (controlName.equalsIgnoreCase("$value")) {
					if (ctrlValue.equalsIgnoreCase("HIGH")) {
						Thread.sleep(20000);
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
					try {
						WebElement dmWaitWebElement = dmWait
								.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(controlName)));
						if (dmWaitWebElement.isEnabled()) {
							Thread.sleep(5000);
							log.info("In Wait_DM : Element is enabled and applied hard wait of 5 seconds");
						} else {
							break;
						}

					} catch (Exception e) {
						log.info(
								"In Wait_DM: Catch : Element is not enabled and so not applied hard wait of 5 seconds");
						break;
					}
				}
				break;

			case Wait_IfValue:
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

			case SPNR_Wait: // DW - 16/Sept/2019: case created to call spinnerWait event.
				switch (actionName) {
				case I:
					if (!ctrlValue.isEmpty()) {
						try {
							log.info("Inside spinner wait");
							WebElement spinner = Automation.driver.findElement(By.xpath(controlName));
							waitForSpinnerToDisappear(Automation.driver, spinner);
						} catch (Exception spnrex) {
							log.info("Spinner closed");
						}
					}
					break;

				case NC:
					try {
						log.info("Inside spinner wait");
						WebElement spinner = Automation.driver.findElement(By.xpath(controlName));
						waitForSpinnerToDisappear(Automation.driver, spinner);
					} catch (Exception spnrex) {
						log.info("Spinner closed");
					}
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
				    if (ctrlValue == null || ctrlValue.trim().isEmpty() || ctrlValue.equalsIgnoreCase("Notdisplayed")) {
				        break;
				    }

				    for (int i = 0; i < 6; i++) {
				        try {

							//System.out.println("Inside SetCheckbox main try block.");

				            WebDriverWait wait = new WebDriverWait(Automation.driver, Duration.ofSeconds(6-i));
				            webElement = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));

			                boolean shouldClick = !webElement.isSelected() && (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes"));
			                boolean otherValue =   webElement.isSelected() && (!ctrlValue.equalsIgnoreCase("Y") && !ctrlValue.equalsIgnoreCase("Yes"));
			                
				            if (webElement.isEnabled() ) {
				                if (shouldClick || otherValue) {
				                	((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView(true);", webElement);
				                	webElement.click();
									Thread.sleep(500);
				                    if (webElement.isSelected() && shouldClick) {
				                        System.out.println("Checkbox is selected through webElement.click action.");
				                        break;
				                    } else if (!webElement.isSelected() && otherValue) {
				                    	System.out.println("Checkbox is unselected through webElement.click action.");
				                        break;
				                    }
				                    else {
				                        System.out.println("Checkbox is not selected through webElement.click action.");
				                        i++;
				                    }
				                } else {
				                    System.out.println("Checkbox not enabled or not required to be clicked.");
				                    i++;
				                }
				            }
				        } catch (Exception e) {

							System.out.println("Inside SetCheckbox main Catch. Try alternative option now.");
							//log.error(e.getMessage(), e);

							try {
								System.out.println("Inside SetCheckbox second try block. Try with JavascriptExecutor.");
								WebDriverWait wait = new WebDriverWait(Automation.driver, Duration.ofSeconds(6-i));
								webElement = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));

				                boolean shouldClick = !webElement.isSelected() && (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes"));
				                boolean otherValue =   webElement.isSelected() && (!ctrlValue.equalsIgnoreCase("Y") && !ctrlValue.equalsIgnoreCase("Yes"));
								
								if (webElement.isEnabled() && (shouldClick || otherValue)) {
									
									((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView(true);", webElement);
									((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();", webElement);
				                	Thread.sleep(500);
									
									if (webElement.isSelected()) {
										System.out.println("Checkbox is selected through JavascriptExecutor.");
										break;
									} else {
										System.out.println("Checkbox is not selected through JavascriptExecutor.");
										i++;
									}
								}
							} catch (Exception e1) {

								System.out.println("Last catch attempt to select Checkbox using MouseclickAction.");
								try {

					                boolean shouldClick = !webElement.isSelected() && (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes"));
					                boolean otherValue =   webElement.isSelected() && (!ctrlValue.equalsIgnoreCase("Y") && !ctrlValue.equalsIgnoreCase("Yes"));

					                if (webElement.isEnabled() && (shouldClick || otherValue)) {
										Actions mouseBuilder = new Actions(Automation.driver);
										Action mouseClickAction = mouseBuilder.moveToElement(webElement).clickAndHold().release().build();
										mouseClickAction.perform();
					                }
								} catch (Exception lastE) {
									System.out.println("Element does not exist.");
									log.error(lastE.getMessage(), lastE);
									i++;
									break;
								}
								i++;
							}
				        }
				    }
				    
				    break;
					

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
					System.out.println("clicked");
					// MousebuilderClick
					// .keyDown(Keys.CONTROL).click(webElement).keyUp(Keys.CONTROL).build().perform();
					break;
				}
				break;

			case WaitForJS:
				waitForCondition();
				break;

			case MaximizeBrowser:
				Automation.driver.manage().window().maximize();
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
							if (!StringUtils.equalsIgnoreCase(Config.applicationPlatform, "digital1st")) {
								((JavascriptExecutor) Automation.driver)
										.executeScript("arguments[0].scrollIntoView(true);", webElement);
								Thread.sleep(500);
							}

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
					try {
						if (!StringUtils.equalsIgnoreCase(Config.applicationPlatform, "digital1st")) {
							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
									webElement);
							Thread.sleep(500);
						}
						Actions builderClick = new Actions(Automation.driver);
						highlightElement(webElement);
						Action clickAction = builderClick.moveToElement(webElement).clickAndHold().release().build();
						clickAction.perform();
					} catch (Exception e) {
						log.error(e.getMessage(), e);
						System.out.println("Element not exist");
					}
					break;

				// tasleem 03-23-2020 change for Scrolling screen
				case Slider:
					WebElement slider = Automation.driver.findElement(By.xpath(controlName));
					Thread.sleep(3000);
					Actions moveSlider = new Actions(Automation.driver);
					Action actionslider = moveSlider.dragAndDropBy(slider, 180, 10).build();
					actionslider.perform();
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
				// Tasleem - NYPIUA changes
				case NC:
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

					// DW: 18-Aug-2020
					if (!StringUtils.equalsIgnoreCase(Config.applicationPlatform, "digital1st")) {
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView(true);",
								webElement);
						Thread.sleep(500);
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
					Thread.sleep(500);

					if (controlName.startsWith("//iframe")) {

						WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(700));// Sel4

						wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

						Automation.driver.switchTo().frame(Automation.driver.findElement(By.xpath(controlName)));

					}

					else {

						Automation.driver.switchTo().frame(controlName);

					}
					break;

				case NC:
					Thread.sleep(500);
					if (controlName.startsWith("//iframe")) {

						WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(700));
						wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
						Automation.driver.switchTo().frame(Automation.driver.findElement(By.xpath(controlName)));
					} else {
						Automation.driver.switchTo().frame(controlName);
					}
					break;

				case NCIF:
					try {
						if (controlName.startsWith("//iframe")) {
							Automation.driver.switchTo().frame(Automation.driver.findElement(By.xpath(controlName)));
						} else {
							Automation.driver.switchTo().frame(controlName);
						}
					} catch (Exception e) {
					}
					break;
				}
				break;

			case Browser:
			case URL:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case NewBrowser:

				switch (actionName) {
				case I: // Mayur_Extra verification...
					if (!ctrlValue.equalsIgnoreCase("null")) {
						String parentWindow = Automation.driver.getWindowHandle();
						Set<String> handles = Automation.driver.getWindowHandles();
						for (String windowHandle : handles) {
							if (!windowHandle.equals(parentWindow)) {
								Automation.driver.switchTo().window(windowHandle);
							}
						}
					}
					break;
				case NC:
					if (!ctrlValue.equalsIgnoreCase("null")) {
						String parentWindow1 = Automation.driver.getWindowHandle();
						Set<String> handles1 = Automation.driver.getWindowHandles();
						for (String windowHandle1 : handles1) {
							if (!windowHandle1.equals(parentWindow1)) {
								Automation.driver.switchTo().window(windowHandle1);
								Thread.sleep(1000);
								Automation.driver.close();
							}
						}
						Automation.driver.switchTo().window(parentWindow1);
					}
					break;
				}
				break;// case NewBrowser://Mayur

			/*
			 * Old code committed below ------------
			 * 
			 * // Thread.sleep(3000); //DS:Check if required casString parentWindow =
			 * Automation.driver.getWindowHandle(); Set<String> handles =
			 * Automation.driver.getWindowHandles(); for (String windowHandle : handles) {
			 * if (!windowHandle.equals(parentWindow)) {
			 * Automation.driver.switchTo().window(windowHandle); // <!--Perform your
			 * operation here for new window--> // Automation.driver.close(); //closing
			 * child window // Automation.driver.switchTo().window(parentWindow); // //cntrl
			 * to parent window } } break;
			 * 
			 */

			case BrowserSwitchUsingURL:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case CloseBrowser:

				Automation.driver.close();

				break;

			case CloseAndLaunchNewBrowser:
				for (int i = 1; i <= 3; i++) {
					try {
						Automation.driver.quit();
						break;
					} catch (Exception e) {
						log.error(e.getMessage(), e);
						Thread.sleep(2000);
					}
				}
				Automation.setUp();
				break;

			// case RowNumbers:
			// System.out.print("");

			case Menu:
				webElement.click();
				break;

			case ClearUniqueNumberFileContent:
				WebHelperUtil.clearUniqueNumberFileContent();
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

				break;

			case ActionClickandEsc:
				switch (actionName) {
				case NC:
					if (webElement != null) {
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								webElement);
						Thread.sleep(500);
						Actions clickandEsc = new Actions(Automation.driver);
						Action clickEscAction = clickandEsc.moveToElement(webElement).click()
								.sendKeys(Keys.ENTER, Keys.ESCAPE).build();
						clickEscAction.perform();
					} else {
						log.info("ActionClickandEsc failed as webelement is not present on screen");
					}
					break;
				case I:
					if (!ctrlValue.equalsIgnoreCase("") || ctrlValue != null) {
						if (webElement != null) {
							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
									webElement);
							Thread.sleep(500);
							Actions clickandEsc = new Actions(Automation.driver);
							Action clickEscAction = clickandEsc.moveToElement(webElement).click()
									.sendKeys(Keys.ENTER, Keys.ESCAPE).build();
							clickEscAction.perform();
						} else {
							log.info("ActionClickandEsc failed as webelement is not present on screen");
						}
					}
					break;
				}
				break;

			case ActionMouseOver:
				switch (actionName) {
				case NC:
					if (webElement != null) {
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								webElement);
						Thread.sleep(500);
						Actions builderMouserOver = new Actions(Automation.driver);
						Action mouseOverAction = builderMouserOver.moveToElement(webElement).build();
						mouseOverAction.perform();
					} else {
						log.info("ActionMouseOver failed as webelement is not present on screen");
					}
					break;
				case I:
					if (!ctrlValue.equalsIgnoreCase("") || ctrlValue != null) {
						if (webElement != null) {
							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
									webElement);
							Thread.sleep(500);
							Actions builderMouserOver = new Actions(Automation.driver);
							Action mouseOverAction = builderMouserOver.moveToElement(webElement).build();
							mouseOverAction.perform();
						} else {
							log.info("ActionMouseOver failed as webelement is not present on screen");
						}
					}
					break;
				}
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
							WebElement yearButton = Automation.driver.findElement(By.cssSelector("td:contains('«')"));
							yearButton.click();
							Monthyear2[1] = Integer.toString(Integer.parseInt(Monthyear2[1]) - 1);
						} else if (Integer.parseInt(Monthyear2[1]) < Integer.parseInt(dtMthYr[2])) {
							WebElement yearButton = Automation.driver.findElement(By.cssSelector("td:contains('»')"));
							yearButton.click();
							Monthyear2[1] = Integer.toString(Integer.parseInt(Monthyear2[1]) + 1);
						}
					}

					while (!WebHelper.month.equalsIgnoreCase(dtMthYr[1])) {
						if (Integer.parseInt(WebHelper.month) > Integer.parseInt(dtMthYr[1])) {
							WebElement monthButton = Automation.driver.findElement(By.cssSelector("td:contains('')"));
							monthButton.click();
							if (Integer.parseInt(WebHelper.month) < 11) {
								WebHelper.month = "0" + Integer.toString(Integer.parseInt(WebHelper.month) - 1);
							} else {
								WebHelper.month = Integer.toString(Integer.parseInt(WebHelper.month) - 1);
							}

						} else if (Integer.parseInt(WebHelper.month) < Integer.parseInt(dtMthYr[1])) {
							WebElement monthButton = Automation.driver.findElement(By.cssSelector("td:contains('')"));
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
			/*
			 * case Window: switch (actionName) { case O: String parentHandle =
			 * Automation.driver.getWindowHandle(); for (String winHandle :
			 * Automation.driver.getWindowHandles()) {
			 * Automation.driver.switchTo().window(winHandle);
			 * 
			 * if (Automation.driver.getTitle().equalsIgnoreCase ( controlName)) {
			 * Automation.driver.close(); }
			 * 
			 * } // Automation.driver.switchTo().window(parentHandle); break; } break;
			 */

			case Window:
				switch (actionName) {
				case O:
					String parentHandle = Automation.driver.getWindowHandle();
					for (String winHandle : Automation.driver.getWindowHandles()) {
						Automation.driver.switchTo().window(winHandle);
						if (Automation.driver.getTitle().equalsIgnoreCase(controlName)) {
							Automation.driver.close();
						}
					}
					Automation.driver.switchTo().window(parentHandle);
					break;
				default:
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

			case Tab:
				switch (actionName) {
				case I:
					if (!ctrlValue.isEmpty()) {
						webElement.sendKeys(Keys.TAB);
					}
					Thread.sleep(500);
					break;

				case NC:
					webElement.sendKeys(Keys.TAB);
					Thread.sleep(500);
					break;
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
					WebHelper.robot.keyPress(KeyEvent.VK_TAB);
					WebHelper.robot.keyRelease(KeyEvent.VK_TAB);

				} else if (controlName.equalsIgnoreCase("SPACE")) {
					WebHelper.robot.keyPress(KeyEvent.VK_SPACE);
					WebHelper.robot.keyRelease(KeyEvent.VK_SPACE);

				} else if (controlName.equalsIgnoreCase("DOWNKEY")) {
					WebHelper.robot.keyPress(KeyEvent.VK_DOWN);
					WebHelper.robot.keyRelease(KeyEvent.VK_DOWN);
					Thread.sleep(500);

				} else if (controlName.equalsIgnoreCase("ENTER")) {
					Thread.sleep(3000);
					WebHelper.robot.keyPress(KeyEvent.VK_ENTER);
					WebHelper.robot.keyRelease(KeyEvent.VK_ENTER);
					Thread.sleep(500);
					// Minaakshi : 01-03-2019
				} else if (!controlName.equalsIgnoreCase("ENTER") && !controlName.equalsIgnoreCase("SPACE")
						&& !controlName.equalsIgnoreCase("TAB") && !controlName.equalsIgnoreCase("SetFilePath")
						&& !controlName.equalsIgnoreCase("DOWNKEY")) {
					webElement = getElementByType(controlId, controlName, WebHelper.control, imageType, ctrlValue);
					webElement.sendKeys(Keys.TAB);
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

			case Database:// Minaakshi : Case added to fetch Unique values from Database
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
							String tempValue = null;
							String[] arrTempVal;
							if (readFromColName.contains(",")) {
								arrTempVal = readFromColName.split(",");
								for (int i = 0; i < arrTempVal.length; i++) {
									String temp = DMProduct.ReadFromExcelUsingColumnName("", arrTempVal[i].toString());
									tempValue = tempValue + "," + temp;
								}
							} else {
								tempValue = DMProduct.ReadFromExcelUsingColumnName("", readFromColName);
							}
							sqlQuery = DMProduct.getDataFetchingSQLQueryPAS(sqlQuery, tempValue);
						}

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
					else if (logicalName.equalsIgnoreCase("RiskCommencementDate") && !ctrlValue.equals(""))
						RiskCommencementDate = ctrlValue;
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
					else if (logicalName.equalsIgnoreCase("EntityCode") && !ctrlValue.equals(""))
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
				WebElement slider = Automation.driver.findElement(By.xpath(controlName));
				Thread.sleep(3000);
				Actions moveSlider = new Actions(Automation.driver);
				Action actionslider = moveSlider.dragAndDropBy(slider, 30, 0).build();
				actionslider.perform();
				break;
			// bhaskar

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

			case SendFileUpload:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

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
					// webElement.isDisplayed() == true
					try {
						if (webElement != null) {
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

			case DownloadDocument:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case WaitTillFileDownload:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case SaveAsDocument:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case MoveDocument:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case MoveXMLDocument:
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

			case UnzipFolderAndRename:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case XMLIgnoreTags_AttributeLevel:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case XMLIgnoreTags_NodeLevel:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case XMLCompare:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);
			case RenameFile:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);

			case MoveFile:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, null);
			case OpenAPI:
				return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName, ctrlValue,
						null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
						rowcount, rowNo, colNo, null, null, controller.controllerTransactionType.toString());

			default:
				log.info("U r in Default");
				break;
			}
			if (StringUtils.equalsIgnoreCase(Config.applicationPlatform, "policy2015")
					|| StringUtils.equalsIgnoreCase(Config.applicationPlatform, "")
					|| StringUtils.equalsIgnoreCase(Config.applicationPlatform, null)) {
				boolean isAlertMissing;
				try {
					WebDriverWait w = new WebDriverWait(Automation.driver, Duration.ofSeconds(1));// Sel4
					if (w.until(ExpectedConditions.alertIsPresent()) == null)
						isAlertMissing = true;
					else
						isAlertMissing = false;
				} catch (Exception e) {
					isAlertMissing = false;
				}

				try {
					if (isAlertMissing) { // swapnil41444 06/30/2021 isAlertMissing changes
						Thread.sleep(200);
						for (int i = 1; i <= 340; i++) {
							List<WebElement> loader = Automation.driver.findElements(By.xpath(
									"//div[@class='v-loading-indicator first' and @style='position: absolute; display: block;']"));
							int loaderSize = loader.size();
							// System.out.println("loaderSize is-->" + loaderSize);
							if (loaderSize >= 1)
								Thread.sleep(300);
							else
								break;
						}
					}
					// else System.out.println("Alert or pop-up window present. Not checking for
					// progress bar");
				} catch (Exception e) {
				}
				// To handle processing spinner wait in policy 2015 | In other applications it
				// will not create any issue.
				try {
					if (isAlertMissing) { // swapnil41444 06/30/2021 isAlertMissing changes
						WebElement spinner = Automation.driver.findElement(By.xpath("//div[@id='statusTextData']"));
						log.info("Inside spinner wait");
						waitForSpinnerToDisappear(Automation.driver, spinner);
						// log.info("Spinner closed");
					}
					// else System.out.println("Alert or pop-up window present. Not checking for
					// progress bar");
				} catch (Exception spnrex) {
				}

			}

			if (StringUtils.equalsIgnoreCase(Config.applicationPlatform, "digital1st")) {
				// To handle processing icon in digital 1st | In other applications it will not
				// create any issue
				try {
					Thread.sleep(200);
					for (int i = 1; i <= 340; i++) {
						List<WebElement> loader = Automation.driver.findElements(By.xpath(
								"//body[contains(@class,'ice-runtime-body-background-color') and @style='cursor: progress;']"));
						int loaderSize = loader.size();
						// System.out.println("loaderSize is-->" + loaderSize);
						if (loaderSize >= 1)
							Thread.sleep(300);
						else
							break;
					}
				} catch (Exception e) {
				}
			}
		} catch (WebDriverException we) {
			log.error(we.getMessage(), we);
			saveScreenShot(controlTypeEnum.toString());
			throw new Exception("Error Occurred from Do Action " + controlName + we.getMessage());
		} catch (Exception e) {
			//log.error(e.getMessage(), e);//
			//String errorMessage = e.getMessage();
			//String firstLine = errorMessage.split("\\r?\\n")[0];
			String errorMessage = (e.getMessage() != null && !e.getMessage().isEmpty()) ? e.getMessage(): "No error message available";
			String firstLine = errorMessage.split("\\r?\\n", 2)[0];
			log.error(firstLine);
			
			
			throw new Exception(e.getMessage());
		}

		//Skip the step for which control value is blank end

		// TM-02/02/2015: Radio button found ("F") & AJAX control ("VA")
				if ((action.equalsIgnoreCase("V") || action.equalsIgnoreCase("F") || action.equalsIgnoreCase("PopUpCount")
				|| action.equalsIgnoreCase("VerifyPopUpElement") || action.equalsIgnoreCase("VA")
				|| action.toString().equalsIgnoreCase("VD") || action.toString().equalsIgnoreCase("V_text")
				|| action.toString().equalsIgnoreCase("V_availabilityStatus")
				|| action.toString().equalsIgnoreCase("V_format") || action.toString().equalsIgnoreCase("V_date")
				|| action.toString().equalsIgnoreCase("V_attributeValue")
				|| action.toString().equalsIgnoreCase("VFCount") || action.toString().equalsIgnoreCase("VFPresence")
				|| action.toString().equalsIgnoreCase("V_checkboxStatus")
				|| action.toString().equalsIgnoreCase("V_toggleDisableStatus")
				|| action.toString().equalsIgnoreCase("V_edit")
				|| action.toString().equalsIgnoreCase("V_toggleCheckboxStatus")
				|| action.toString().equalsIgnoreCase("V_toggleOptionStatus")
				|| action.toString().equalsIgnoreCase("V_toggleOptionDisableStatus")
				|| action.toString().equalsIgnoreCase("V_buttonOptionStatus")
				|| action.toString().equalsIgnoreCase("V_backgroundStatus")
				|| action.toString().equalsIgnoreCase("V_disableStatus")
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
		WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(30));// Sel4
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

	// swapnil41444 07/07/2021: commenting code related to TakesScreenshot
	// functionality...
	// Instead adding code for AShot. Also added pom.xml reference for the same.
	// reason for change: 1) TakesScreenshot doesn't scroll the page while taking
	// the screenshot
	// 2) When iTAF integrated with Jenkins, Jenkins executes scripts in 4*3 screen
	// ratio
	// & then barely gets any screen details as majority details needs page
	// scrolling

public static void saveScreenShot(String controlType) {
		if (ITAFWebDriver.isSuiteApplication()) {
			WebHelperUtil.saveScreenShot("");
		} else {
			if (!(Automation.driver instanceof TakesScreenshot)) {

				System.out.println(
						"Not able to take screenshot: Current WebDriver does not support TakesScreenshot interface.");
				return;
			}
			// Sel4
			TakesScreenshot screenshotDriver = (TakesScreenshot) Automation.driver;
			File screenshotFile = null;
		     try {
		            // Get screen size using Toolkit
		            java.awt.Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		            int width = (int) screenSize.getWidth();
		            int height = (int) screenSize.getHeight();
		            log.info("Screen size: " + width + "x" + height);

		            // Set the browser window size
		            if (Boolean.parseBoolean(Config.headlessBrowser)) {
		                Dimension browserSize = new Dimension(width, height);
		                Automation.driver.manage().window().setSize(browserSize);
		            }
				screenshotFile = screenshotDriver.getScreenshotAs(OutputType.FILE);

			} catch (Exception e) {
				log.error(e.getMessage(), e);
				System.out.println("Taking screenshot failed for: " + webDriver.getReport().getTestcaseId());
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
					Files.copy(screenshotFile.toPath(), new File(location).toPath());

				} catch (IOException e) {
					log.error(e.getMessage(), e);
					e.printStackTrace();
					return;
				}
			} else {
				try {
					String subfolder = Config.resultFilePath + "\\ScreenShots";
					File createDir = new File(subfolder);
					if (!createDir.exists()) {
						createDir.mkdirs();
					}

					location = subfolder + "\\" + fileName + ".jpeg";
					// controller.testDescription = location;
					webDriver.getReport().setScreenShot("file:\\\\" + location);
					// ImageIO.write(screenshot.getImage(), "jpeg", new File(new_location));

					// Files.copy(screenshotFile.toPath(), new File(location).toPath());
					Files.copy(screenshotFile.toPath(), new File(location).toPath(),
							StandardCopyOption.REPLACE_EXISTING);

					controller.FailScreen = new File(location).getAbsolutePath();
				} catch (IOException e) {
					//String errorMessage = e.getMessage();
					//String firstLine = errorMessage.split("\\r?\\n")[0];
					String errorMessage = (e.getMessage() != null && !e.getMessage().isEmpty()) ? e.getMessage(): "No error message available";
					String firstLine = errorMessage.split("\\r?\\n", 2)[0];
					log.error("Error :" + firstLine);
					// log.error(e.getMessage(), e);
					// e.printStackTrace();
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

	/**
	 * Author: DW 16/Sept/2019 - This function is used for to wait till spinner goes
	 * off.
	 **/
	public static boolean waitForSpinnerToDisappear(WebDriver driver, final WebElement spinnerName) {
		/*
		 * FluentWait<WebDriver> fWait = new FluentWait<WebDriver>(driver);
		 * fWait.withTimeout(180, TimeUnit.SECONDS); fWait.pollingEvery(1,
		 * TimeUnit.SECONDS); fWait.ignoring(NoSuchElementException.class);
		 */
		// Sel4

		FluentWait<WebDriver> fWait = new FluentWait<>(driver);
		fWait.withTimeout(Duration.ofSeconds(180));
		fWait.pollingEvery(Duration.ofSeconds(1));
		fWait.ignoring(NoSuchElementException.class);

		Function<WebDriver, Boolean> func = new Function<WebDriver, Boolean>() {
			@Override
			public Boolean apply(WebDriver driver) {
				// System.out.println(spinnerName.getCssValue("display"));
				if (spinnerName.getCssValue("display").equalsIgnoreCase("none")) {
					return true;
				}
				return false;
			}
		};
		return fWait.until(func);
	}
}
