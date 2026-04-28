package com.majesco.itaf.main;

import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.time.Duration;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
//import java.util.Map;
import java.util.Random;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.time.format.DateTimeFormatter;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementNotInteractableException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.sikuli.script.Screen;
import com.majesco.itaf.recovery.StartRecoveryClaims;
import com.majesco.itaf.util.BillingProduct;
import com.majesco.itaf.util.Constants;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.verification.WebVerification;
import com.majesco.itaf.vo.Reporter;
import com.majesco.itaf.util.DMProduct;
import com.majesco.itaf.util.JDBCConnection;

public class WebHelperClaims {

	public static String claimsloader = "//div[@class='overlay']/div[@class='logo-wrapper']/div";
	public static String CRMloader = "//div[@class='spinner-three-bounce full-screen']";
	public static Reporter report = new Reporter();
	private final static Logger log = LogManager.getLogger(WebHelperClaims.class.getName());
	public static Cell WebservicecycleDate = null;
	public static String wscycledate;
	public static Boolean blank = false;
	@SuppressWarnings("unused")
	private static HashMap<String, Object> maincontrolsheet = new HashMap<String, Object>();
	@SuppressWarnings("unused")
	private static Row currRow = null;
	public static boolean recovery_done = false;
	@SuppressWarnings("unused")
	private static int current_SC_NO = 0;
	@SuppressWarnings("unused")
	private static int currentscenario_num = 0;
	@SuppressWarnings("unused")
	private static int failedscenario_num = 0;
	public static Sheet MainControllerSheet = null;
	public static Boolean colnotfound = false;
	public static HashMap<String, Object> vColumnheaderIndex = new HashMap<String, Object>();
	public static HashMap<String, Object> vColumnheaderValues = new HashMap<String, Object>();
	public static Date toDate = null;
	public static String testCase = null;
	public static Screen sikuliScreen = null;
	public static List<String> searchValue1 = null;
	public static Boolean pageLoaded = false;
	public static String sqlQuery = "";
	public static String readFromColName = "";
	public static String toBeFetchedDBColName = "";
	public static String expectedDBStatus = "";
	public static String writeToColName = "";
	public static String verifyColNameExpected = "";
	public static Cell transactionType = null;
	public static int TotalpassCount;
	public static int TotalfailCount;
	public static int DResult = 1;
	private static List<String> formList = new ArrayList<String>();
	private static List<String> presentForms = new ArrayList<String>();
	private static List<String> additionalForms = new ArrayList<String>();
	private static List<String> missingForms = new ArrayList<String>();
	private static List<String> expectedForms = new ArrayList<String>();
	private static List<String> lookupValuesList = new ArrayList<String>();
	private static List<String> presentLookupValues = new ArrayList<String>();
	private static List<String> additionalLookupValues = new ArrayList<String>();
	private static List<String> missingLookupValues = new ArrayList<String>();
	private static List<String> expectedLookupValues = new ArrayList<String>();
	static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	static MainController controller = ObjectFactory.getMainController();

	static void implementWait() throws InterruptedException {
		Automation.driver = WebHelper.currentdriver;
		WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(90));
		WaitForPageLoad.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
		WebHelper.wait = new FluentWait<>(WebHelper.currentdriver).withTimeout(Duration.ofSeconds(60))
				.pollingEvery(Duration.ofSeconds(5)).ignoring(NoSuchElementException.class)
				.ignoring(ElementNotInteractableException.class);
		// .ignoring(ElementNotInteractableException.class);//Sel4
		// .ignoring(ElementNotFoundException.class);

	}

	public static void calldoAction(Sheet headerValues, String logicalName, Row rowValues, String TransactionType,
			int valuesRowIndex, String action, String controlName, Sheet sheetStructure, int rowIndex,
			String TestCaseID, String controltype, String controlID, String indexVal, String imageType, String FilePath,
			int rowCount, String rowNo, String colNo, String operationType) throws Exception {
		String ctrlValue1 = null;
		String ctrlValue2 = null;
		String ctrlValue = null;
		String cycleDate = TestCaseID;
		WebElement webElement = null;
		List<WebElement> controlList = null;
		colnotfound = false;

		if (WebHelper.valuesHeader.isEmpty() == true) {
			WebHelper.valuesHeader = WebHelperUtil.getValueFromHashMap(headerValues);
		}
		Object actualValue = null;

		if (logicalName != null) {
			actualValue = WebHelper.valuesHeader.get(logicalName.toString());
		} // headerRow.getCell(colIndex);

		if (actualValue == null) {
			colnotfound = true; // log.info("Null");
		}

		WebHelper.testcaseID = rowValues.getCell(Integer.parseInt(WebHelper.valuesHeader.get("TestCaseID").toString()));
		if (WebHelper.testcaseID == null) {
			testCase = "";
		} else {
			testCase = WebHelper.testcaseID.toString();
		}
		WebHelper.transactionType = rowValues
				.getCell(Integer.parseInt(WebHelper.valuesHeader.get("TransactionType").toString()));
		@SuppressWarnings("unused")
		String stransactionType = TransactionType.toString();

		if (colnotfound == false) {
			ctrlValue = WebHelperUtil.getCellData(logicalName, headerValues, valuesRowIndex, WebHelper.valuesHeader);
		} else {
			ctrlValue = "";
		}
		Pattern trimregex = Pattern.compile("^\\s+|\\s+$");
		Matcher match = trimregex.matcher(ctrlValue);
		StringBuffer ctrlValue_output = new StringBuffer();
		while (match.find())
			match.appendReplacement(ctrlValue_output, "");
		match.appendTail(ctrlValue_output);
		if (action.equalsIgnoreCase("Capture")) {
			// Reporter report = new Reporter();
			log.info("Inside Capture Case");
			controlName = WebHelperUtil.getCellData("ControlName", sheetStructure, rowIndex, WebHelper.structureHeader);
			logicalName = WebHelperUtil.getCellData("LogicalName", sheetStructure, rowIndex, WebHelper.structureHeader);
			if (ctrlValue.equalsIgnoreCase("Y")) {
				// log.info("CYCLEDATE is :" + cycleDate);
				TransactionMapping.TransactionCaptureData("NA", TestCaseID, controlName,
						Config.transactionInputFilePath);
			}
		}

		if (((action.equals("I") && !StringUtils.isEmpty(ctrlValue))
				|| (action.equals("InputDate") && !StringUtils.isEmpty(ctrlValue))
				|| (action.equals("V_text") && !ctrlValue.isEmpty())
				|| (action.equals("V_attributeValue") && !StringUtils.isEmpty(ctrlValue))
				|| (action.equals("T") && !StringUtils.isEmpty(ctrlValue))
				|| (action.equals("V") && !StringUtils.isEmpty(ctrlValue))
				|| !(action.equals("I") || action.equals("V_text"))) && !action.equalsIgnoreCase("Capture")
				&& !action.equalsIgnoreCase("FIND") && !action.equalsIgnoreCase("TABLEINPUT")
				&& !(action.equals("CA") && ctrlValue.isEmpty())) {

			if ((logicalName.equalsIgnoreCase("WAIT")) || (controltype.equalsIgnoreCase("WAIT"))) {
				log.info("Wait time is:" + controlName + "second");
			} else {
				log.info(action + " On " + controltype + " " + logicalName);
				log.info("ctrlValue is :" + ctrlValue);
			}

			if (logicalName.equalsIgnoreCase("CreateBatch")) {
				log.info("wait");
			}

			if (!controltype.startsWith("Sikuli")) {
				if (!action.equalsIgnoreCase("V")
						// && !action.equals("V_attributeValue")
						&& !action.equalsIgnoreCase("LOOP") && !controltype.equalsIgnoreCase("Wait")
						&& !controltype.equalsIgnoreCase("Wait_IfValue") && !controltype.equalsIgnoreCase("WaitToLoad")
						&& !action.equalsIgnoreCase("END_LOOP") && !controltype.equalsIgnoreCase("Browser")
						&& !controltype.equalsIgnoreCase("NewBrowser") && !controltype.equalsIgnoreCase("CloseBrowser")
						&& !controltype.equalsIgnoreCase("CloseAndLaunchNewBrowser")
						&& !controltype.equalsIgnoreCase("ParentWindow")
						&& !controltype.equalsIgnoreCase("Window") && !controltype.equalsIgnoreCase("Alert")
						&& !controltype.equalsIgnoreCase("URL") && !controltype.equalsIgnoreCase("WaitForJS")
						&& !controltype.contains("Robot") && !controltype.equalsIgnoreCase("Calendar")
						&& !controltype.equalsIgnoreCase("CalendarNew") && !controltype.equalsIgnoreCase("CalendarIPF")
						&& !controltype.equalsIgnoreCase("CalendarEBP")
						&& (!action.equalsIgnoreCase("Read")
								|| ((action.equalsIgnoreCase("Read") && !StringUtils.isEmpty(controlName))))
						&& (!action.equalsIgnoreCase("Read")
								|| ((action.equalsIgnoreCase("Read") && !StringUtils.isEmpty(controlName))))// Repeated
						&& (!action.equalsIgnoreCase("Read2")
								|| (action.equalsIgnoreCase("Read2") && !StringUtils.isEmpty(ctrlValue)))
						&& (!action.equalsIgnoreCase("VFormCount")
								|| (action.equalsIgnoreCase("VFormCount") && !StringUtils.isEmpty(ctrlValue)))
						&& !action.equalsIgnoreCase("VFormPresence")
						&& !action.equalsIgnoreCase("FormVerificationDetails")
						&& (!action.equalsIgnoreCase("VLookupValuesCount")
								|| (action.equalsIgnoreCase("VLookupValuesCount") && !StringUtils.isEmpty(ctrlValue)))
						&& !action.equalsIgnoreCase("VLookupValuesPresence")
						&& !action.equalsIgnoreCase("LookupVerificationDetails")
						&& (!action.equalsIgnoreCase("V_date")
								|| (action.equalsIgnoreCase("V_date") && !StringUtils.isEmpty(ctrlValue)))
						&& (!action.equalsIgnoreCase("V_availabilityStatus")
								|| (action.equalsIgnoreCase("V_availabilityStatus") && !StringUtils.isEmpty(ctrlValue)))
						&& !controltype.equalsIgnoreCase("JSScript") && !controltype.equalsIgnoreCase("DB")
						&& !controlID.equalsIgnoreCase("XML") && !controltype.startsWith("Process")
						&& !controltype.startsWith("Destroy") && !controltype.startsWith("ReadSikuli")
						&& !controltype.equalsIgnoreCase("WebService") && !action.equalsIgnoreCase("VA")
						&& !action.equalsIgnoreCase("FileCompare") && !controltype.equalsIgnoreCase("Screenshot")// devishree
						&& !controltype.equalsIgnoreCase("WebService1") && !controltype.equalsIgnoreCase("WebService2")
						&& !controltype.equalsIgnoreCase("WebService3") && !controltype.equalsIgnoreCase("WebServiceV")
						&& !controltype.equalsIgnoreCase("WebServiceC") && !controltype.equalsIgnoreCase("WebServiceRP")
						&& !controltype.equalsIgnoreCase("WebServiceV1")
						&& !controltype.equalsIgnoreCase("WebServiceV2")
						&& !controltype.equalsIgnoreCase("WebServiceVAG")
						&& !controltype.equalsIgnoreCase("WebServiceV3") && !controltype.equalsIgnoreCase("Multiselect") // devishree
						&& !controltype.equalsIgnoreCase("PDFDocumentCompare")
						&& !controltype.equalsIgnoreCase("MoveDocument")) {
					if ((indexVal.equalsIgnoreCase("") || indexVal.equalsIgnoreCase("0"))
							&& !controlID.equalsIgnoreCase("TagValue") && !controlID.equalsIgnoreCase("TagText")
							&& !action.equalsIgnoreCase("NoException") && !action.equalsIgnoreCase("FIND")
							&& !action.equalsIgnoreCase("TABLEINPUT")) {

						try {
							if (controlName.contains("+")) {
								controlName = controlName.replace("+", ctrlValue);
							}
							for (int i = 0; i < 25; i++) {
								if (((JavascriptExecutor) WebHelper.currentdriver)
										.executeScript("return document.readyState").toString().equals("complete")) {
									break;
								} else {
									Thread.sleep(1000);
								}
							}
							if ((!(logicalName.equalsIgnoreCase("Last"))
									|| (!(logicalName.equalsIgnoreCase("PaginationLast"))))
									&& (!controlName.contains("//li[@class='last']"))) {
								webElement = getElementByType(controlID, controlName, WebHelper.control, imageType,
										ctrlValue, logicalName);
							}
						} catch (Exception e) {
							throw new Exception(e.getMessage());
						}

					}
					// bhaskar Supressing exception when element not found START
					else if (action.equalsIgnoreCase("NoException") || action.equalsIgnoreCase("FIND")
							|| action.equalsIgnoreCase("TABLEINPUT")) {
						Boolean elementexists = false;
						Constants.ControlIdEnum scontrolID = Constants.ControlIdEnum.valueOf(controlID);
						Thread.sleep(1000);

						switch (scontrolID) {
						case Id:
							elementexists = WebHelper.currentdriver.findElements(By.id(controlName)).size() > 0;
							break;

						case XPath:
							elementexists = WebHelper.currentdriver.findElements(By.xpath(controlName)).size() > 0;
							break;

						case Name:
							elementexists = WebHelper.currentdriver.findElements(By.name(controlName)).size() > 0;
							break;

						case ClassName:
							elementexists = WebHelper.currentdriver.findElements(By.className(controlName)).size() > 0;
							break;

						case LinkText:
							elementexists = WebHelper.currentdriver.findElements(By.linkText(controlName)).size() > 0;
							break;

						case CSSSelector:
							elementexists = WebHelper.currentdriver.findElements(By.cssSelector(controlName))
									.size() > 0;
							break;

						case Id_p:
						case HTMLID_p:
							elementexists = WebHelper.currentdriver.findElements(By.id(controlName)).size() > 0;
							break;

						case XPath_p:
							elementexists = WebHelper.currentdriver.findElements(By.xpath(controlName)).size() > 0;
							break;

						case XPath_if:

							try {
								elementexists = WebHelper.currentdriver.findElements(By.xpath(controlName)).size() > 0;
							} catch (Exception e) {
								// log.error(e.getMessage(), e);
								elementexists = null;
								log.info("XPATH not found on UI: " + controlName);
							}
							break;

						default:
							break;
						}

						if (elementexists == true) {
							try {

								// Passing logicalName for CLaims Autouser creation feature- sync user issue -
								// 10\07\2023
								webElement = getElementByType(controlID, controlName, WebHelper.control, imageType,
										ctrlValue, logicalName);
							} catch (Exception e) {
								throw new Exception(e.getMessage());
							}

						} else {
							return;
						}
					}
					// bhaskar Supressing exception when element not found END
					else {
						controlList = WebHelperUtil.getElementsByType(controlID, controlName, WebHelper.control,
								imageType, ctrlValue);

						if (controlList != null && controlList.size() > 1) {
							String expMessage = "";
							try {
								webElement = WebHelperUtil.GetControlByIndex(indexVal, controlList, controlID,
										controlName, WebHelper.control, ctrlValue); // , ISelenium selenium)
							} catch (NoSuchElementException nse) {
								expMessage = "Failed to find Elements using FindBy with index. " + nse.getMessage();
								StartRecoveryClaims.initiateRecovery(expMessage);
								throw new Exception(expMessage);
							} catch (StaleElementReferenceException sere) {
								expMessage = "Element is no longer appearing on the DOM page with index. "
										+ sere.getMessage();
								StartRecoveryClaims.initiateRecovery(expMessage);
								throw new Exception(expMessage);
								/*
								 * catch (ElementNotVisibleException env) { expMessage =
								 * "Element is not visible with index. " + env.getMessage();
								 * StartRecoveryClaims.initiateRecovery(expMessage); throw new
								 * Exception(expMessage);
								 */// Sel4
							} catch (Exception ex) {
								expMessage = ex.getMessage();
								StartRecoveryClaims.initiateRecovery(expMessage);
								throw new Exception(expMessage);
							}

						} else {
							return;
						}
					}
				}
			} else {
				sikuliScreen = new Screen();
			}
		}

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
				if (tempValues2[j].trim().equals("pickFromUniqueNumbers")) {
					tempValues2[j] = WebHelperUtil.ReadFromExcel(tempValues2[j].trim(), WebHelper.columnName);
				}
				tempReplaceString = tempReplaceString.replace("$" + k + "value", tempValues2[j].trim());
				k++;
			}
			controlName = tempReplaceString;
		}

		/*** Perform action on the identified control ***/
		// log.info("go to method doAction");
		if (!action.equalsIgnoreCase("Capture")) {
			doAction(FilePath, rowValues, testCase, imageType, controltype, controlID, controlName, ctrlValue,
					ctrlValue1, ctrlValue2, wscycledate, logicalName, action, webElement, true, sheetStructure,
					headerValues, rowIndex, rowCount, rowNo, colNo, operationType, cycleDate, TransactionType);
		}
	}

	// Adding "WriteToDetailResults" for PDF Comparision from

	public static Reporter WriteToDetailResults(String expectedValue, String actualValue, String columnName)
			throws IOException

	{
		if (WebHelper.file.exists() == true && DResult == 1) {
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
		if ((expectedValue.trim()).equalsIgnoreCase(actualValue.trim())) {
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

		// if(file.exists() == false)
		if (WebHelper.file.exists() == false) {
			WebHelper.print = new PrintStream(WebHelper.file);
		}

		WebHelper.print = new PrintStream(new FileOutputStream(WebHelper.file, true));

		int usedRows = WebHelperUtil.count(WebHelper.file);
		if (usedRows == 0) {
			WebHelper.print
					.print("Iteration,TestCaseID,TransactionType,CurrentDate,RowType,Status,PassCount,FailCount");
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
	public static WebElement getElementByType(String controlId, String controlName, String controlType,
			String imageType, String controlValue, String logicalName) throws Exception {

		WebElement controlList = null;
		String expMessage;
		try {

			if (controlId.equalsIgnoreCase("Id") || controlId.equalsIgnoreCase("HTMLID")) {
				controlList = WebHelper.wait.until(ExpectedConditions.elementToBeClickable(By.id(controlName)));

			} else if (controlId.equalsIgnoreCase("XPath")) {
				controlList = WebHelper.wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

			} else if (controlId.equalsIgnoreCase("XPath_WithBlankValue")) {
				if (!controlValue.equalsIgnoreCase("")) {
					controlList = WebHelper.wait
							.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
				} else {
					log.info("Blank value found");
				}
			} else if (controlId.equalsIgnoreCase("Name")) {
				controlList = WebHelper.wait.until(ExpectedConditions.presenceOfElementLocated(By.name(controlName)));
			} else if (controlId.equalsIgnoreCase("ClassName")) {
				controlList = WebHelper.wait.until(ExpectedConditions.elementToBeClickable(By.className(controlName)));
			} else if (controlId.equalsIgnoreCase("LinkText") || controlId.equalsIgnoreCase("LinkValue")) {
				controlList = WebHelper.wait.until(ExpectedConditions.elementToBeClickable(By.linkText(controlName)));
			} else if (controlId.equalsIgnoreCase("TagText") || controlId.equalsIgnoreCase("TagValue")
					|| controlId.equalsIgnoreCase("TagOuterText")) {
				controlList = WebHelper.wait.until(ExpectedConditions.elementToBeClickable(By.tagName(imageType)));
			} else if (controlId.equalsIgnoreCase("CSSSelector")) {
				controlList = WebHelper.wait
						.until(ExpectedConditions.elementToBeClickable(By.cssSelector(controlName)));
			} else if (controlId.equalsIgnoreCase("AjaxPath")) {
				controlList = WebHelper.wait.until(ExpectedConditions
						.elementToBeClickable(By.xpath(controlName + "[contains(text(),'" + controlValue + "')]")));
			} else if (controlId.equalsIgnoreCase("AutoUserPath")) {
				controlList = WebHelper.wait.until(ExpectedConditions
						.presenceOfElementLocated(By.xpath(controlName + "[contains(text(),'" + controlValue + "')]")));
			} else if (controlId.equalsIgnoreCase("AjaxPath_Dynamic")) {
				controlList = WebHelper.wait.until(ExpectedConditions
						.elementToBeClickable(By.xpath(controlName + "[contains(text(),'" + controlValue)));
			} else if (controlId.equalsIgnoreCase("Id_p") || controlId.equalsIgnoreCase("HTMLID_p")) {
				controlList = WebHelper.wait.until(ExpectedConditions.presenceOfElementLocated(By.id(controlName)));
			} else if (controlId.equalsIgnoreCase("XPath_p")) {
				controlList = WebHelper.wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
			} else if (controlId.equalsIgnoreCase("Xpath_Dynamic")) {
				controlList = WebHelper.wait.until(ExpectedConditions.elementToBeClickable(
						By.xpath("//div[contains(text(),'" + controlValue + "')]" + controlName)));

			} else if (controlId.equalsIgnoreCase("XPathValue")) {
				if (!controlValue.equalsIgnoreCase("")) {
					if (controlValue.trim().equals("pickFromUniqueNumbers")) {
						controlValue = WebHelperUtil.ReadFromExcel(controlValue, WebHelper.columnName);
					}
					String tempCtrlName = controlName;
					String tempReplaceString = tempCtrlName.replace("$Value", controlValue);
					controlList = WebHelper.wait
							.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString)));
				} else {
					log.info("No Value to select element");
				}
			}

			else if (controlId.equalsIgnoreCase("XPath2$Value")) {
				if (!controlValue.equalsIgnoreCase("")) {
					if (controlValue.contains(";")) {
						String[] MultipleControlValue = controlValue.split(";");

						String tempControlname1 = controlName.replace("$Value1", MultipleControlValue[0]);
						String tempControlname2 = tempControlname1.replace("$Value2", MultipleControlValue[1]);
						// Thread.sleep(500);
						WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(60));
						WaitForPageLoad.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
						controlList = WebHelper.wait
								.until(ExpectedConditions.elementToBeClickable(By.xpath(tempControlname2)));
					}

				} else {
					log.info("No Value to select element");
				}
			}

			else if (controlId.equalsIgnoreCase("XPathClaimNo")) {
				String tempCtrlName = controlName;
				String uniqueNumber = ReadFromExcel(controlValue);
				String tempReplaceString = tempCtrlName.replace("$Value", uniqueNumber);
				controlList = WebHelper.wait
						.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString)));
			}

			else if (controlId.equalsIgnoreCase("XPath_R")) {
				if (!controlValue.isEmpty()) {

					try {

						controlList = WebHelper.wait
								.until(ExpectedConditions.elementToBeClickable(By.xpath(controlValue)));

					} catch (Exception e) {
						log.error(e.getMessage(), e);
						controlList = null;
					}
				}
			} else if (controlId.equalsIgnoreCase("XPath_if")) {
				WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(3));
				try {

					if (wait1.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName))).isDisplayed()) {
						controlList = wait1.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
					}
				} catch (Exception e) {
					// log.error(e.getMessage(), e);
					controlList = null;
				}
			} else if (controlId.equalsIgnoreCase("RetriveFormsCountAndName")) {
				missingForms.removeAll(missingForms);
				additionalForms.removeAll(additionalForms);
				presentForms.removeAll(presentForms);
				expectedForms.removeAll(expectedForms);
				formList.removeAll(formList);
				if (!controlValue.trim().equals("")) {
					List<WebElement> forms = Automation.driver.findElements(By.xpath(controlName));
					for (int i = 0; i < forms.size(); i++) {
						formList.add(forms.get(i).getText());
					}
				}
			} else if (controlId.equalsIgnoreCase("RetriveLookupValuesCountAndName")) {

				missingLookupValues.removeAll(missingLookupValues);
				additionalLookupValues.removeAll(additionalLookupValues);
				presentLookupValues.removeAll(presentLookupValues);
				expectedLookupValues.removeAll(expectedLookupValues);
				lookupValuesList.removeAll(lookupValuesList);
				if (!controlValue.trim().equals("")) {
					List<WebElement> lookupValues = Automation.driver.findElements(By.xpath(controlName));
					for (int i = 0; i < lookupValues.size(); i++) {
						if (!lookupValues.get(i).getText().trim().equals(""))
							lookupValuesList.add(lookupValues.get(i).getText());
					}
				}
			} else if (controlId.equalsIgnoreCase("XPathValueMultiWithInput_p")) {
				try {
					WebDriverWait wait2 = new WebDriverWait(Automation.driver, Duration.ofSeconds(2));
					String tempcontrolValue3 = controlValue;
					String[] tempValues3 = tempcontrolValue3.split(";;");
					String tempReplaceString6 = controlName;
					int s = 1;
					for (int j = 0; j < tempValues3.length - 1; j++) {
						if (tempValues3[j].trim().equals("pickFromUniqueNumbers")) {
							tempValues3[j] = WebHelperUtil.ReadFromExcel(tempValues3[j].trim(), WebHelper.columnName);
						}
						tempReplaceString6 = tempReplaceString6.replace("$" + s + "value", tempValues3[j].trim());
						s++;
					}
					controlList = wait2.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString6)));
				} catch (Exception e) {
					controlList = null;
				}
			} else if (controlId.equalsIgnoreCase("XPathValueMultiWithInput")) {
				WebDriverWait wait = new WebDriverWait(Automation.driver, Duration.ofSeconds(2));
				String tempcontrolValue2 = controlValue;
				String[] tempValues2 = tempcontrolValue2.split(";;");
				String tempReplaceString4 = controlName;
				int k = 1;
				for (int j = 0; j < tempValues2.length - 1; j++) {
					if (tempValues2[j].trim().equals("pickFromUniqueNumbers")) {
						tempValues2[j] = WebHelperUtil.ReadFromExcel(tempValues2[j].trim(), WebHelper.columnName);
					}
					tempReplaceString4 = tempReplaceString4.replace("$" + k + "value", tempValues2[j].trim());
					k++;
				}
				controlList = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString4)));
			} else if (controlId.equalsIgnoreCase("XPathValue_p")) {
				try {
					WebDriverWait wait = new WebDriverWait(Automation.driver, Duration.ofSeconds(2));
					if (!controlName.equals("")) {
						String tempCtrlName1 = controlName;
						String tempReplaceString1 = tempCtrlName1.replace("$value", controlValue);
						controlList = wait
								.until(ExpectedConditions.presenceOfElementLocated(By.xpath(tempReplaceString1)));
					} else {
						controlList = null;
					}
				} catch (Exception e) {
					controlList = null;
				}
			}

			else if (controlId.equalsIgnoreCase("XPath_value")) {
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

							controlList = WebHelper.wait
									.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));

						}

					} catch (Exception e) {
						log.error(e.getMessage(), e);
						controlList = null;
					}
				} else {

					System.out.print("don't perform any action");
				}
			}

			return controlList;
		} catch (NoSuchElementException nse) {
			expMessage = "Failed to find Element(s). " + nse.getMessage();
			StartRecoveryClaims.initiateRecovery(expMessage);
			throw new Exception(expMessage);
		} catch (StaleElementReferenceException sere) {
			expMessage = "Element is no longer appearing on the DOM page. " + sere.getMessage();
			StartRecoveryClaims.initiateRecovery(expMessage);
			throw new Exception(expMessage);
		} catch (ElementNotInteractableException env) {
			expMessage = "Element is not visible. " + env.getMessage();
			StartRecoveryClaims.initiateRecovery(expMessage);
			throw new Exception(expMessage);
		} catch (Exception ex) {
			expMessage = ex.getMessage();
			StartRecoveryClaims.initiateRecovery(expMessage);
			throw new Exception(expMessage);
		}
	}

	public static String doAction(String FilePath, Row rowValues, String testCase, String imageType, String controlType,
			String controlId, String controlName, String ctrlValue, String ctrlValue1, String ctrlValue2,
			String wscycledate, String logicalName, String action, WebElement webElement, Boolean Results,
			Sheet strucSheet, Sheet valSheet, int rowIndex, int rowcount, String rowNo, String colNo,
			String operationType, String cycleDate, String TransactionType)
			throws WebDriverException, IOException, Exception {

		@SuppressWarnings("unused")
		String clocation;
		List<WebElement> WebElementList = null;
		String currentValue = null;
		String uniqueNumber = "";
		wscycledate = testCase;
		WebVerification.isFromVerification = false;
		Constants.ControlTypeEnum controlTypeEnum = Constants.ControlTypeEnum.valueOf(controlType);
		Constants.ControlTypeEnum actionName = Constants.ControlTypeEnum.valueOf(action.toString());
		WebHelper.sikscreen = Config.SikuliScr;
		if (controlType.contains("Robot") && !WebHelper.isIntialized) {
			log.info("In method doaction debug1");
			WebHelper.robot = new Robot();
			WebHelper.isIntialized = true;
		}

		if (!WebHelperUtil.stringIn(action, new String[] { "I", "V", "F", "VA", "VV", "CA" })
				|| !ctrlValue.equalsIgnoreCase("")) {

			try {
				switch (controlTypeEnum) {

				case WebEdit_NoScroll:
					switch (actionName) {
					case I:
						if (ctrlValue != null && !ctrlValue.trim().equalsIgnoreCase("")) {
							// if (ctrlValue != null || !ctrlValue.trim().equalsIgnoreCase("")) {//Sel4
							log.info("ctrlValue is :" + ctrlValue);
							// Thread.sleep(500);
							webElement.clear();
							webElement.sendKeys(ctrlValue);
							Thread.sleep(100);
						}
						break;
					default:
						break;
					}
					break;
				case WebEdit:
					switch (actionName) {
					case Read:
						uniqueNumber = ReadFromExcel(ctrlValue);
						log.info(uniqueNumber);
						claimLoaderWait();
						// log.info("!!!!!!!!!!!!!!!!");
						// log.info("uniqueNumber:"+uniqueNumber);
						webElement.clear();
						webElement.sendKeys(uniqueNumber);
						break;

					case Read_if:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
							uniqueNumber = ReadFromExcel(ctrlValue);
							log.info(uniqueNumber);
							claimLoaderWait();
							webElement.clear();
							webElement.sendKeys(uniqueNumber);
						}
						break;

					case Read2:
						if (ctrlValue.trim().equals(""))
							break;

						uniqueNumber = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
						webElement.clear();
						Thread.sleep(500);
						webElement.sendKeys(uniqueNumber);
						Thread.sleep(500);
						webElement.sendKeys(Keys.TAB);
						break;

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
								WebHelperPAS.doAction(null, "ReplaceDefault", null, null, ctrlValue, null, "I",
										webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo);
							}
						} catch (Exception e) {
							log.error(e.getMessage(), e);
						}
						break;

					case Write:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
							claimLoaderWait();
							WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName,
									rowNo, colNo);
						} else {
							// log.info("Not needed");
						}
						break;

					case CA: // Select all and clear the value, case added for claims ticket iTDM-911
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
							webElement.click();
							Thread.sleep(200);
							webElement.sendKeys(Keys.chord(Keys.CONTROL, "a"), ctrlValue);
							// webElement.clear();
							// webElement.sendKeys(ctrlValue);
						}
						break;

					case I:
						if (!ctrlValue.equalsIgnoreCase("null") || !(ctrlValue.trim().equalsIgnoreCase(""))) {
							scroll(controlName, webElement);

							WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
									Duration.ofSeconds(60));
							if (controlName.contains("//input[@type='file']")) {
								claimLoaderWait();
								// webElement.clear();
								if (ctrlValue.contains("PICK_INPUT_DATA_FILEPATH")) {
									ctrlValue = ctrlValue.replace("PICK_INPUT_DATA_FILEPATH", Config.inputDataFilePath);
									ctrlValue = ctrlValue.replace("\\\\", "\\");
								}

								webElement.sendKeys(ctrlValue);
							} else {
								claimLoaderWait();
								// WaitForPageLoad.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
								WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
								WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));

								if (logicalName.contains("Date") && (ctrlValue.equalsIgnoreCase("Today")
										|| ctrlValue.equalsIgnoreCase("Future") || ctrlValue.equalsIgnoreCase("Past"))) // 
								{
									if (ctrlValue.equalsIgnoreCase("Today")) {
										DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
										Date date = new Date();
										String todayDate = dateFormat.format(date);
										log.info(todayDate);
										webElement.clear();
										// Thread.sleep(500);
										if (logicalName.contains("VendorManagemtBusiness")) {
											String vendorName = "Automation" + todayDate;
											webElement.sendKeys(vendorName);
										} else {
											webElement.sendKeys(todayDate);
											// webElement.sendKeys(Keys.TAB);
											// log.info("Performed TAB out");//Mandar 21/10/24
										}
									} else if (ctrlValue.equalsIgnoreCase("Future")) {
										DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
										Date date = new Date();
										Calendar c = Calendar.getInstance();
										c.setTime(date); // Now use today
										// date.
										c.add(Calendar.DATE, 1);
										date = c.getTime();
										String todayDate = dateFormat.format(date);
										log.info(todayDate);
										webElement.clear();
										// Thread.sleep(500);
										webElement.sendKeys(todayDate);
									} else {
										DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
										Date date = new Date();
										Calendar c = Calendar.getInstance();
										c.setTime(date); // Now use today
										// date.
										c.add(Calendar.DATE, -1);
										date = c.getTime();
										String todayDate = dateFormat.format(date);
										log.info(todayDate);
										webElement.clear();
										// Thread.sleep(500);
										webElement.sendKeys(todayDate);
									}
								} else if ((logicalName.contains("ClaimInput") || logicalName.contains("PolicyInput")
										|| logicalName.contains("$Input")) && ctrlValue.equalsIgnoreCase("Y")) // 
								{
									uniqueNumber = ReadFromExcel(ctrlValue);
									webElement.clear();
									webElement.sendKeys(uniqueNumber);
								} else if ((logicalName.contains("CheckNo") || logicalName.equalsIgnoreCase("Invoice"))
										&& ctrlValue.equalsIgnoreCase("Y")) // 
								{
									String SALTCHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890";
									StringBuilder salt = new StringBuilder();
									Random rnd = new Random();
									while (salt.length() < 7) { // length of
										// the
										// random
										// string.
										int index = (int) (rnd.nextFloat() * SALTCHARS.length());
										salt.append(SALTCHARS.charAt(index));
									}
									String saltStr = salt.toString();
									log.info(saltStr);
									webElement.clear();
									webElement.sendKeys(saltStr);
									if (TransactionType.contains("HO") && ((logicalName.contains("CheckNo")
											&& TransactionType.contains("LossPayment"))
											|| logicalName.equalsIgnoreCase("Invoice")))
										WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType,
												controlName, rowNo, colNo);
								} else if (logicalName.contains("Email") && ctrlValue.contains("$randomvalue")) {
									String REmail = AddRandomCodeGenerator(ctrlValue);
									// Thread.sleep(500);
									webElement.clear();
									webElement.sendKeys(REmail);
									WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType,
											controlName, rowNo, colNo);
								} else if (ctrlValue.contains("Today+") || ctrlValue.contains("TODAY+")) {
									int ctrlValueLength = ctrlValue.length();
									String PlusDays = ctrlValue.substring(6, ctrlValueLength);
									int PlusDaysValue = Integer.parseInt(PlusDays);

									LocalDate localDate = LocalDate.now();
									LocalDate Futuredate = localDate.plusDays(PlusDaysValue);
									log.info(Futuredate);
									DateTimeFormatter FOMATTER = DateTimeFormatter.ofPattern("MM/dd/yyyy");
									// String FutureDateFormatted =
									// Futuredate.format(DateTimeFormatter.ofPattern("MM/DD/YYYY"));
									String FutureDateFormatted = FOMATTER.format(Futuredate);
									log.info(FutureDateFormatted);
									webElement.clear();
									Thread.sleep(500);
									webElement.sendKeys(FutureDateFormatted);

								} else if (ctrlValue.equalsIgnoreCase("Clear")) {
									Thread.sleep(100);
									webElement.clear();
									Thread.sleep(100);
								}

								else if (logicalName.contains("PayeeCode")) {

									String SALTCHARS = "1234567890";
									StringBuilder salt = new StringBuilder();
									Random rnd = new Random();
									while (salt.length() < 9) {
										int index = (int) (rnd.nextFloat() * SALTCHARS.length());
										salt.append(SALTCHARS.charAt(index));
									}
									String saltStr = salt.toString();
									log.info(saltStr);
									webElement.clear();
									webElement.sendKeys("V" + saltStr);

								}

								else {
									log.info("ctrlValue is :" + ctrlValue);
									// Thread.sleep(500);
									webElement.clear();
									webElement.sendKeys(ctrlValue);
									Thread.sleep(100);
									// webElement.sendKeys(Keys.TAB);
								}
							}
						} else {
							webElement.clear();
						}

						break;
					case V:

						if (!(ctrlValue.equalsIgnoreCase(""))) {
							if (ctrlValue.contains("$ClaimNo$")) {
								@SuppressWarnings("unused")
								WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
										Duration.ofSeconds(60));

								uniqueNumber = ReadFromExcel(ctrlValue);
								String ExpectedValue = ctrlValue.replace("$ClaimNo$", uniqueNumber);
								WebElement WebElement = WebHelper.wait
										.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));

								log.info(WebElement);
								String ActualValue = WebElement.getText();
								log.info("ActualValue is : " + ActualValue);
								log.info("ExpectedValue is : " + ExpectedValue);

								if (ExpectedValue.equalsIgnoreCase(ActualValue)) {
									log.info("Expected Value got match ");
									report.setStatus("PASS");
									report.setStatus(report.getStatus());
									report.setMessage("Values Matched");
								} else {
									log.info("Expected Value not match ");
									report.setStatus("FAIL");
									report.setStatus(report.getStatus());
									report.setMessage("Values Not Matched");
									WebHelperUtil.saveScreenShot();
									controller.pauseFun("");
								}

								break;

							} else if (ctrlValue.contains("FALSE") || ctrlValue.contains("TRUE")) {
								claimLoaderWait();
								Reporter report = new Reporter();
								report.setReport(report);

								WebElement Verifyelement1 = null;
								Verifyelement1 = WebHelper.currentdriver.findElement(By.xpath(controlName));
								currentValue = Verifyelement1.getAttribute("value");
								if (currentValue.isEmpty())
									currentValue = "FALSE";
								else
									currentValue = "TRUE";

								String ActualValue = currentValue;
								String ExpectedValue = ctrlValue;

								log.info("ActualValue is : " + ActualValue);
								log.info("ExpectedValue is : " + ExpectedValue);

								if (ExpectedValue.equalsIgnoreCase(ActualValue)) {
									log.info("Expected Value got match ");
									report.setStatus("PASS");
									report.setStatus(report.getStatus());
									report.setMessage("Values Matched");
								} else {
									log.info("Expected Value not match ");
									report.setStatus("FAIL");
									report.setStatus(report.getStatus());
									report.setMessage("Values Not Matched");
									WebHelperUtil.saveScreenShot();
									controller.pauseFun("");
								}
								break;
							} else {
								claimLoaderWait();
								WebElement Verifyelement = null;
								Verifyelement = WebHelper.wait
										.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
								// break;

								currentValue = Verifyelement.getText();
								Verifyelement = null;
							}
							break;
						}

						else {
							log.info("Element not Present");
						}
						break;

					case V_TextImage:
						if (!(ctrlValue.equalsIgnoreCase(""))) {
							@SuppressWarnings("unused")
							WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
									Duration.ofSeconds(60));
							WebElement Verifyelement = null;
							Verifyelement = WebHelper.currentdriver.findElement(By.xpath(controlName));
							currentValue = Verifyelement.getAttribute("src");
							Verifyelement = null;
							if (currentValue.equals(ctrlValue)) {
								log.info("The text matched with the ctrl value provided");
							}
						}

						else {
							log.info("Element not Present");
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
						String attribute = "", vType = "";
						String[] tempImageType = imageType.split(";;");
						if (!imageType.isEmpty() && tempImageType.length == 2) {
							attribute = tempImageType[1];
							vType = tempImageType[0];
						} else {
							attribute = "value";
							vType = "full";
						}

						try {
							if (attribute.equalsIgnoreCase("value")) {
								currentValue = webElement.getAttribute(attribute);
								if (StringUtils.equalsIgnoreCase(currentValue, null)
										|| currentValue.equalsIgnoreCase(""))
									currentValue = webElement.getText();
							} else {
								currentValue = webElement.getAttribute(attribute);
							}
							if (vType.equalsIgnoreCase("partial")) {
								if (currentValue.contains(ctrlValue)) {
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

					case VF:
						if (ctrlValue.contains("$ClaimNo$")) {
							@SuppressWarnings("unused")
							WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
									Duration.ofSeconds(60));

							uniqueNumber = ReadFromExcel(ctrlValue);
							String ExpectedValue = ctrlValue.replace("$ClaimNo$", uniqueNumber);
							WebElement WebElement = WebHelper.wait
									.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));

							log.info(WebElement);
							currentValue = WebElement.getText();
							log.info("ActualValue is : " + currentValue);
							log.info("ExpectedValue is : " + ExpectedValue);
							webDriver.setReport(
									WebHelperUtil.WriteToDetailResults(ExpectedValue, currentValue, logicalName));
							break;

						} else if (!(ctrlValue.equalsIgnoreCase("") || ctrlValue.contains("$ClaimNo$"))) {
							WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
									Duration.ofSeconds(60));
							WaitForPageLoad
									.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
							WebElement Verifyelement = null;
							Verifyelement = WebHelper.wait
									.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
							Verifyelement = WebHelper.currentdriver.findElement(By.xpath(controlName));
							// break;

							currentValue = Verifyelement.getText();

							Verifyelement = null;
						}
						break;

					case VV:
						if (!(ctrlValue.equalsIgnoreCase(""))) {
							Reporter report = new Reporter();
							report.setReport(report);
							if (!(ctrlValue.contains("/"))) {
								String ExpectedValue = ctrlValue;
								String ActualValue = webElement.getText();

								log.info("ActualValue is : " + ActualValue);
								log.info("ExpectedValue is : " + ExpectedValue);

								if (ActualValue.contains(ExpectedValue)) {
									report.setStatus("PASS");
									report.setStatus(report.getStatus());
									report.setMessage("Values Matched");
									// WebHelper.saveScreenShot();
								} else {
									report.setStatus("FAIL");
									report.setStatus(report.getStatus());
									// report.setMessage("Values Not Matched");
									report.setMessage("Actual Value =>" + " " + ActualValue + " "
											+ "Not Matching Expected Value =>" + " " + ExpectedValue);
									WebHelperUtil.saveScreenShot();
									controller.pauseFun(report.getMessage());
								}
								WebHelper.columns.add("");
								WebHelper.columnsData.add(WebHelper.columns);
								int temprowcount = 0;
								int tempcolcount = 0;
								ExcelUtility.WriteToCompareDetailResults(testCase, TransactionType, WebHelper.columns,
										WebHelper.columnsData, temprowcount, tempcolcount, report, ExpectedValue,
										ActualValue, logicalName, operationType, cycleDate);

							} else {
								String ExpectedValue1 = ctrlValue;
								String[] EParts = ExpectedValue1.split("/");
								String Expected = EParts[0];
								String EData = EParts[1];
								// int EDataNumber = Integer.parseInt(EData) ;

								String ActualValue1 = webElement.getText();
								String[] Actual = ActualValue1.split(" ");
								String[] parts = EData.split(",");
								String ActualValue = "";

								try {
									for (String P : parts) {
										ActualValue = ActualValue + Actual[Integer.parseInt(P)];
									}
								} catch (Exception ex) {
									log.info("Array Exception");
								}

								String ExpectedValue = Expected.replaceAll("\\s+", "");

								log.info("ActualValue is : " + ActualValue);
								log.info("ExpectedValue is : " + ExpectedValue);

								if (ActualValue.contains(ExpectedValue)) {
									report.setStatus("PASS");
									report.setStatus(report.getStatus());
									report.setMessage("Values Matched");
									// WebHelper.saveScreenShot();
								} else {
									report.setStatus("FAIL");
									report.setStatus(report.getStatus());
									report.setMessage("Values Not Matched");
									WebHelperUtil.saveScreenShot();
									// controller.pauseFun("");
								}
								WebHelper.columns.add("");
								WebHelper.columnsData.add(WebHelper.columns);
								int temprowcount = 0;
								int tempcolcount = 0;
								ExcelUtility.WriteToCompareDetailResults(testCase, TransactionType, WebHelper.columns,
										WebHelper.columnsData, temprowcount, tempcolcount, report, ExpectedValue,
										ActualValue, logicalName, operationType, cycleDate);
							}
						}
						break;
					default:
						break;
					}
					break;

				case WebButton:
					switch (actionName) {
					case I:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| ctrlValue.equalsIgnoreCase("YY") || !(ctrlValue.trim().equalsIgnoreCase(""))) {

							claimLoaderWait();
							Thread.sleep(500);
							scroll(controlName, webElement);
							// Thread.sleep(1000);

							if (Automation.browserType.toString().toUpperCase().contains("INTERNETEXPLORER")
									|| Automation.browserType.toString().toUpperCase().contains("MsEdge")) {
								log.info("IE");
								WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
										Duration.ofSeconds(90));
								WaitForPageLoad
										.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
								WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
								WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
								webElement.click();
							} else {
								if (!TransactionType.contains("HO") && (TransactionType.contains("Payment")
										&& (logicalName.equalsIgnoreCase("Preview")
												|| logicalName.equalsIgnoreCase("BulkPaymentSubmit")
												|| logicalName.equalsIgnoreCase("ReleaseBulk")
												|| logicalName.equalsIgnoreCase("CloseBatch")))) {
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(90));
									// ClaimLoader.claimLoaderWait();
									WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
									WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
									webElement.click();
									try {
										Thread.sleep(1000);
										claimLoaderWait();
										WebDriverWait WaitForPageLoadDP = new WebDriverWait(Automation.driver,
												Duration.ofSeconds(7));
										WebElement webElement1 = WaitForPageLoadDP
												.until(ExpectedConditions.elementToBeClickable(By.xpath(
														"//div[@data-labelkey='mm.icd.ClaimModification.Common.DuplicatePaymentFound']//div[@data-cid='btnContinue']")));
										claimLoaderWait();
										webElement1.click();

										Thread.sleep(1000);
										log.info("Pop-up found on the web page" + "ctrlValue");
									} catch (Exception e) {
										// log.error(e.getMessage(), e);
										log.info("No Duplicate-Payment pop up found");
									}
								} else if ((logicalName.equalsIgnoreCase("Login")
										|| logicalName.equalsIgnoreCase("Home"))) // - Handling Chat-bot
								{
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(90));
									// ClaimLoader.claimLoaderWait();
									WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
									WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
									webElement.click();

									try {
										Thread.sleep(1000);
										claimLoaderWait();
										WebDriverWait WaitForPageLoadWhtsNew = new WebDriverWait(Automation.driver,
												Duration.ofSeconds(15));

										WebElement WhatsNewClose = WaitForPageLoadWhtsNew
												.until(ExpectedConditions.elementToBeClickable(By.xpath(
														"//div[@class='modal-dialog']//button[@class='close']")));
										WhatsNewClose.click();

									} catch (Exception ex1) {
										log.info("No What's New Pop up found");
									}
									try { // Handling Chat-Bot
										Thread.sleep(2000);
										claimLoaderWait();
										WebDriverWait WaitForPageLoadChatBot = new WebDriverWait(Automation.driver,
												Duration.ofSeconds(10));
										;
										WebElement webElementChatBot = WaitForPageLoadChatBot.until(ExpectedConditions
												.elementToBeClickable(By.xpath("//a[@id='floatBtn']//img")));
										webElementChatBot.click();

										Thread.sleep(1000);
										WebElement webElementChatBotIFrame = WaitForPageLoadChatBot
												.until(ExpectedConditions.presenceOfElementLocated(By.xpath(
														"//div[@data-cid='mlChatWrapper']//iframe[contains(@name,'ml_frame')]")));
										Automation.driver.switchTo().frame(webElementChatBotIFrame);

										WebElement webElementChatBotDock = WaitForPageLoadChatBot
												.until(ExpectedConditions.elementToBeClickable(
														By.xpath("//div[@class='dockCheckBox']//label")));
										// webElementChatBotDock.click();
										((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
												webElementChatBotDock);

										Automation.driver.switchTo().defaultContent();
									} catch (Exception ex) {
										log.info("No Chat-Bot found");
									}
								}

								else if (logicalName.equalsIgnoreCase("LoginWhatsNewPopup")) {
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(90));
									// Thread.sleep(1000);
									WaitForPageLoad.until(
											ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
									WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
									WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
									log.info("Login/Home Click");
									webElement.click();

									try {

										WaitForPageLoad.until(ExpectedConditions
												.invisibilityOfElementLocated(By.xpath(claimsloader)));
										WebDriverWait WaitForPageLoadChatBot = new WebDriverWait(Automation.driver,
												Duration.ofSeconds(30));
										// WebElement webElementChatBot =
										// WaitForPageLoadChatBot.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@id='floatBtn']//img")));
										@SuppressWarnings("unused")
										WebElement webElementChatBot = WaitForPageLoadChatBot
												.until(ExpectedConditions.elementToBeClickable(
														By.xpath("//iframe[contains(@src,'loadwhatsnewpopup')]")));
										Thread.sleep(100);
										log.info("WhatsNewPopup Found");
										// Thread.sleep(1000);
										Automation.driver.switchTo().frame(Automation.driver
												.findElement(By.xpath("//iframe[contains(@src,'loadwhatsnewpopup')]")));
										// Thread.sleep(1000);

										WebElement WhatsNewPopupNew = Automation.driver
												.findElement(By.xpath("//button[@id='wn-skip-btn']"));
										WhatsNewPopupNew.click();

										Thread.sleep(1000);
										WebElement WhatsNewPopupOK = Automation.driver
												.findElement(By.xpath("//button[@id='wn-confirm-btn']"));
										WhatsNewPopupOK.click();
										Thread.sleep(500);
										Automation.driver.switchTo().defaultContent();
									} catch (Exception ex) {
										log.info("No WhatsNewPopup found");
									}
								}

								else if ((logicalName.contains("ClaimInput") || logicalName.contains("PolicyInput"))
										&& ctrlValue.equalsIgnoreCase("Y")) {
									uniqueNumber = ReadFromExcel(ctrlValue);
									WebElement webElement2 = WebHelper.wait
											.until(ExpectedConditions.elementToBeClickable(By
													.xpath(controlName + "[contains(text(),'" + uniqueNumber + "')]")));
									log.info(webElement2);
									@SuppressWarnings("unused")
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(120));
									// ClaimLoader.claimLoaderWait();
									webElement2.click();

								} else if (ctrlValue.equalsIgnoreCase("YY")) // -double Click
								{
									webElement.click();
									Thread.sleep(400);
									webElement.click();
									// builderdoubleClick.doubleClick(webElement).perform();

								} else if ((logicalName.equalsIgnoreCase("Last") && ctrlValue.equalsIgnoreCase("Y"))) // )
								{
									try {
										log.info("Thread.sleep(500)");
										Thread.sleep(500);
										WebElement ActualElement = WebHelper.currentdriver
												.findElement(By.xpath(controlName));
										WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
												Duration.ofSeconds(90));
										// ClaimLoader.claimLoaderWait();
										WaitForPageLoad.until(ExpectedConditions.visibilityOf(ActualElement));
										Thread.sleep(100);
										WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(ActualElement));
										Thread.sleep(500);
										((JavascriptExecutor) WebHelper.currentdriver)
												.executeScript("arguments[0].scrollIntoView();", ActualElement);
										((JavascriptExecutor) WebHelper.currentdriver)
												.executeScript("arguments[0].click();", ActualElement);

										// ActualElement.click();
										Thread.sleep(500);
										log.info("Last page clicked");
										// break;
									} catch (Exception e) {
										log.info("Last page button not enabled or does not exist - check Xpath");
										// log.error(e.getMessage(), e);
										break;
									}

								} else if ((logicalName.equalsIgnoreCase("PaginationLast"))
										&& ctrlValue.equalsIgnoreCase("Y")) {
									try {
										log.info("Thread.sleep(500)");
										Thread.sleep(500);

										WebElement ActualElement = WebHelper.currentdriver
												.findElement(By.xpath(controlName));

										log.info(ActualElement);

										if (ActualElement.isEnabled()) {
											log.info("Element is enabled");

											((JavascriptExecutor) WebHelper.currentdriver)
													.executeScript("arguments[0].scrollIntoView();", ActualElement);
											((JavascriptExecutor) WebHelper.currentdriver)
													.executeScript("arguments[0].click();", ActualElement);

											log.info("Last page button was clicked ");

										} else {
											log.info("Last page button was disabled");
											break;

										}

									} catch (Exception e) {
										log.info("Last page button not enabled or doesnot exist -check Xpath");
										// log.error(e.getMessage(), e);
										break;
									}

								} else {

									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(90));
									// ClaimLoader.claimLoaderWait();
									WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
									WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
									Thread.sleep(500);
									webElement.click();
									Thread.sleep(500);

								}
							}
						}
						break;

					case NC:// Empty case added for User Profile transaction - Autouser creation -
							// 11\07\2023
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| ctrlValue.equalsIgnoreCase("")) {

							((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].click();",
									webElement);

						}
						break;

					case C:

						WebDriverWait WaitForPageLoad1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(90));
						claimLoaderWait();
						WaitForPageLoad1.until(ExpectedConditions.visibilityOf(webElement));
						WaitForPageLoad1.until(ExpectedConditions.elementToBeClickable(webElement));
						webElement.click();

						break;

					case ReadD:
						if (ctrlValue.equalsIgnoreCase("Y")) {
							uniqueNumber = ReadFromExcel(ctrlValue);
							String controlNameNew = controlName.replace("$Value", uniqueNumber);
							WebElement webElementNew = WebHelper.wait
									.until(ExpectedConditions.elementToBeClickable(By.xpath(controlNameNew)));

							log.info(webElementNew);
							WebDriverWait WaitForPageLoadNew = new WebDriverWait(Automation.driver,
									Duration.ofSeconds(120));
							WaitForPageLoadNew
									.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
							webElementNew.click();
							Thread.sleep(1000);
						} else {
							log.info("Control Value found Blank or N ");
							break;
						}
						break;

					case V:
						WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(5));
						claimLoaderWait();
						WebElement Verifyelement = null;
						try {
							Verifyelement = WaitForPageLoad
									.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
							if (Verifyelement.isDisplayed()) {
								if (Verifyelement.isEnabled() == true)
									currentValue = "True";
								else
									currentValue = "False";
							}
						} catch (Exception e) {
							log.error(e.getMessage(), e);
							currentValue = "False";
						}
						break;

					case V_Disable:
						WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(5));
						WaitForPageLoad.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
						Verifyelement = null;
						try {
							WebElementList = WebHelper.currentdriver.findElements(By.xpath(controlName));
							if (!ctrlValue.equalsIgnoreCase("")) {
								if (WebElementList.isEmpty()) {

									currentValue = "Disable";
								}
								/*
								 * else { currentValue = "Enable"; }
								 */
							}

						} catch (Exception e) {
							log.error(e.getMessage(), e);
							currentValue = "Enable";
						}
						break;

					case Read:
						Reporter report = new Reporter();
						report.setReport(report);

						WebHelper.ActualValue = webElement.getText();
						uniqueNumber = ReadFromExcel(ctrlValue);
						WebHelper.ExpectedValue = uniqueNumber;

						log.info("ActualValue is : " + WebHelper.ActualValue);
						log.info("ExpectedValue is : " + WebHelper.ExpectedValue);
						// if(ActualValue.equalsIgnoreCase(ExpectedValue))

						if (WebHelper.ActualValue.contains(WebHelper.ExpectedValue)) {
							report.setStatus("PASS");
							report.setStatus(report.getStatus());
							report.setMessage("Values Matched");
							// WebHelper.saveScreenShot();
						} else {
							report.setStatus("FAIL");
							report.setStatus(report.getStatus());
							report.setMessage("Values Not Matched");
						}
						WebHelper.columns.add("");
						WebHelper.columnsData.add(WebHelper.columns);
						int temprowcount = 0;
						int tempcolcount = 0;
						ExcelUtility.WriteToCompareDetailResults(testCase, TransactionType, WebHelper.columns,
								WebHelper.columnsData, temprowcount, tempcolcount, report, WebHelper.ExpectedValue,
								WebHelper.ActualValue, logicalName, operationType, cycleDate);
						/*
						 * } else { log.info("Element is not clicked"); }
						 */
						break;
					case T: // 
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
							// Thread.sleep(1200);
							WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(90));
							claimLoaderWait();
							// WaitForPageLoad.until(ExpectedConditions.invisibilityOfElementLocated(By.className("overlay
							// hide")));
							WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
							WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
							webElement.click();
						} else {
							log.info("Element is not clicked");
						}
						break;

					case NCIF_Button:// Sel4-26/07/24

						List<String> browserList1 = Arrays
								.asList(new String[] { "InternetExplorer", "Chrome", "MsEdge", "Firefox", "Safari" });
						// for Button
						try {
							if (Automation.driver.findElement(By.xpath(controlName)).isDisplayed()) {
								List<WebElement> ls = WebHelper.currentdriver.findElements(By.xpath(controlName));
								System.out.println(ls.size());
								if (ls.size() > 0) {
									if (browserList1.contains(Automation.browserType.toString())) {
										((JavascriptExecutor) WebHelper.currentdriver).executeScript(
												"arguments[0].click();",
												Automation.driver.findElement(By.xpath(controlName)));
										log.info("controlName: " + controlName + "clicked");
									} else {
										((JavascriptExecutor) WebHelper.currentdriver).executeScript(
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

					case NCIF: // to be used with XPath_if controlID
						try {
							if (webElement.isDisplayed()) {
								webElement.click();
								Thread.sleep(500);
							}

						} catch (Exception e) {
							// log.error(e.getMessage(), e);
							log.info("Element Not Found on Page in NCIF case");
						}
						break;
					default:
						break;
					}
					break;

				case AjaxWebButton:
					switch (actionName) {
					case I:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
							if (ctrlValue.equalsIgnoreCase("Y")) {
								((JavascriptExecutor) WebHelper.currentdriver)
										.executeScript("arguments[0].scrollIntoView();", webElement);
								((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].click();",
										webElement);
							} else {
								((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].click();",
										webElement);
							}
						}
						break;
					case V:
						if (!(ctrlValue.equalsIgnoreCase(""))) {
							claimLoaderWait();
							WebElement Verifyelement = null;

							Verifyelement = WebHelper.wait
									.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
							// break;

							currentValue = Verifyelement.getAttribute("value");
							Verifyelement = null;
						} else {
							log.info("Element not Present");
						}
						break;
					default:
						break;
					}
					break;

				case WebElement:
					switch (actionName) {
					case NC: // Scroll and Click on JAvaScript
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
							uniqueNumber = ReadFromExcel(ctrlValue);

							scroll(controlName, webElement);

							webElement.clear();
							webElement.sendKeys(uniqueNumber);
						} else {
							log.info("Element is not Clicked");
						}
						break;

					/*
					 * case I: if (ctrlValue.equalsIgnoreCase("Y") ||
					 * ctrlValue.equalsIgnoreCase("Yes") ||
					 * !(ctrlValue.trim().equalsIgnoreCase(""))) { WebDriverWait WaitForPageLoad
					 * =new WebDriverWait(WebHelper.currentdriver, Duration.ofSeconds(90));
					 * claimLoaderWait(); scroll(controlName, webElement);
					 * WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
					 * WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
					 * 
					 * // Attempt clicking the element up to 3 times int clickAttempts = 0; boolean
					 * clicked = false; while (clickAttempts < 3 && !clicked) { try { // Scroll
					 * element into view ((JavascriptExecutor)
					 * WebHelper.currentdriver).executeScript("arguments[0].scrollIntoView();",
					 * webElement); // Click the element webElement.click(); clicked = true; // Set
					 * clicked to true if click succeeds System.out.println("Click attempt " +
					 * (clickAttempts + 1) + " performed."); } catch (Exception e) { // Log the
					 * exception if click fails log.error("Failed to click element. Attempt: " +
					 * (clickAttempts + 1)); clickAttempts++; // Increment click attempts // You can
					 * optionally add a delay between attempts if needed // Thread.sleep(1000); //
					 * Delay for 1 second } }
					 * 
					 * // If clicking failed after 3 attempts, throw an exception if (!clicked) {
					 * throw new RuntimeException("Failed to click element after 3 attempts."); } }
					 * else { log.info("Element is not clicked"); } break;
					 */

					/*
					 * case I: if (ctrlValue.equalsIgnoreCase("Y") ||
					 * ctrlValue.equalsIgnoreCase("Yes") || !ctrlValue.trim().isEmpty()) {
					 * 
					 * log.info("Selenium4-Fall'24-Update1 version...");
					 * //!(ctrlValue.trim().equalsIgnoreCase(""))) WebDriverWait WaitForPageLoad =
					 * new WebDriverWait(WebHelper.currentdriver, Duration.ofSeconds(90));
					 * claimLoaderWait(); scroll(controlName, webElement); // Scroll element into
					 * view
					 * 
					 * 
					 * ((JavascriptExecutor) WebHelper.currentdriver).executeScript(
					 * "arguments[0].scrollIntoViewIfNeeded(true);", webElement); Thread.sleep(500);
					 * 
					 * 
					 * WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
					 * 
					 * // Get the current position of the element Object isElementInView =
					 * ((JavascriptExecutor) WebHelper.currentdriver)
					 * .executeScript("var elem = arguments[0],         " +
					 * "box = elem.getBoundingClientRect(), " + "elemTop = box.top, " +
					 * "elemBottom = box.bottom, " +
					 * "isVisible = (elemTop >= 0) && (elemBottom <= window.innerHeight); " +
					 * "return isVisible;", webElement);
					 * 
					 * // Check if the element is in view if (!(Boolean) isElementInView) { //
					 * Scroll the element into view if it's not visible
					 * log.info(" Scrolling the element:" +webElement+"- into view");
					 * 
					 * ((JavascriptExecutor) WebHelper.currentdriver)
					 * .executeScript("arguments[0].scrollIntoView();", webElement);
					 * Thread.sleep(500); }else {
					 * 
					 * log.info(" Not Scrolling the element:"
					 * +webElement+"- as its in the view already"); }
					 * 
					 * 
					 * 
					 * ((JavascriptExecutor) WebHelper.currentdriver).executeScript(
					 * "arguments[0].scrollIntoViewIfNeeded(true);", webElement);
					 * 
					 * 
					 * // Attempt clicking the element up to 3 times int clickAttempts = 0; boolean
					 * clicked = false; while (clickAttempts < 3 && !clicked) { try {
					 * 
					 * // Create an instance of Actions class Actions actions = new
					 * Actions(WebHelper.currentdriver);//Disabled - 21/10/24
					 * 
					 * 
					 * if (webElement.isEnabled()) { log.info("The element is enabled."); } else {
					 * log.info("The element is not enabled."); }
					 * 
					 * if (isElementPresent(webElement)) { Thread.sleep(1000); // Perform the click
					 * action using Actions class actions.moveToElement(webElement);//Mandar try {
					 * webElement.sendKeys(Keys.DOWN); } catch (ElementNotInteractableException e) {
					 * log.info("Element not interactable when sending DOWN key: " +
					 * e.getMessage()); // Optionally, you can decide to refresh the element or take
					 * another action here } actions.click(webElement).perform();////Disabled -
					 * 21/10/24
					 * 
					 * ((JavascriptExecutor)
					 * WebHelper.currentdriver).executeScript("arguments[0].click();", webElement);
					 * Thread.sleep(1000);
					 * 
					 * //webElement.click(); //((JavascriptExecutor)
					 * WebHelper.currentdriver).executeScript("arguments[0].click();", webElement);
					 * clicked = true; // Set clicked to true if click succeeds
					 * log.info("Click attempt " + (clickAttempts + 1) + " performed.");
					 * 
					 * } else { log.info("Element is stale. Attempting to refresh."); webElement =
					 * WebHelper.currentdriver.findElement(By.xpath(controlName)); // Refresh
					 * element } }catch (StaleElementReferenceException e) { // Handle
					 * StaleElementReferenceException by refreshing the element reference
					 * log.info("Stale element reference detected. Refreshing element reference.");
					 * webElement = WebHelper.currentdriver.findElement(By.xpath(controlName)); }
					 * catch (Exception e) { // Log other exceptions log.info("Click attempt " +
					 * (clickAttempts + 1) + " failed due to: " + e.getMessage()); }
					 * clickAttempts++; // Increment click attempts }
					 * 
					 * // If clicking failed after 3 attempts, throw an exception if (!clicked) {
					 * throw new RuntimeException("Failed to click element after 3 attempts."); } }
					 * else { log.info("Element is not clicked"); }
					 * 
					 * break;
					 */
					case I:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes") || !ctrlValue.trim().isEmpty()) {
						    log.info("Selenium4-Fall'24-Update3  version...");
						    //WebDriverWait wait = new WebDriverWait(WebHelper.currentdriver, Duration.ofSeconds(90));
						    claimLoaderWait();
						    
						    // Scroll the element into view
						    log.info("BSP-61096 - Issue test, scrollIntoView added at this step for webelement, commented focus() - 23042025");
						    ((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].scrollIntoView(true);", webElement);
						    //((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].focus();", webElement);
						    //((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].parentNode.scrollTop += arguments[0].getBoundingClientRect().top - window.innerHeight/2;", webElement);
						     Thread.sleep(1000);
						    ((JavascriptExecutor) WebHelper.currentdriver).executeScript("window.scrollBy(0, arguments[0].getBoundingClientRect().top - window.innerHeight/2);", webElement);

						    Thread.sleep(500); // Give time for scroll

						    // Wait for the element to be clickable
						    //wait.until(ExpectedConditions.elementToBeClickable(webElement));//Mandar

						    // Attempt clicking the element with fallback to JavaScript click
						    try {
						        Actions actions = new Actions(WebHelper.currentdriver);
						        actions.moveToElement(webElement).click().perform();
						        log.info("Element clicked using Actions.");
						    } catch (Exception e) {
						    	String errorMessage = (e.getMessage() != null && !e.getMessage().isEmpty()) ? e.getMessage(): "No error message available";
								String firstLine = errorMessage.split("\\r?\\n", 2)[0];
								//log.error("Error :" + firstLine);
						        //log.info("Actions click failed, attempting JavaScript click due to: " + e.getMessage());
								log.info("Actions click failed, attempting JavaScript click due to: " + firstLine);
						        ((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].click();", webElement);
						        log.info("Element clicked using JavaScript.");
						    }
						} else {
						    log.info("Element is not clicked");
						}
						break;


					case C:
						if (controlId.equalsIgnoreCase("XPath$Value$ClaimNo$") && ctrlValue.equalsIgnoreCase("")) {
							log.info("Blank Value Found");
						} else {
							String tempCtrlName = controlName;
							String uniqueNumber1 = ReadFromExcel(ctrlValue);
							String tempReplaceString = tempCtrlName.replace("$ClaimNo$", uniqueNumber1);

							String ExpectedValue = tempReplaceString.replace("$Value", ctrlValue);
							scroll(controlName, webElement);
							WebElement DynamicElement = WebHelper.wait
									.until(ExpectedConditions.elementToBeClickable(By.xpath(ExpectedValue)));
							Thread.sleep(1000);

							DynamicElement.click();
						}
						break;

					case Read:

						uniqueNumber = ReadFromExcel(ctrlValue);
						log.info("Claim no from unique no sheet is :" + uniqueNumber);

						if (ctrlValue.trim().equalsIgnoreCase("")) {
							boolean elementexists;
							if (elementexists = WebHelper.currentdriver
									.findElements(By.xpath(controlName + "[contains(text(),'" + uniqueNumber + "')]"))
									.size() > 0)
								;// Semicolon removed - 23/10/23
							{
								if (elementexists == true) {
									log.info("Claim no :" + uniqueNumber + " " + "exist on the last page");

									WebElement webElementLastPage = WebHelper.wait
											.until(ExpectedConditions.presenceOfElementLocated(By
													.xpath(controlName + "[contains(text(),'" + uniqueNumber + "')]")));
									// log.info(webElement2));
									claimLoaderWait();
									Thread.sleep(2000);
									scrollAndHighlight(controlName, webElementLastPage);// 20/10/23
									try {
										webElementLastPage.click();
										Thread.sleep(6000);// Post click page takes time to load
										WebHelperUtil.saveScreenShot();// 20/10/23 log.info("Claim no :" +" "+
																		// uniqueNumber + " " + "is Selected");
									} catch (Exception e) {
										// Handle the click error
										log.error("Error clicking the element: " + e.getMessage());
									}

								} else {

									WebElement preButton = null;
									int count = 0;
									while (preButton == null & count <= 3) {
										log.info("Previous button click attempt :" + ++count);
										try {

											WebElement prevbutton = WebHelper.currentdriver
													.findElement(By.xpath(".//li[@class='prev']//a"));

											((JavascriptExecutor) WebHelper.currentdriver)
													.executeScript("arguments[0].scrollIntoView();", prevbutton);
											// Thread.sleep(1000);
											prevbutton.click();
											// log.info("Previous button clicked");
											elementexists = WebHelper.currentdriver
													.findElements(By.xpath(
															controlName + "[contains(text(),'" + uniqueNumber + "')]"))
													.size() > 0;
											if (elementexists == true) {
												WebElement webElement2 = WebHelper.wait.until(
														ExpectedConditions.presenceOfElementLocated(By.xpath(controlName
																+ "[contains(text(),'" + uniqueNumber + "')]")));
												log.info(webElement2);
												claimLoaderWait();
												Thread.sleep(2000);
												scrollAndHighlight(controlName, webElement2);
												webElement2.click();
												Thread.sleep(5000);
												WebHelperUtil.saveScreenShot();// 20/10/23
												log.info("Claim Id selected: " + uniqueNumber);
												break;
											}
											log.info("Claim Id not found on this page");
											WebHelperUtil.saveScreenShot();
										}

										catch (Exception ex) {
											log.info("Claim Id does not exist:" + ex.getMessage());
										}
									}

									// log.info("Claim Id does not exist on last 4 pages");
								}
							}
						} else {
							WebHelperUtil.saveScreenShot();

						}
						break;

					case Write:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
							WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName,
									rowNo, colNo);
						} else {
							log.info("Not needed");
						}
						break;

					case V_date:
						if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
								|| ctrlValue.equalsIgnoreCase("IGNORE")) {
							break;
						}
						currentValue = webElement.getAttribute("value");
						if (StringUtils.equalsIgnoreCase(currentValue, null) || currentValue.equalsIgnoreCase("")) {
							currentValue = webElement.getText();
							if (currentValue == null) {
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

					case V_text:
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

					case VFormCount:
						if (ctrlValue == null || ctrlValue.trim().equals("")) {
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

					case VLookupValuesPresence:
						if (!ctrlValue.equals("")) {
							expectedLookupValues.add(ctrlValue);
							if (!lookupValuesList.isEmpty()) {
								boolean lookupValueFound = false;
								for (int k = 0; k < lookupValuesList.size(); k++) {
									if (lookupValuesList.get(k).equals(ctrlValue)) {
										lookupValueFound = true;
										break;
									}
								}
								if (lookupValueFound == true) {
									presentLookupValues.add(ctrlValue);
									currentValue = ctrlValue;
								} else {
									missingLookupValues.add(ctrlValue);
									currentValue = "Lookup value is not present";
								}

							} else {
								ctrlValue = "";
								currentValue = "No lookup value is displayed on screen.";
							}
						}
						break;

					case FormVerificationDetails:
						if (!ctrlValue.equals("")) {
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
						}
						missingForms.removeAll(missingForms);
						additionalForms.removeAll(additionalForms);
						presentForms.removeAll(presentForms);
						formList.removeAll(formList);
						expectedForms.removeAll(expectedForms);

						break;

					case LookupVerificationDetails:
						if (!ctrlValue.trim().equalsIgnoreCase("")) {
							String temp3 = "", temp4 = "";
							boolean fail2 = false;

							// List missing LookupValues
							if (!missingLookupValues.isEmpty()) {
								fail2 = true;
								temp3 = missingLookupValues.size() + " lookup value(s) are missing.\n";
								for (int p = 0; p < missingLookupValues.size(); p++) {
									temp3 = temp3 + missingLookupValues.get(p) + "\n";
								}
							} else {
								temp3 = "No lookup value is missing.";
							}

							// List additional LookupValues
							List<String> tempLookupValues = new ArrayList<String>();
							// copy LookupValueList to tempLookupValues
							for (int k = 0; k < lookupValuesList.size(); k++) {
								tempLookupValues.add(lookupValuesList.get(k));
							}
							tempLookupValues.removeAll(expectedLookupValues);
							// fill additionalLookupValues
							for (int k = 0; k < tempLookupValues.size(); k++) {
								additionalLookupValues.add(tempLookupValues.get(k));
							}
							if (!additionalLookupValues.isEmpty()) {
								fail2 = true;
								temp4 = additionalLookupValues.size() + " lookup value(s) are additional.\n";
								for (int p = 0; p < additionalLookupValues.size(); p++) {
									temp4 = temp4 + additionalLookupValues.get(p) + "\n";
								}
							} else {
								temp4 = "No additional lookup value present.";
							}
							if (fail2 == true) {
								ctrlValue = temp3;
								currentValue = "\n" + temp4;
							} else {
								ctrlValue = temp3 + "\n" + temp4;
								currentValue = ctrlValue;
							}
						}
						missingLookupValues.removeAll(missingLookupValues);
						additionalLookupValues.removeAll(additionalLookupValues);
						presentLookupValues.removeAll(presentLookupValues);
						lookupValuesList.removeAll(lookupValuesList);
						expectedLookupValues.removeAll(expectedLookupValues);
						break;

					case VLookupValuesCount:
						if (ctrlValue == null || ctrlValue.trim().equals("")) {
							break;
						}
						if (!lookupValuesList.isEmpty()) {
							// check duplicate lookupValue instances
							List<String> templookupValues = new ArrayList<String>();
							List<String> duplicatelookupValues = new ArrayList<String>();
							String duplicatelookupValueStatus = "";
							// copy lookupValueList to templookupValues
							for (int k = 0; k < lookupValuesList.size(); k++) {
								templookupValues.add(lookupValuesList.get(k));
							}

							for (int j = 0; j < templookupValues.size(); j++) {
								for (int l = 1; l < templookupValues.size(); l++) {
									if (templookupValues.get(j).equals(templookupValues.get(l)) && (j != l)) {
										boolean presentInDuplicate = false;
										for (int k = 0; k < duplicatelookupValues.size(); k++) {
											if (duplicatelookupValues.get(k).equals(templookupValues.get(j))) {
												presentInDuplicate = true;
											}
										}
										if (presentInDuplicate == false) {
											duplicatelookupValues.add(templookupValues.get(j));
										}
									}
								}
							}
							if (duplicatelookupValues.isEmpty()) {
								duplicatelookupValueStatus = "No duplicate lookupValue present.";
							} else {
								duplicatelookupValueStatus = "Multiple instances of " + duplicatelookupValues.size()
										+ " lookupValues present.";
								for (int k = 0; k < duplicatelookupValues.size(); k++) {
									duplicatelookupValueStatus = duplicatelookupValueStatus + "\n"
											+ duplicatelookupValues.get(k);
								}
							}

							// compare expected and actual lookupValue count
							if (lookupValuesList.size() == Integer.parseInt(ctrlValue)) {
								ctrlValue = "Expected and Actual Lookup Values Count Matches. Count = " + ctrlValue
										+ ". ";
								ctrlValue = ctrlValue + "\n" + duplicatelookupValueStatus;
								currentValue = ctrlValue;
							} else {
								currentValue = "Mismatch in Lookup Values count. Expected Count = " + ctrlValue;
								ctrlValue = "Actual Count = " + lookupValuesList.size();
								ctrlValue = ctrlValue + "\n" + duplicatelookupValueStatus;
							}
						} else {
							if (ctrlValue.equalsIgnoreCase("0")) {
								ctrlValue = "No lookup value is displayed on screen. Expected lookup value Count = "
										+ ctrlValue;
								currentValue = ctrlValue;
							} else {
								ctrlValue = "Expected lookup value Count = " + ctrlValue;
								currentValue = "No lookup value is displayed on screen.";
							}
						}
						break;

					case V:// _ Verification of Text-fields
						if (!(ctrlValue.equalsIgnoreCase(""))) {

							if (!(ctrlValue.equalsIgnoreCase(""))) {

								if (ctrlValue.contains("<>")) // Pradeep: Added code to get Dynamic XPath value from
																// Inputdata to verify table data.12/26/2019
								{
									String[] MultipleControlValue = ctrlValue.split("<>");

									String tempControlname1 = controlName.replace("<>Value", MultipleControlValue[0]);

									Thread.sleep(500);
									WebDriverWait WaitForPageLoad1 = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(60));
									WaitForPageLoad1.until(
											ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
									WebElement Verifyelement = null;
									Verifyelement = WebHelper.wait
											.until(ExpectedConditions.elementToBeClickable(By.xpath(tempControlname1)));
									// currentValue = Verifyelement.getAttribute(tempControlname1.toString());
									currentValue = Verifyelement.getText();
									log.info("ActualValue is : " + currentValue);
									log.info("ExpectedValue is : " + MultipleControlValue[1]);
									// Verifyelement = null;
									ctrlValue = MultipleControlValue[1];

								}

								// Pradeep - Added case to verify randaom value starting with "M" , "Null" and
								// Random Numbers, TextVerify, Todays date, Count
								else if (ctrlValue.equalsIgnoreCase("M%") || ctrlValue.equalsIgnoreCase("Null")
										|| ctrlValue.equalsIgnoreCase("Number") || logicalName.contains("TextVerify")
										|| ctrlValue.equalsIgnoreCase("Today") || logicalName.contains("Count")) {
									if (ctrlValue.equalsIgnoreCase("M%")) {
										// WebElement Verifyelement = null;
										// Verifyelement =
										// WebHelper.wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
										WebElement CurrentElement = WebHelper.currentdriver
												.findElement(By.xpath(controlName));
										currentValue = CurrentElement.getText();
										char FirstChar = currentValue.charAt(0);
										String ActualValue = Character.toString(FirstChar);

										if (ActualValue.equalsIgnoreCase("M")) {
											currentValue = "M%";
										} else {
											currentValue = "NotStartedWithM";
										}
										break;

									} else if (ctrlValue.contains("Null")) {
										WebElement CurrentElement = WebHelper.currentdriver
												.findElement(By.xpath(controlName));
										String currentValue1 = CurrentElement.getAttribute("Value");

										if (currentValue1 == null) {
											currentValue = "Null";
										} else {
											currentValue = "Not Null";
										}
										break;
									}

									else if (ctrlValue.equalsIgnoreCase("Number")) {
										WebElement Verifyelement = null;
										Verifyelement = WebHelper.wait
												.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
										String ActualValue1 = Verifyelement.getText();
										boolean ActualValue2 = ActualValue1.chars().allMatch(Character::isDigit);

										if (ActualValue2 == true) {
											currentValue = "Number";
										} else {
											currentValue = "Not a Number";
										}
										break;
									}

									else if (ctrlValue.equalsIgnoreCase("Today")) {
										DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
										Date date = new Date();
										String ExpectedDate = dateFormat.format(date);
										log.info(ExpectedDate);

										WebElement AutualElement = WebHelper.wait
												.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
										// String ActualDate = AutualElement.getAttribute("value");
										String ActualDate = AutualElement.getText();

										if (ExpectedDate.equalsIgnoreCase(ActualDate)) {

											log.info("Date Matched");
											currentValue = "Today";
										} else {
											log.info("Date Not Matched");
											currentValue = ActualDate;
										}
										break;

									}

									else if (logicalName.contains("TextVerify")) {
										WebElement ActualElement = WebHelper.currentdriver
												.findElement(By.xpath(controlName));
										// Modified by Shubhashree-Uncommented to fetch the attribute property
										String ActualValue1 = ActualElement.getAttribute("value");
										String ActualValue = ActualElement.getText();
										if (ActualValue.contains(ctrlValue)) {
											log.info("Text Matched");
											currentValue = ctrlValue;
										}
										// Shubhashree-- Addedthe extra condition to handle the attribute scenario
										else if (ActualValue1.equals(ctrlValue)) {
											log.info("The text matches with the orginal value");
											currentValue = ctrlValue;
										} else {
											log.info("Text Not Matched");
											currentValue = ActualValue;
										}
										break;
									}

									else if (logicalName.contains("Count")) // Pradeep- Added case to verify the "Count"
																			// of the element present
									{
										List<WebElement> CurrentElement = WebHelper.currentdriver
												.findElements(By.xpath(controlName));
										// String currentValue1 = CurrentElement.getAttribute("Value");
										int ElementCount = CurrentElement.size();
										int ExpectedCount = Integer.parseInt(ctrlValue);
										// int ActualCount1=ActualCount;
										if (ElementCount == ExpectedCount) {
											currentValue = String.valueOf(ExpectedCount);
										} else {
											currentValue = String.valueOf(ElementCount);
										}
										break;

									}

								} else {
									claimLoaderWait();
									WebElement Verifyelement = null;
									Verifyelement = WebHelper.wait
											.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
									currentValue = Verifyelement.getAttribute(controlName.toString());
									Verifyelement = null;
								}
							} else {
								log.info("Element not Present");
							}
						}
						break;

					case V_ReadOnly:

						try {
							if (ctrlValue.equalsIgnoreCase("ReadOnly")) {
								WebElement webElement1 = WebHelper.currentdriver.findElement(By.xpath(controlName));
								String Verifyelement1 = webElement1.getAttribute("readonly");
								if (Verifyelement1 != null && Verifyelement1.equalsIgnoreCase("true")) {
									currentValue = "ReadOnly";
								} else {
									currentValue = "Enable";
								}
							} else {
								log.info("No control value to be verified");
							}
						} catch (Exception e) {
							log.error(e.getMessage(), e);
							currentValue = "False";
						}
						// bhaskar Action Verify Values VV START
					case T: // _Claims (Payment I Icon
						// Verification)
						Reporter report = new Reporter();
						report.setReport(report);
						WebDriverWait WaitForPageLoad11 = new WebDriverWait(Automation.driver, Duration.ofSeconds(5));
						// Thread.sleep(1000);
						WebElement logout = WaitForPageLoad11.until(ExpectedConditions.elementToBeClickable(By.xpath(
								"//div[@class='v-label v-widget v-has-width' and contains(text(),'Pending Approval')]")));
						logout.click();

						String ActualValue1 = webElement.getText();
						String ExpectedValue1 = ctrlValue;

						String[] parts = ActualValue1.split(" ");
						String ActualValue = parts[0] + parts[1] + parts[2];

						String ExpectedValue = ExpectedValue1.replaceAll("\\s+", "");

						log.info("ActualValue is : " + ActualValue);
						log.info("ExpectedValue is : " + ExpectedValue);

						if (ActualValue.contains(ExpectedValue)) {
							report.setStatus("PASS");
							report.setStatus(report.getStatus());
							report.setMessage("Values Matched");
							// WebHelper.saveScreenShot();
						} else {
							report.setStatus("FAIL");
							report.setStatus(report.getStatus());
							report.setMessage("Values Not Matched");
							WebHelperUtil.saveScreenShot();
							controller.pauseFun("");
						}
						WebHelper.columns.add("");
						WebHelper.columnsData.add(WebHelper.columns);
						int temprowcount = 0;
						int tempcolcount = 0;
						ExcelUtility.WriteToCompareDetailResults(testCase, TransactionType, WebHelper.columns,
								WebHelper.columnsData, temprowcount, tempcolcount, report, ExpectedValue, ActualValue,
								logicalName, operationType, cycleDate);
						break;
					case V_NotBlank: // Pradeep - Added case to verify empty field or field with value.12/26/2019.
						try {
							if (ctrlValue.equalsIgnoreCase("NotBlank")) {
								WebElement webElement1 = WebHelper.currentdriver.findElement(By.xpath(controlName));
								String Verifyelement1 = webElement1.getAttribute("value");
								if (Verifyelement1.isEmpty() == true) {
									currentValue = "NotBlank";
								} else {
									currentValue = "Blank";
								}
							} else {
								log.info("No control value to be verified");
							}
						} catch (Exception e) {
							log.error(e.getMessage(), e);
							currentValue = "False";
						}
						break;

					case V_ClaimNo: // Pradeep - Added case to verify ClaimNo.12/26/2019.
						try {
							if (!ctrlValue.equalsIgnoreCase("")) {
								WebElement webElement1 = WebHelper.currentdriver.findElement(By.xpath(controlName));
								String Verifyelement1 = webElement1.getAttribute("value");
								String uniqueNumber1 = ReadFromExcel(ctrlValue);
								if (Verifyelement1.equalsIgnoreCase(uniqueNumber1) == true) {
									currentValue = "True";
								} else {
									currentValue = "False";
								}
							} else {
								log.info("No value to be verified");
							}
						} catch (Exception e) {
							log.error(e.getMessage(), e);
							currentValue = "False";
						}
						break;

					default:
						break;
					}
					break;

				// Below given code added to support AddUSerUtility for Claims - 07\07\2023

				case LevelSelect:
					switch (actionName) {
					case I:

						@SuppressWarnings("unused")
						WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(7));
						// WebDriverWait = new WebDriverWait(Automation.driver, 15L);
						uniqueNumber = ReadFromExcel(ctrlValue);
						uniqueNumber = uniqueNumber.toLowerCase();
						if (uniqueNumber.contains("supervisor")) {
							uniqueNumber = "Level3";
						} else if (uniqueNumber.contains("manager")) {
							uniqueNumber = "Level2";
						} else if (uniqueNumber.contains("vp")) {
							uniqueNumber = "Level1";
						}
						claimLoaderWait();

						/*
						 * WebElement webElement1 = (WebElement) WebHelper.wait .until((Function)
						 * ExpectedConditions.presenceOfElementLocated(By.xpath(
						 * String.valueOf(controlName) + "[contains(text(),'" + uniqueNumber + "')]")));
						 */

						WebElement webElement1 = (WebElement) WebHelper.wait
								.until(ExpectedConditions.presenceOfElementLocated(By.xpath(
										String.valueOf(controlName) + "[contains(text(),'" + uniqueNumber + "')]")));

						Thread.sleep(500L);
						webElement1.click();
						break;
					case T:
						uniqueNumber = ReadFromExcel(ctrlValue);
						uniqueNumber = uniqueNumber.concat("%");
						webElement.clear();
						webElement.sendKeys(new CharSequence[] { uniqueNumber.toLowerCase() });
						break;
					default:
						break;
					}
					break;

				case RoleAssign:
					switch (actionName) {
					case I:
						String[] parts;
						int j;
						String[] arrayOfString3;
						byte b;
						uniqueNumber = ReadFromExcel(ctrlValue);
						parts = uniqueNumber.split(",");
						for (j = (arrayOfString3 = parts).length, b = 0; b < j;) {
							String role = arrayOfString3[b];
							claimLoaderWait();
							WebDriverWait webDriverWait = new WebDriverWait(Automation.driver, Duration.ofSeconds(15));
							Thread.sleep(1000L);
							/*
							 * WebElement searchBox = (WebElement) webDriverWait.until((Function)
							 * ExpectedConditions
							 * .elementToBeClickable(By.xpath("(//input[contains(@class,'v-textfield')])[2]"
							 * )));
							 */ // Sel4
							WebElement searchBox = (WebElement) webDriverWait.until(ExpectedConditions
									.elementToBeClickable(By.xpath("(//input[contains(@class,'v-textfield')])[2]")));

							searchBox.clear();
							searchBox.sendKeys(new CharSequence[] { role });
							Thread.sleep(1000L);
							/*
							 * WebElement searchBoxImg = (WebElement) webDriverWait.until((Function)
							 * ExpectedConditions .elementToBeClickable(By.xpath(
							 * "(//img[contains(@src,'images/search.png')])[2]")));
							 */ // Sel4

							WebElement searchBoxImg = (WebElement) webDriverWait.until((ExpectedConditions
									.elementToBeClickable(By.xpath("(//img[contains(@src,'images/search.png')])[2]"))));

							searchBoxImg.click();
							Thread.sleep(1000L);

							/*
							 * WebElement Checkbox = (WebElement) webDriverWait .until((Function)
							 * ExpectedConditions.elementToBeClickable(By.xpath( "//div[contains(text(),'" +
							 * role + "')]/ancestor::td/..//td//input")));
							 */

							WebElement Checkbox = (WebElement) webDriverWait
									.until((ExpectedConditions.elementToBeClickable(By.xpath(
											"//div[contains(text(),'" + role + "')]/ancestor::td/..//td//input"))));

							Checkbox.click();
							Thread.sleep(1000L);
							searchBox.clear();
							Thread.sleep(1000L);
							searchBoxImg.click();
							b++;
						}
						break;
					default:
						break;
					}
					break;

				// End.

				case JSScript:
					((JavascriptExecutor) WebHelper.currentdriver).executeScript(controlName, ctrlValue);
					break;

				case WaitForElementToVisible: // swapnil41444 06/30/2021: Added try-catch block
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
							log.info("WebElement not found: " + controlName);
							Thread.sleep(1000);
							x--;
						}
					}

					if (x == 0)
						log.info("WebElement not found finally: " + controlName);

					break;

				case WaitForPageToLoad:
					switch (actionName) {
					case I:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !ctrlValue.trim().equals("")) {

							WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
									Duration.ofSeconds(30));// Sel4
							WaitForPageLoad
									.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
							WaitForPageLoad.until(ExpectedConditions
									.visibilityOf(WebHelper.currentdriver.findElement(By.xpath(controlName))));
							WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
							Thread.sleep(1500);

						}
						break;
					case NC:
						// Automation.driver.manage().timeouts().pageLoadTimeout(120,
						// TimeUnit.SECONDS);//Sel4
						Automation.driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(30));
						WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(30));

						WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
						Thread.sleep(1500);
						break;
					default:
						break;

					}
					break;

				case Wait: // Implicit wait
					switch (actionName) {
					case NC:
						if (StringUtils.equalsIgnoreCase(Config.applyStaticWait, "false")) {
							log.info("Wait not applied");
						} else {
							Thread.sleep(Integer.parseInt(controlName) * 1000);
						}

						break;
					case I:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
							Thread.sleep(Integer.parseInt(controlName) * 1000);
						}
						break;
					default:
						break;
					}
					break;

				case Wait_IfValue:
					switch (actionName) {
					case I:
						if (!ctrlValue.isEmpty()) {
							Thread.sleep(Integer.parseInt(controlName) * 1000);
						}
						break;
					default:
						break;
					}
					break;

				case WaitToLoad:
					return WebHelper.doAction(null, null, null, null, controlType, controlId, controlName, ctrlValue,
							null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
							rowcount, rowNo, colNo, null, null, null);

				case CheckBox:
					switch (actionName) {
					case I:
						if (ctrlValue.equalsIgnoreCase("") || StringUtils.isEmpty(ctrlValue)) {
							break;
						} else if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !ctrlValue.equalsIgnoreCase(""))
						// Pradeep-Added non blank condition for
						// scenarios having some input value.
						{
							claimLoaderWait();
							if (!webElement.isSelected()) // Added By
							// Dharmendra
							// to Check whether
							// CheckBox select or
							// not
							{
								// Thread.sleep(1000);
								// ((JavascriptExecutor)
								// WebHelper.currentdriver).executeScript("arguments[0].scrollIntoView();",
								// webElement);
								scroll(controlName, webElement);
								webElement.click(); // Select Checkbox,
								// if it is
								// not selectd
								Thread.sleep(1000);
							}
						}

						else if (ctrlValue.equalsIgnoreCase("N") || ctrlValue.equalsIgnoreCase("No")) {
							// Thread.sleep(1000);
							if (webElement.isSelected()) // Added By
							// Dharmendra
							// to Check whether
							// CheckBox select or
							// not
							{
								// Thread.sleep(1000);
								// ((JavascriptExecutor)
								// WebHelper.currentdriver).executeScript("arguments[0].scrollIntoView();",
								// webElement);
								scroll(controlName, webElement);
								webElement.click(); // Deselect
								// Checkbox, if it
								// is selectd
								Thread.sleep(1000);
							}
						}
						break;
					case T:
						if (ctrlValue.equalsIgnoreCase("") || StringUtils.isEmpty(ctrlValue)) {
							break;
						} else if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
							if (!webElement.isSelected())

							{
								scroll(controlName, webElement);
								// Thread.sleep(1000);
								webElement.click(); // Select Checkbox,
							}
						}

						else if (ctrlValue.equalsIgnoreCase("N") || ctrlValue.equalsIgnoreCase("No")) {
							if (webElement.isSelected())

							{
								scroll(controlName, webElement);

								webElement.click(); // Deselect Checkbox, if it

							}
						}
						break;
					case V:
						/*
						 * if (webElement.isSelected()) { //
						 * log.info("logical Name is not Selected:"+logicalName); log.info(logicalName +
						 * "Is Selected"); // currentValue = //
						 * webElement.getAttribute(controlName.toString()); } else {
						 * log.info(logicalName + "Is not Selected"); }
						 */
						if (ctrlValue.contains("FALSE") || ctrlValue.contains("TRUE")) { // - Verify Checkbox is
																							// selcted or Not
							claimLoaderWait();
							Reporter report = new Reporter();
							report.setReport(report);

							WebElement Verifyelement1 = null;
							// Verifyelement1 =
							// WebHelper.wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
							Verifyelement1 = WebHelper.currentdriver.findElement(By.xpath(controlName));

							if (Verifyelement1.isSelected() == true)
								currentValue = "TRUE";
							else
								currentValue = "FALSE";

							String ActualValue = currentValue;
							String ExpectedValue = ctrlValue;

							log.info("ActualValue is : " + ActualValue);
							log.info("ExpectedValue is : " + ExpectedValue);

							if (ExpectedValue.equalsIgnoreCase(ActualValue)) {
								log.info("Expected Value got match ");
								report.setStatus("PASS");
								report.setStatus(report.getStatus());
								report.setMessage("Values Matched");
							} else {
								log.info("Expected Value not match ");
								report.setStatus("FAIL");
								report.setStatus(report.getStatus());
								report.setMessage("Values Not Matched");
								WebHelperUtil.saveScreenShot();
								controller.pauseFun("");
							}
						}
						break;

					case V_EnableDisable:// Pradeep- Added case to verify EnableDisable based on element attribute as
											// "disabled"
						if (ctrlValue.equalsIgnoreCase("") || StringUtils.isEmpty(ctrlValue)) {
							log.info("Blank Value Found");
							break;
						} else {
							String ActualValue = webElement.getAttribute("disabled");

							if (ActualValue.equalsIgnoreCase("true")) {
								// String ActualValue = webElement.getAttribute("disabled");
								log.info(logicalName + "  Is Disabled");
								log.info(ActualValue);
								currentValue = "Disabled";
							} else {
								log.info(logicalName + "Element is Enabled");
								currentValue = "Enabled";
							}
							// break;
						}
						break;

					case NC:
						if (!webElement.isSelected()) {
							((JavascriptExecutor) WebHelper.currentdriver)
									.executeScript("arguments[0].scrollIntoView();", webElement);
							webElement.click();
						}
						break;
					case Read: // _Claims for delete/void bulk
						uniqueNumber = ReadFromExcel(ctrlValue);
						try {
							WebElement webElement2 = WebHelper.wait.until(
									ExpectedConditions.elementToBeClickable(By.xpath(controlName + "[contains(text(),'"
											+ uniqueNumber + "')]//following::td[@data-colid='Col_7']/div/div/input")));
							log.info(webElement2);
							((JavascriptExecutor) WebHelper.currentdriver)
									.executeScript("arguments[0].scrollIntoView();", webElement2);
							webElement2.click();
							Thread.sleep(1000);
						} catch (Exception e) {
							log.error(e.getMessage(), e);
							WebElement webElement1 = WebHelper.wait.until(ExpectedConditions.elementToBeClickable(By
									.xpath("//div[@id='mainRegion']/div[@class='page-region']//div[@data-name='tabExpPayDetl_tab']//li[@class='last']/a")));
							webElement1.click();
							WebElement webElement2 = WebHelper.wait.until(
									ExpectedConditions.elementToBeClickable(By.xpath(controlName + "[contains(text(),'"
											+ uniqueNumber + "')]//following::td[@data-colid='Col_7']/div/div/input")));
							log.info(webElement2);
							webElement2.click();
							Thread.sleep(1000);
						}
						/*
						 * webElement.clear(); webElement.sendKeys(uniqueNumber);
						 */
						break;
					default:
						break;
					}
					break;

				case Radio:
					switch (actionName) {
					case I: // _Claims for release bulk
						if (ctrlValue.equalsIgnoreCase("Y") || !ctrlValue.equalsIgnoreCase("")) {
							uniqueNumber = ReadFromExcel(ctrlValue);
							claimLoaderWait();
							if (ctrlValue.equalsIgnoreCase("Y")
									&& (logicalName.contains("ReleaseClaimInputCheckBox"))) {
								try {
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(7));
									WebElement webElement2 = WaitForPageLoad
											.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName
													+ "[contains(text(),'" + uniqueNumber
													+ "')]//ancestor::div[contains(@id,'OpenBatchData')]//following::td[@data-colid]//input[@type='checkbox']")));
									log.info(webElement2);
									// Thread.sleep(1000);
									((JavascriptExecutor) WebHelper.currentdriver)
											.executeScript("arguments[0].scrollIntoView();", webElement2);
									// Thread.sleep(1000);
									webElement2.click();
								} catch (Exception e) {
									log.error(e.getMessage(), e);
									WebElement webElement1 = WebHelper.wait
											.until(ExpectedConditions.elementToBeClickable(By.xpath(
													"//div[@id='mainRegion']/div[@class='page-region']//div[@data-name='tabExpPayDetl_tab']//li[@class='last']/a")));
									((JavascriptExecutor) WebHelper.currentdriver)
											.executeScript("arguments[0].scrollIntoView();", webElement1);
									webElement1.click();
									WebElement webElement2 = WebHelper.wait
											.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName
													+ "[contains(text(),'" + uniqueNumber
													+ "')]//ancestor::div[contains(@id,'OpenBatchData')]//following::td[@data-colid]//input[@type='checkbox']")));
									log.info(webElement2);
									// Thread.sleep(1000);
									((JavascriptExecutor) WebHelper.currentdriver)
											.executeScript("arguments[0].scrollIntoView();", webElement2);
									Thread.sleep(500);
									webElement2.click();
								}
							} else if (ctrlValue.equalsIgnoreCase("Y")
									&& (logicalName.contains("DeleteClaimBulkPayment"))) {
								try {
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(7));
									WebElement webElement2 = WaitForPageLoad
											.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName
													+ "[contains(text(),'" + uniqueNumber
													+ "')]//ancestor::div[contains(@id,'OpenBatchData')]//following::td[@data-colid]//img[contains(@src,'delete')]")));
									log.info(webElement2);
									// Thread.sleep(1000);
									((JavascriptExecutor) WebHelper.currentdriver)
											.executeScript("arguments[0].scrollIntoView();", webElement2);
									// Thread.sleep(1000);
									webElement2.click();
								} catch (Exception e) {
									log.error(e.getMessage(), e);
									WebElement webElement1 = WebHelper.wait
											.until(ExpectedConditions.elementToBeClickable(By.xpath(
													"//div[@id='mainRegion']/div[@class='page-region']//div[@data-name='tabExpPayDetl_tab']//li[@class='last']/a")));
									((JavascriptExecutor) WebHelper.currentdriver)
											.executeScript("arguments[0].scrollIntoView();", webElement1);
									webElement1.click();
									WebElement webElement2 = WebHelper.wait
											.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName
													+ "[contains(text(),'" + uniqueNumber
													+ "')]//ancestor::div[contains(@id,'OpenBatchData')]//following::td[@data-colid]//img[contains(@src,'delete')]")));
									log.info(webElement2);
									// Thread.sleep(1000);
									((JavascriptExecutor) WebHelper.currentdriver)
											.executeScript("arguments[0].scrollIntoView();", webElement2);
									// Thread.sleep(1000);
									webElement2.click();
								}
							} else if (ctrlValue.equalsIgnoreCase("Y")
									&& (logicalName.contains("EditClaimBulkPayment"))) {
								try {
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(7));
									WebElement webElement2 = WaitForPageLoad
											.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName
													+ "[contains(text(),'" + uniqueNumber
													+ "')]//ancestor::div[contains(@id,'OpenBatchData')]//following::td[@data-colid]//img[contains(@src,'edit')]")));
									log.info(webElement2);
									Thread.sleep(1000);
									((JavascriptExecutor) WebHelper.currentdriver)
											.executeScript("arguments[0].scrollIntoView();", webElement2);
									// Thread.sleep(1000);
									webElement2.click();
								} catch (Exception e) {
									log.error(e.getMessage(), e);
									WebElement webElement1 = WebHelper.wait
											.until(ExpectedConditions.elementToBeClickable(By.xpath(
													"//div[@id='mainRegion']/div[@class='page-region']//div[@data-name='tabExpPayDetl_tab']//li[@class='last']/a")));
									((JavascriptExecutor) WebHelper.currentdriver)
											.executeScript("arguments[0].scrollIntoView();", webElement1);
									webElement1.click();
									WebElement webElement2 = WebHelper.wait
											.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName
													+ "[contains(text(),'" + uniqueNumber
													+ "')]//ancestor::div[contains(@id,'OpenBatchData')]//following::td[@data-colid]//img[contains(@src,'edit')]")));
									log.info(webElement2);
									// Thread.sleep(1000);
									((JavascriptExecutor) WebHelper.currentdriver)
											.executeScript("arguments[0].scrollIntoView();", webElement2);
									// Thread.sleep(1000);
									webElement2.click();
								}
							}
						}
						break;

					case V:
					case F:
						return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
								controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
								webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo,
								operationType, cycleDate, TransactionType);

					case NC: // - To verify radio or check box
						// selected or
						// not
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
							if (!webElement.isSelected()) {
								log.info("IS NOT SELECTED");
								log.info(logicalName + " is not Selected ");
								// webElement.click();
							} else {
								log.info("IS SELECTED");
								log.info(logicalName + " is Selected ");
							}
						} else {
							log.info("No action CHECKBOX/ RADIO Button");
						}
						break;
					case T: // _Select depending on the
						// requirement
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
							if (!webElement.isSelected()) {
								webElement.click();
							} else {
								log.info("No radio found");
								// currentValue =
								// webElement.getAttribute(controlName.toString());
							}
						} else {
							log.info("NO for this Scenario");
						}
						break;
					default:
						break;
					}
					break;

				case WebLink:
					switch (actionName) {
					case I:
						if (ctrlValue.equalsIgnoreCase("B")) {
							// Suit Automation(WIP-Unassigned Claims)
							Reporter report = new Reporter();
							report.setReport(report);
							WebDriverWait WaitForPageLoad11 = new WebDriverWait(Automation.driver,
									Duration.ofSeconds(10));
							;
							// Thread.sleep(1000);
							WebElement WIPUnssignedClaims = WaitForPageLoad11
									.until(ExpectedConditions.elementToBeClickable(By.xpath(
											"//div[@class='v-label v-widget v-has-width' and contains(text(),'Unassigned Claims')]")));
							// logout.click();
							String ActualValue1 = WIPUnssignedClaims.getText();
							String[] parts = ActualValue1.split(" ");
							String ActualValue = parts[0];

							WIPUnssignedClaims.click();

							WebElement ExpectedValue1 = WaitForPageLoad11.until(ExpectedConditions.elementToBeClickable(
									By.xpath("//div[@class='totalrec']//label[@data-section='record_count']")));
							String ExpText = ExpectedValue1.getText();
							String[] ExpTextparts = ExpText.split(" ");
							String ExpectedValue = ExpTextparts[4];

							log.info("ActualValue is : " + ActualValue);
							log.info("ExpectedValue is : " + ExpectedValue);

							if (ActualValue.contains(ExpectedValue)) {
								report.setStatus("PASS");
								report.setStatus(report.getStatus());
								report.setMessage("Values Matched");
								// WebHelper.saveScreenShot();
							} else {
								report.setStatus("FAIL");
								report.setStatus(report.getStatus());
								report.setMessage("Values Not Matched");
								WebHelperUtil.saveScreenShot();
								// MainController.pauseFun("");
							}
							WebHelper.columns.add("");
							WebHelper.columnsData.add(WebHelper.columns);
							int temprowcount = 0;
							int tempcolcount = 0;
							ExcelUtility.WriteToCompareDetailResults(testCase, TransactionType, WebHelper.columns,
									WebHelper.columnsData, temprowcount, tempcolcount, report, ExpectedValue,
									ActualValue, logicalName, operationType, cycleDate);
							break;
						} else if (ctrlValue.equalsIgnoreCase("A")) {
							// Suit Automation(WIP-Pending Approval)
							Reporter report = new Reporter();
							report.setReport(report);
							WebDriverWait WaitForPageLoad11 = new WebDriverWait(Automation.driver,
									Duration.ofSeconds(10));
							;
							// Thread.sleep(1000);
							WebElement WIPPendingApproval = WaitForPageLoad11
									.until(ExpectedConditions.elementToBeClickable(By.xpath(
											"//div[@class='v-label v-widget v-has-width' and contains(text(),'Pending Approval')]")));
							// logout.click();
							String ActualValue1 = WIPPendingApproval.getText();
							String[] parts = ActualValue1.split(" ");
							String ActualValue = parts[0];

							WIPPendingApproval.click();
							WebElement PendingForMyApprovalClick = WaitForPageLoad11
									.until(ExpectedConditions.elementToBeClickable(
											By.xpath("//a[@data-labelkey='mm.icd.Approval.PendingForApprovals']")));
							PendingForMyApprovalClick.click();

							WebElement ReserveString = WaitForPageLoad11.until(ExpectedConditions.elementToBeClickable(
									By.xpath("//div[@id='dspPendingReserve']//span[@data-section='displayvalue']")));
							String ExpText1 = ReserveString.getText();
							String[] ExpText1parts = ExpText1.split(" ");
							String ExpPart1 = ExpText1parts[1];
							int beginInd = ExpPart1.indexOf("(");
							int endInd = ExpPart1.indexOf(")");
							String SumText1 = ExpPart1.substring(beginInd + 1, endInd);
							int SumNum1 = Integer.parseInt(SumText1);

							WebElement PaymentsString = WaitForPageLoad11.until(ExpectedConditions.elementToBeClickable(
									By.xpath("//div[@id='dspPendingPayments']//span[@data-section='displayvalue']")));
							String ExpText2 = PaymentsString.getText();
							String[] ExpText2parts = ExpText2.split(" ");
							String ExpPart2 = ExpText2parts[1];
							int beginInd2 = ExpPart2.indexOf("(");
							int endInd2 = ExpPart2.indexOf(")");
							String SumText2 = ExpPart2.substring(beginInd2 + 1, endInd2);
							int SumNum2 = Integer.parseInt(SumText2);

							int Sum = SumNum1 + SumNum2;
							String ExpectedValue = String.valueOf(Sum);

							log.info("ActualValue is : " + ActualValue);
							log.info("ExpectedValue is : " + ExpectedValue);

							if (ActualValue.contains(ExpectedValue)) {
								report.setStatus("PASS");
								report.setStatus(report.getStatus());
								report.setMessage("Values Matched");
								// WebHelper.saveScreenShot();
							} else {
								report.setStatus("FAIL");
								report.setStatus(report.getStatus());
								report.setMessage("Values Not Matched");
								WebHelperUtil.saveScreenShot();
								// MainController.pauseFun("");
							}
							WebHelper.columns.add("");
							WebHelper.columnsData.add(WebHelper.columns);
							int temprowcount = 0;
							int tempcolcount = 0;
							ExcelUtility.WriteToCompareDetailResults(testCase, TransactionType, WebHelper.columns,
									WebHelper.columnsData, temprowcount, tempcolcount, report, ExpectedValue,
									ActualValue, logicalName, operationType, cycleDate);
						} else if (ctrlValue.equalsIgnoreCase("C")) {
							{
								// Suit Automation(WIP-Pending FNOL)
								Reporter report1 = new Reporter();
								report1.setReport(report1);
								Thread.sleep(1000);
								WebDriverWait WaitForPageLoad12 = new WebDriverWait(Automation.driver,
										Duration.ofSeconds(10));
								;
								WebElement WIPPendingFNOL = WaitForPageLoad12
										.until(ExpectedConditions.elementToBeClickable(By.xpath(
												"//div[@class='v-label v-widget v-has-width' and contains(text(),'Pending FNOL')]")));
								// logout.click();
								String ActualValue2 = WIPPendingFNOL.getText();
								String[] PFParts = ActualValue2.split(" ");
								String PFActualValue = PFParts[0];

								WIPPendingFNOL.click();

								WebElement ExpectedValue2 = WaitForPageLoad12
										.until(ExpectedConditions.elementToBeClickable(By.xpath(
												"//div[@class='totalrec']//label[@data-section='record_count']")));
								String ExpText2 = ExpectedValue2.getText();
								String[] PFExpTextparts = ExpText2.split(" ");
								String PFExpectedValue = PFExpTextparts[4];

								log.info("ActualValue is : " + PFActualValue);
								log.info("ExpectedValue is : " + PFExpectedValue);

								if (PFActualValue.contains(PFExpectedValue)) {
									report1.setStatus("PASS");
									report1.setStatus(report1.getStatus());
									report1.setMessage("Values Matched");
									// WebHelper.saveScreenShot();
								} else {
									report1.setStatus("FAIL");
									report1.setStatus(report1.getStatus());
									report1.setMessage("Values Not Matched");
									WebHelperUtil.saveScreenShot();
									// MainController.pauseFun("");
								}
								WebHelper.columns.add("");
								WebHelper.columnsData.add(WebHelper.columns);
								int temprowcount = 0;
								int tempcolcount = 0;
								ExcelUtility.WriteToCompareDetailResults(testCase, TransactionType, WebHelper.columns,
										WebHelper.columnsData, temprowcount, tempcolcount, report1, PFExpectedValue,
										PFActualValue, logicalName, operationType, cycleDate);
								break;
							}
						}
					default:
						break;
					}
					break;

				case WindowAlertOk:
					// Pradeep-Added case from PAS to handle Alert popup
					switch (actionName) {
					case I:
						if (ctrlValue == null || ctrlValue.trim().equals("")) {
							break;
						}
						Automation.driver.switchTo().alert().accept();
						break;
					default:
						break;
					}
					break;
				case CloseWindow:// added this Case to bypass page loading
					// after clicking the event
				case WaitForJS:
				case ListBox:
				case WebList:
					return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
							controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
							webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo, operationType,
							cycleDate, TransactionType);

				case AjaxWebList:
					switch (actionName) {
					case I:
						// 10\07\23
						if (controlId.equalsIgnoreCase("AutoUserPath")) {
							if (logicalName.equalsIgnoreCase("AuthorityTypeSelect")) {
								WebElement element1 = Automation.driver
										.findElement(By.xpath("//span[contains(@id,'select2-cmbAuthorityType-')]"));

								String titleValue = element1.getAttribute("title");
								log.info("Title value (already selected value) is: " + titleValue);

								if (titleValue.equalsIgnoreCase(ctrlValue)) {
									log.info(ctrlValue + " is already selected, performing a click action...");

									WebElement newElement = Automation.driver
											.findElement(By.xpath("//span[contains(@id,'select2-cmbAuthorityType')"
													+ "and text()='" + ctrlValue + "']"));

									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(30));
									WaitForPageLoad.until(ExpectedConditions.visibilityOf(newElement));
									WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(newElement));
									Thread.sleep(500);
									newElement.click();

									// newElement.sendKeys(Keys.TAB);
									log.info("Click completed");
									break;

								} else {

									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(30));
									// claimLoaderWait();
									WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
									WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
									Thread.sleep(500);
									webElement.click();
									log.info("Authority Type Selected : " + ctrlValue);
									break;
								}

							} else {
								WebElement element = Automation.driver
										.findElement(By.xpath("//span[contains(@id,'select2-cmbAuthorityLevel-')]"));
								String titleValue = element.getAttribute("title");

								log.info("Title value (already selected value) is: " + titleValue);
								if (titleValue.equalsIgnoreCase(ctrlValue)) {
									log.info(ctrlValue + " is already selected, performing a click action...");

									WebElement newElement = Automation.driver
											.findElement(By.xpath("//span[contains(@id,'select2-cmbAuthorityLevel')"
													+ "and text()='" + ctrlValue + "']"));

									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(30));
									WaitForPageLoad.until(ExpectedConditions.visibilityOf(newElement));
									WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(newElement));
									Thread.sleep(500);
									newElement.click();

									// newElement.sendKeys(Keys.TAB);
									log.info("Click completed");
									break;

								} else {
									log.info("BSP-61096 ...1...- Issue test...");
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(30));
									// claimLoaderWait();
									WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
									WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
									Thread.sleep(500);
									webElement.click();
									log.info("Authority Level Selected : " + ctrlValue);
									break;
								}

							}
						} else {

							WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
									Duration.ofSeconds(30));
							log.info("BSP-61096 - Issue test, scroll added at this step fro Ajaxweblist - 23042025");
							claimLoaderWait();
							scroll(controlName, webElement);
							Thread.sleep(500);
							WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
							WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
							Thread.sleep(500);
							
							//((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].focus();", webElement);
						    //((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].parentNode.scrollTop += arguments[0].getBoundingClientRect().top - window.innerHeight/2;", webElement);
						    // Thread.sleep(1000);
							//((JavascriptExecutor) WebHelper.currentdriver).executeScript("window.scrollBy(0, arguments[0].getBoundingClientRect().top - window.innerHeight/2);", webElement);
							webElement.click();
							break;
						}

					case V_List:// Pradeep- Added case to verify
						// Dropdown value with expected list
						// Thread.sleep(2000);
						currentValue = new String();
						@SuppressWarnings("unused")
						String[] ExpectedValue;
						String string = ctrlValue;
						String[] ExpectedValue1 = string.split(",");
						for (int i = 1; i < ExpectedValue1.length + 5; i++) {
							String webElement1 = WebHelper.currentdriver.findElements(By.xpath(
									"//span[starts-with(@class,'select2-container select2-container--default select2-container--open')]//span[@class='select2-results']//ul[starts-with(@id,'select2-cmbVendorType-') and contains(@id,'results')]//li"))
									.get(i).getText();

							if (webElement1.equals(ExpectedValue1[i])) {
								log.info("MATCHED");
							}

							else {
								log.info("Value Not MATCHED");
							}
						}

						break;

					/*
					 * case V_Select :// Pradeep- Added case to verify // Dropdown value with
					 * expected list Thread.sleep(2000); if (!ctrlValue.equalsIgnoreCase("")) {
					 * String ExpectedValue_String = ctrlValue; String[] ExpectedValue2 =
					 * ExpectedValue_String.split(","); for (int i = 0; i < ExpectedValue2.length;
					 * i++) { Thread.sleep(1000); WebElement VendorSearchClick =
					 * WebHelper.currentdriver .findElement(By
					 * .xpath("//span[starts-with(@aria-labelledby,'select2-cmbVendorType-') and contains(@aria-labelledby,'container')]//span[@class='select2-selection__arrow']/b"
					 * )); VendorSearchClick.click(); Thread.sleep(1000); WebElement
					 * VendorSearchSelect = WebHelper.currentdriver.findElement(By.xpath(controlName
					 * + "[contains(text(),'" + ExpectedValue2[i] + "')]"));
					 * VendorSearchSelect.click(); // WebElement VendorSearch = //
					 * WebHelper.currentdriver.findElement(By.
					 * xpath("//span/input[@class='select2-search__field' and @type='search']")); //
					 * VendorSearch.sendKeys(ExpectedValue2[i]); // Thread.sleep(1000); //
					 * VendorSearch.click(); Thread.sleep(1000); WebElement webElement1 =
					 * WebHelper.currentdriver .findElement(By
					 * .xpath("//div[@class='col-md-3']/div[@style='display: block;']//span[starts-with(@id,'select2-cmbAssociatedOrganization-') and contains(@id,'container')]"
					 * ));
					 * 
					 * boolean AssoOrgDropDown = webElement1.isDisplayed(); if (AssoOrgDropDown ==
					 * true) { log.info("Associated Organization dropDown found for " +
					 * ExpectedValue2[i]); String Actualvalue = webElement1.getText();
					 * 
					 * if (Actualvalue.equals("Select")) { log.info("Value Select found"); }
					 * 
					 * else { log.info("Select not found");
					 * 
					 * } }
					 * 
					 * else { log.info("Associated Organization dropDown not found for " +
					 * ExpectedValue2[i]); } } } break;
					 */
					case V_Select:// Pradeep- Added case to verify
						// Dropdown value with expected list
						// Thread.sleep(2000);
						if (!ctrlValue.equalsIgnoreCase("")) {
							String ExpectedValue_String = ctrlValue;
							String[] ExpectedValue2 = ExpectedValue_String.split(",");
							for (int i = 0; i < ExpectedValue2.length; i++) {
								// Thread.sleep(1000);
								WebElement VendorSearchClick = WebHelper.currentdriver.findElement(By.xpath(
										"//span[starts-with(@aria-labelledby,'select2-cmbVendorType-') and contains(@aria-labelledby,'container')]//span[@class='select2-selection__arrow']/b"));
								VendorSearchClick.click();
								Thread.sleep(1000);
								WebElement VendorSearchSelect = WebHelper.currentdriver.findElement(
										By.xpath(controlName + "[contains(text(),'" + ExpectedValue2[i] + "')]"));
								VendorSearchSelect.click();
								// WebElement VendorSearch =
								// WebHelper.currentdriver.findElement(By.xpath("//span/input[@class='select2-search__field'
								// and @type='search']"));
								// VendorSearch.sendKeys(ExpectedValue2[i]);
								// Thread.sleep(1000);
								// VendorSearch.click();
								Thread.sleep(1000);
								WebElement webElement1 = WebHelper.currentdriver.findElement(By.xpath(
										"//div[@class='col-md-3']/div[@style='display: block;']//span[starts-with(@id,'select2-cmbAssociatedOrganization-') and contains(@id,'container')]"));

								boolean AssoOrgDropDown = webElement1.isDisplayed();
								if (AssoOrgDropDown == true) {
									System.out
											.println("Associated Organization dropDown found for " + ExpectedValue2[i]);
									String Actualvalue = webElement1.getText();

									if (Actualvalue.equals("Select")) {
										log.info("Value Select found");
										currentValue = Actualvalue;
										ctrlValue = ExpectedValue2[i];
									}

									else {
										log.info("Select not found");
										currentValue = Actualvalue;
										ctrlValue = ExpectedValue2[i];
									}

								}

								else {
									log.info("Associated Organization dropDown not found for " + ExpectedValue2[i]);
								}
							}
						}
						break;

					case V_Disable:// Pradeep- Added case to verify
						// Dropdown value with expected list
						// Thread.sleep(2000);
						if (!ctrlValue.equalsIgnoreCase("")) {
							String ExpectedValue1_String = ctrlValue;
							// Start-Change the logic to take value from controlValue
							String[] ExpectedValue_String = ExpectedValue1_String.split("<>");
							String[] ExpectedValue3 = ExpectedValue_String[1].split(",");
							// End
							for (int i = 0; i < ExpectedValue3.length; i++) {
								// Thread.sleep(1000);
								WebElement VendorSearchClick = WebHelper.currentdriver.findElement(By.xpath(
										"//span[starts-with(@aria-labelledby,'select2-cmbVendorType-') and contains(@aria-labelledby,'container')]//span[@class='select2-selection__arrow']/b"));
								VendorSearchClick.click();
								Thread.sleep(1000);
								WebElement VendorSearchSelect = WebHelper.currentdriver.findElement(
										By.xpath(controlName + "[starts-with(text(),'" + ExpectedValue3[i] + "')]"));
								VendorSearchSelect.click();
								// WebElement VendorSearch =
								// WebHelper.currentdriver.findElement(By.xpath("//span/input[@class='select2-search__field'
								// and @type='search']"));
								// VendorSearch.sendKeys(ExpectedValue2[i]);
								// Thread.sleep(1000);
								// VendorSearch.click();
								Thread.sleep(1000);
								WebElement webElement1 = WebHelper.currentdriver.findElement(By.xpath(
										"//div[contains(@class,'hide')]/div[@style='display: none;']//span[starts-with(@id,'select2-cmbAssociatedOrganization-') and contains(@id,'container')]"));

								boolean AssoOrgDropDown = webElement1.isDisplayed();

								if (AssoOrgDropDown == false) {
									log.info("Associated Organization dropDown not found for " + ExpectedValue3[i]);
									// Start-Added the verification point
									currentValue = ExpectedValue_String[0];
									ctrlValue = ExpectedValue_String[0];
									// end
									report.setStatus("PASS");
									report.setStatus(report.getStatus());
									report.setMessage("Values Matched");
									// WebHelper.saveScreenShot();

								}

								else {
									System.out
											.println("Associated Organization dropDown found for " + ExpectedValue3[i]);

									report.setStatus("FAIL");
									report.setStatus(report.getStatus());
									report.setMessage("Values Not Matched");
									WebHelperUtil.saveScreenShot();
									// MainController.pauseFun("");
								}
							}
						}
						break;

					case VA:
						return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
								controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
								webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo,
								operationType, cycleDate, TransactionType);

					case V: // Pradeep: Added case to verify selected value from drop down
						if (!ctrlValue.contains(",")) {
							WebElement webElement1 = WebHelper.currentdriver.findElement(By.xpath(controlName));
							currentValue = webElement1.getText();
							if (StringUtils.isEmpty(currentValue)) {
								currentValue = webElement1.getAttribute("value");
							}
						}
						break;
					default:
						break;
					}
					break;

				case ComboList: // Pradeep- To handle Combo Dropdown List
					// through by searching value, facing
					// issue while selecting value by
					// AjaxWebList
					switch (actionName) {
					case I:
						WebDriverWait WaitForPageLoad1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(30));
						WaitForPageLoad1.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
						WaitForPageLoad1.until(ExpectedConditions.visibilityOf(webElement));
						WaitForPageLoad1.until(ExpectedConditions.elementToBeClickable(webElement));
						Thread.sleep(200);
						webElement.sendKeys(ctrlValue);
						log.info("Element Entered");
						// webElement.sendKeys(Keys.ARROW_DOWN);
						Thread.sleep(200);
						webElement.sendKeys(Keys.ENTER);
						log.info("Enter key pressed");
						Thread.sleep(400);
						break;
					default:
						break;

					}
					break;

				case Dropdown: // Pradeep-Added case from PAS to Select
					// Drop down value.
					switch (actionName) {
					case I:
						if (ctrlValue == null || ctrlValue.trim().equals("")) {
							break;
						}
						Select s1 = new Select(webElement);
						s1.selectByVisibleText(ctrlValue);
						break;

					case Select:// Pradeep: Added case to select drop down value having "ul", "li" list in one
								// step.
						// it will reduce the additional steps
						// to click and then select value.12/26/2017

						WebElement dropdown = WebHelper.currentdriver.findElement(By.xpath(controlName));
						dropdown.click();
						if (controlName.equalsIgnoreCase("//div[@id='mainRegion']/div[@class='page-region']//span")) {
							String NewControlName = controlName.replace("//span", "//ul");
							List<WebElement> options = dropdown.findElements(By.xpath(NewControlName + "//li"));
							for (WebElement option : options) {
								if (option.getText().equals(ctrlValue)) {
									option.click(); // click the expected option
									break;
								}
							}
						} else if (controlName.equalsIgnoreCase("//span")) {

						}

					case V_List: // Pradeep-Added case to verify DropDown value through select.getOptions()
									// method.
						if (ctrlValue == null || ctrlValue.trim().equals("")) {
							break;
						}

						String[] ExpectedValue1 = ctrlValue.split("<>");
						@SuppressWarnings("unused")
						List<WebElement> AllWebElement = WebHelper.currentdriver.findElements(By.xpath(controlName));
						WebElement dropdown1 = WebHelper.currentdriver.findElement(By.xpath(controlName));

						Select select = new Select(dropdown1);

						List<WebElement> options = select.getOptions();
						int count = 0;
						for (WebElement we : options) {
							for (int i = 0; i < ExpectedValue1.length; i++) {

								if (we.getText().equals(ExpectedValue1[i])) {
									count++;
									// String addArray[]= (String) ExpectedValue1[i];
									log.info("Matched");
									log.info("Actual Value: " + we.getText());
									log.info("Expected Value: " + ExpectedValue1[i]);
								} else {
									log.info("Actual Fail Value: " + we.getText());
								}
							}
						}
						if (count == ExpectedValue1.length) {
							log.info("Expected Count: " + ExpectedValue1.length);
							log.info("Acutal Count: " + count);
							currentValue = "Matched";
							ctrlValue = "Matched";
							break;
						} else {
							log.info("Expected Count: " + ExpectedValue1.length);
							log.info("Acutal Count: " + count);
							currentValue = "Value Not Matched";
							ctrlValue = "Matched";
						}
						// break;
					default:
						break;

					}
					break;

				case TAB: // Pradeep- To handle TAB to focus out from
					// element
					switch (actionName) {
					case NC:
						log.info("Webelement is: " + webElement);

						webElement.sendKeys(Keys.TAB);
						log.info("Performed TAB out");
						break;

					case I: // _ TAB optional
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !ctrlValue.equalsIgnoreCase("")) {
							// log.info("Webelement is: " + webElement);

							webElement.sendKeys(Keys.TAB);
							log.info("Performed TAB out");

						}
						break;
					default:
						break;
					}
					break;

				case IFrame: // Pradeep-Added case from PAS
					switch (actionName) {
					case I:
						if (ctrlValue == null || ctrlValue.trim().equals("")) {
							break;
						} else if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| ctrlValue.equalsIgnoreCase("YES") || !ctrlValue.trim().equalsIgnoreCase("")) { // Pradeep-Added

							log.info("In method doaction debug4");
							Thread.sleep(500);

							if (controlName.startsWith("//iframe")) {

								WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(700));

								wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

								Automation.driver.switchTo()
										.frame(Automation.driver.findElement(By.xpath(controlName)));

							}

							else {

								Automation.driver.switchTo().frame(controlName);
							}

						}

						log.info("In method doaction debug5");
						break;

					case NC:
						log.info("In method doaction debug4");
						Thread.sleep(500);

						if (controlName.startsWith("//iframe")) {

							WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(700));

							wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

							Automation.driver.switchTo().frame(Automation.driver.findElement(By.xpath(controlName)));

						}

						else {

							Automation.driver.switchTo().frame(controlName);

						}

						log.info("In method doaction debug5");
						break;

					case Back: // Pradeep-Added case to handle
						// Switching back to parent window
						log.info("Switching back to Parent window");
						Thread.sleep(500);

						Automation.driver.switchTo().defaultContent();

						log.info("Done the Switching back to Parent window");
						break;
					default:
						break;

					}
					break;

				case Multiselect: // _ MultiSelect/Deselct for JS
					// (Ajax)
					// Claims
					switch (actionName) {
					case I:
						if (!ctrlValue.equalsIgnoreCase("null")) {

							if (!ctrlValue.contains(",")) {
								webElement = WebHelper.wait.until(ExpectedConditions.elementToBeClickable(
										By.xpath(controlName + "[contains(text(),'" + ctrlValue + "')]")));
								// ((JavascriptExecutor)currentdriver).executeScript("arguments[0].scrollIntoView();",
								// webElement);
								WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
										Duration.ofSeconds(30));
								claimLoaderWait();
								WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
								WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));

								webElement.click();
							}

							else {
								String string = ctrlValue;
								String[] parts = string.split(",");
								String partWebelemt = parts[0];
								for (int i = 1; i < parts.length; i++) {
									WebElement webElement1 = WebHelper.wait
											.until(ExpectedConditions.elementToBeClickable(
													By.xpath(controlName + "[contains(text(),'" + parts[i] + "')]")));
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(30));
									claimLoaderWait();
									webElement1.click();
									Thread.sleep(300); // Pradeep-Added
									// wait
									if (i < parts.length - 1) {
										webElement = WebHelper.wait
												.until(ExpectedConditions.elementToBeClickable(By.xpath(partWebelemt)));
										claimLoaderWait();
										WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
										WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
										webElement.click();
										Thread.sleep(300); // Pradeep-Added
										// wait
										/*
										 * Thread.sleep(1000); robot = new Robot(); robot.keyPress(KeyEvent .VK_ENTER);
										 * robot.keyRelease(KeyEvent .VK_ENTER); Thread.sleep(1000);
										 */
									}
								}
							}
						} else {
							log.info("No Data Entered");
						}
						break;

					case T: //  - Deselect multi choice
						if (!(ctrlValue.equalsIgnoreCase(""))) {
							/*
							 * if(!ctrlValue.contains(",")){ webElement= wait.until
							 * (ExpectedConditions.elementToBeClickable (
							 * By.xpath(controlName+"[@title='"+ctrlValue +
							 * "']//span[@class='select2-selection__choice__remove']" )));
							 * webElement.click(); } else{
							 */
							String string = ctrlValue;
							String[] parts = string.split(",");
							String partWebelemt = parts[0];
							for (int i = 1; i < parts.length; i++) {
								// webElement=
								// wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName+"[@title='"+parts[i]+"']//span[@class='select2-selection__choice__remove']")));
								webElement = WebHelper.wait.until(ExpectedConditions
										.elementToBeClickable(By.xpath(controlName + "[contains(@title,'" + parts[i]
												+ "')]//span[@class='select2-selection__choice__remove']")));
								WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
										Duration.ofSeconds(30));
								// Thread.sleep(1000);
								claimLoaderWait();
								WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
								WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
								webElement.click();
								Thread.sleep(1000);
								if (i == parts.length - 1) {
									webElement = WebHelper.wait
											.until(ExpectedConditions.elementToBeClickable(By.xpath(partWebelemt)));
									claimLoaderWait();
									WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
									WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
									webElement.click();
								}
							}
							// }

							/*
							 * robot = new Robot(); robot.keyPress(KeyEvent.VK_TAB);
							 * robot.keyRelease(KeyEvent.VK_TAB);
							 */
						} else {
							log.info("No Data Entered");
						}
					default:
						break;
					}
					break;

				case Browser:
					// Thread.sleep(3000); //DS:Check if required
					switch (actionName) {
					case I:
						Set<String> handlers = null;
						handlers = WebHelper.currentdriver.getWindowHandles();
						for (String handler : handlers) {
							WebHelper.currentdriver = WebHelper.currentdriver.switchTo().window(handler);

							// TM-19/01/2015: Changed following
							// comparison from
							// equalsIgnoreCase to contains
							if (WebHelper.currentdriver.getTitle().contains(controlName)) {
								log.info("Focus on window with title: " + WebHelper.currentdriver.getTitle());
								break;
							}
						}
						break;
					default:
						break;
					}
					break;

				case NewBrowser:// : Added 31-May :Start
					switch (actionName) {
					case I: // _Extra verification...
						if (!ctrlValue.trim().equalsIgnoreCase("")) {
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
						if (!ctrlValue.trim().equalsIgnoreCase("")) {
							String parentWindow1 = Automation.driver.getWindowHandle();
							Set<String> handles1 = Automation.driver.getWindowHandles();
							for (String windowHandle1 : handles1) {
								if (!windowHandle1.equals(parentWindow1)) {
									Automation.driver.switchTo().window(windowHandle1);
									// Thread.sleep(1000);
									Automation.driver.close();
								}
							}
							Automation.driver.switchTo().window(parentWindow1);
						}
						break;
					default:
						break;
					}
					break;// case NewBrowser://

					
				case NewWindow: 
				    switch (actionName) {
				        case I: 
				            if (!ctrlValue.trim().isEmpty()) { // Ensure ctrlValue is not empty
				                String parentWindow = Automation.driver.getWindowHandle(); // Get the parent window handle
				                Set<String> handles = Automation.driver.getWindowHandles(); // Get all open window handles

				                if (handles.size() > 1) { // Check if there are multiple windows
				                    for (String windowHandle : handles) {
				                        if (!windowHandle.equals(parentWindow)) { // Switch to non-parent window
				                            Automation.driver.switchTo().window(windowHandle);
				                           log.info("Switched to new window: " + Automation.driver.getTitle());
				                            break; // Exit loop once switched
				                        }
				                    }
				                } else {
				                	Automation.driver.switchTo().window(parentWindow);
				                	log.info("Only one window (parent) is open. Staying on parent window.");
				                }
				            }
				        default:
				            break;
				    }
				    break;
				
				// Below given case added to support new browser functionality in claims Note feature - 30/01/2025    
				case ParentWindow:
					switch (actionName) {

					case I:
						String parentWindow = null;
						if (!ctrlValue.trim().isEmpty()) { // Ensure ctrlValue is not empty
							// If parentWindow is not already stored, store it

							Set<String> handles = Automation.driver.getWindowHandles();
							// ✅ If parentWindow is not assigned, take the first (or only) available window
							// handle
							if (parentWindow == null && !handles.isEmpty()) {
								parentWindow = handles.iterator().next();
								log.info("Stored Parent Window Handle: " + parentWindow);
							}
							// Set<String> handles = Automation.driver.getWindowHandles(); // Get all open
							// window handles

							if (handles.size() > 1) { // In claims only 1 window exists
								log.info("Multiple window handle found");
								for (String windowHandle : handles) {
									if (!windowHandle.equals(parentWindow)) { // Switch to non-parent window
										Automation.driver.switchTo().window(windowHandle);
										log.info("Switched to new window: " + Automation.driver.getTitle());
										// Close the new window
										Automation.driver.close();
										log.info("New window closed.");
										break; // Exit loop after switching
									}
								}
								// Switch back to the parent window
								if (parentWindow != null) {
									Automation.driver.switchTo().window(parentWindow);
									System.out.println("Returned to parent window: " + Automation.driver.getTitle());
								} else {
									System.err.println("Parent window handle is not available!");
								}
							} else {
								log.info("Only one window (parent) is open. Staying on parent window.");
								for (String windowHandle : handles) {
									if (windowHandle.equals(parentWindow)) { // Switch to non-parent window
										Automation.driver.switchTo().window(windowHandle);
										log.info("Switched to parent window: " + Automation.driver.getTitle());
										// Close the new window
										//Automation.driver.close();
										//log.info("New window closed.");
										break; // Exit loop after switching
									}

								}
							}
						}
					default:
						break;
					}
					break;
				    
					
				case CloseBrowser:
					switch (actionName) {
					case I:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
							Automation.driver.close();
						}
						break;// case

					case NC:
						if (!ctrlValue.trim().equalsIgnoreCase("")) {
							Set<String> handles1 = Automation.driver.getWindowHandles();
							for (String windowHandle1 : handles1) {
								Automation.driver.switchTo().window(windowHandle1);
							}
						}
						break;

					default:
						break;
					}
					break;

				case CloseAndLaunchNewBrowser:
					switch (actionName) {
					case I:
						if (ctrlValue.trim().equalsIgnoreCase(""))
							break;
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
						WebHelper.currentdriver = Automation.driver;
						WebHelperClaims.implementWait();
						break;
					case NC:
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
						WebHelper.currentdriver = Automation.driver;
						WebHelperClaims.implementWait();
						break;
					default:
						break;
					}
					break;

				case URL:
					switch (actionName) {
					case I:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
							if (ITAFWebDriver.isClaimsApplication() || (ITAFWebDriver.isSuiteApplication()
									&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims"))) {
								try {
									WebHelper.currentdriver.navigate().to(controlName);
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(5));
									@SuppressWarnings("unused")
									WebElement webElement1 = WaitForPageLoad
											.until(ExpectedConditions.elementToBeClickable(
													By.xpath("//div[contains(@class,'loginPage-majescologo')]")));
								} catch (Exception e) {
									try {
										WebHelper.currentdriver.navigate().refresh();
										WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
												Duration.ofSeconds(5));
										@SuppressWarnings("unused")
										WebElement webElement1 = WaitForPageLoad
												.until(ExpectedConditions.elementToBeClickable(
														By.xpath("//div[contains(@class,'loginPage-majescologo')]")));
									} catch (Exception ex) {
										WebHelper.currentdriver.navigate().to(controlName);
									}

								}
							} else {
								WebHelper.currentdriver.navigate().to(ctrlValue);
							}
						}
						break;

					case NC:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
							if (ITAFWebDriver.isClaimsApplication() || (ITAFWebDriver.isSuiteApplication()
									&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims"))) // For CS
							// Enviornment
							{
								WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
										Duration.ofSeconds(7));
								try {
									Automation.driver.get(controlName);
									Thread.sleep(200);
									claimLoaderWait();
									@SuppressWarnings("unused")
									WebElement webElementLogin = WaitForPageLoad.until(ExpectedConditions
											.elementToBeClickable(By.xpath("//input[@name='username']")));
									Automation.driver.get(controlName);
								} catch (Exception exx) {
									// WebHelper.currentdriver.navigate().to(controlName);
									try {

										Automation.driver.get(controlName);
										Thread.sleep(200);
										claimLoaderWait();

										try {
											Thread.sleep(200);
											claimLoaderWait();
											WebDriverWait WaitForPageLoadWhtsNew = new WebDriverWait(Automation.driver,
													Duration.ofSeconds(15));

											WebElement WhatsNewClose = WaitForPageLoadWhtsNew
													.until(ExpectedConditions.elementToBeClickable(By.xpath(
															"//div[@class='modal-dialog']//button[@class='close']/span")));
											WhatsNewClose.click();
										} catch (Exception ex1) {
											log.info("No What's New Pop up found");
										}

										Thread.sleep(200);
										WebElement webElementLogoutArrow = WaitForPageLoad
												.until(ExpectedConditions.elementToBeClickable(
														By.xpath("//div[@class='loginDetailSecDownArrow']")));
										Thread.sleep(200);
										webElementLogoutArrow.click();
										WebElement webElementLogoutB = WaitForPageLoad.until(
												ExpectedConditions.elementToBeClickable(By.xpath("//a[@id='logout']")));
										webElementLogoutB.click();
										WebElement webElementLogin = WaitForPageLoad.until(ExpectedConditions
												.elementToBeClickable(By.xpath("//a[contains(text(),'Log in')]")));
										webElementLogin.click();

									} catch (Exception exxx) {
										Automation.driver.navigate().refresh();
										Thread.sleep(200);
										claimLoaderWait();

										WebElement webElementLogoutArrow = WaitForPageLoad
												.until(ExpectedConditions.elementToBeClickable(
														By.xpath("//div[@class='loginDetailSecDownArrow']")));
										webElementLogoutArrow.click();
										WebElement webElementLogoutB = WaitForPageLoad.until(
												ExpectedConditions.elementToBeClickable(By.xpath("//a[@id='logout']")));
										webElementLogoutB.click();
										WebElement webElementLogin = WaitForPageLoad.until(ExpectedConditions
												.elementToBeClickable(By.xpath("//a[contains(text(),'Log in')]")));
										webElementLogin.click();
									}
								}
							} else {
								WebHelper.currentdriver.navigate().to(controlName);
								WebHelper.currentdriver.navigate().refresh();
							}

						}
						break;
					default:
						break;
					}
					break;

				case Alert: // Pradeep-Added case from PAS
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

							Thread.sleep(1000);
						}
						break;
					case NC:
						// wait.until(ExpectedConditions.presenceOfElementLocated(alert()));
						Alert alert1 = Automation.driver.switchTo().alert();
						if (alert1 != null) {
							alert1.accept();
							Thread.sleep(1000);
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
					default:
						break;
					}
					break;
				// case Alert :
				case Menu:
				case WebImage:
				case ActionClick:

					switch (actionName) {
					case I:
						if (!ctrlValue.equalsIgnoreCase("")) {

							try {

								@SuppressWarnings("unused")
								WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
										Duration.ofSeconds(90));
								claimLoaderWait();
								scroll(controlName, webElement);
								// ((JavascriptExecutor)
								// Automation.driver).executeScript("arguments[0].scrollIntoView();",
								// webElement);
								// Thread.sleep(1000);

								Actions builderClick = new Actions(Automation.driver);
								// highlightElement(webElement);
								Action clickAction = builderClick.moveToElement(webElement).clickAndHold().release()
										.build();
								clickAction.perform();

							} catch (Exception e) {
								log.error(e.getMessage(), e);
								log.info("Element not exist");
							}
						}
						break;

					case NC:
						scroll(controlName, webElement);
						Actions builderClick = new Actions(Automation.driver);
						// highlightElement(webElement);
						Action clickAction = builderClick.moveToElement(webElement).clickAndHold().release().build();
						clickAction.perform();
						break;
					default:
						break;
					}
					break;

				case ActionDoubleClick:
					switch (actionName) {
					case I:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
							Actions builderdoubleClick = new Actions(Automation.driver);
							builderdoubleClick.doubleClick(webElement).build().perform();
						}
						break;
					default:
						break;
					}
					break;

				case ActionClickandEsc:
					return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
							controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
							webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo, operationType,
							cycleDate, TransactionType);

				case ActionMouseOver:
					switch (actionName) {
					/*
					 * case I: if (ctrlValue.equalsIgnoreCase("Y") ||
					 * ctrlValue.equalsIgnoreCase("Yes") ||
					 * !(ctrlValue.trim().equalsIgnoreCase(""))) { WebDriverWait WaitForPageLoad =
					 * new WebDriverWait(Automation.driver, Duration.ofSeconds(30));
					 * claimLoaderWait();
					 * WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
					 * WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
					 * Actions builderMouserOver = new Actions(WebHelper.currentdriver);
					 * builderMouserOver.moveToElement(webElement).click().build().perform(); //
					 * Pradeep-Click } default: break;
					 */
					case NC:
						if (webElement != null) {
							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
									webElement);
							Thread.sleep(500);
							claimLoaderWait();
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
								claimLoaderWait();
								Actions builderMouserOver = new Actions(Automation.driver);
								Action mouseOverAction = builderMouserOver.moveToElement(webElement).build();
								mouseOverAction.perform();
							} else {
								log.info("ActionMouseOver failed as webelement is not present on screen");
							}
						}
						break;
					default:
						break;

					}
					break;

				case Calendar:
				case CalendarNew:
				case CalendarIPF:
				case CalendarEBP:
					return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
							controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
							webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo, operationType,
							cycleDate, TransactionType);

				case Window:
					switch (actionName) {
					case O:
						return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
								controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
								webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo,
								operationType, cycleDate, TransactionType);
					case I: // _ Claims_handle AJAX pop-up
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("WK")
								|| ctrlValue.equalsIgnoreCase("D") || ctrlValue.equalsIgnoreCase("DP")
								|| ctrlValue.equalsIgnoreCase("C") || ctrlValue.equalsIgnoreCase("Ok")
								|| ctrlValue.equalsIgnoreCase("Ok1") || !(ctrlValue.trim().equalsIgnoreCase(""))) {
							claimLoaderWait();
							if (ctrlValue.equalsIgnoreCase("Y")) {
								try {
									webElement = WebHelper.wait.until(ExpectedConditions.elementToBeClickable(
											By.xpath("//button[@name='Ok' or @name='btnOk' or @name='OK']")));// Added
									WebHelperUtil.saveScreenShot();// 23/10/23 // popup
									Thread.sleep(500);
									// webElement.click();
									Actions builderClick = new Actions(Automation.driver);
									Action clickAction = builderClick.moveToElement(webElement).clickAndHold().release()
											.build();
									clickAction.perform();
									log.info(ctrlValue + " " + "button clicked successfully");
								} catch (Exception e) {
									log.error(e.getMessage(), e);
									webElement = WebHelper.wait.until(ExpectedConditions
											.elementToBeClickable(By.xpath("//button[@name='btnOk']")));
									Thread.sleep(500);
									// webElement.click();
									Actions builderClick = new Actions(Automation.driver);
									Action clickAction = builderClick.moveToElement(webElement).clickAndHold().release()
											.build();
									clickAction.perform();
									log.info("Pop-up found on the web page" + "ctrlValue");
									Thread.sleep(1000);
								}
								
							} else if ((ctrlValue.equalsIgnoreCase("D"))) {
								try {
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(7));
									WebElement webElement1 = WaitForPageLoad.until(
											ExpectedConditions.elementToBeClickable(By.name("btnProceedWithFNOL")));
									// webElement=
									// wait.until(ExpectedConditions.elementToBeClickable(By.name("btnProceedWithFNOL")));
									Thread.sleep(500);
									// webElement1.click();
									Actions builderClick = new Actions(Automation.driver);
									Action clickAction = builderClick.moveToElement(webElement1).clickAndHold()
											.release().build();
									clickAction.perform();

									log.info("Pop-up found on the web page" + ctrlValue);
									Thread.sleep(1000);
								} catch (Exception e) {
									// log.error(e.getMessage(), e);
									log.info("No duplicate pop up found");
								}
								
								//New condition added to handle - btnProceedWithFROI option - 18/02/2025
							} else if ((ctrlValue.equalsIgnoreCase("D1"))) {
								log.info("ctrlValue is:" + ctrlValue);
								try {
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(7));
									WebElement webElement1 = WaitForPageLoad.until(
											ExpectedConditions.elementToBeClickable(By.name("btnProceedWithFROI")));
									// webElement=
									// wait.until(ExpectedConditions.elementToBeClickable(By.name("btnProceedWithFNOL")));
									Thread.sleep(500);
									// webElement1.click();
									Actions builderClick = new Actions(Automation.driver);
									Action clickAction = builderClick.moveToElement(webElement1).clickAndHold()
											.release().build();
									clickAction.perform();

									log.info("Pop-up found on the web page" + ctrlValue);
									Thread.sleep(1000);
								} catch (Exception e) {
									// log.error(e.getMessage(), e);
									log.info("No duplicate pop up found");
								}
								
							} else if ((ctrlValue.equalsIgnoreCase("DP"))) {
								WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
										Duration.ofSeconds(5));
								WebElement webElement1 = WaitForPageLoad
										.until(ExpectedConditions.elementToBeClickable(By.xpath(
												"//div[@data-labelkey='mm.icd.ClaimModification.Common.DuplicatePaymentFound']//div[@data-cid='btnContinue']")));
								// webElement=
								// wait.until(ExpectedConditions.elementToBeClickable(By.name("btnProceedWithFNOL")));
								Thread.sleep(500);
								// webElement1.click();
								Actions builderClick = new Actions(Automation.driver);
								Action clickAction = builderClick.moveToElement(webElement1).clickAndHold().release()
										.build();
								clickAction.perform();
								log.info("Pop-up found on the web page" + "ctrlValue");
								Thread.sleep(1000);
							} else if ((ctrlValue.equalsIgnoreCase("R"))) {
								try {
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(7));
									WebElement webElement1 = WaitForPageLoad
											.until(ExpectedConditions.elementToBeClickable(By.name("btnReturn")));
									// webElement=
									// wait.until(ExpectedConditions.elementToBeClickable(By.name("btnProceedWithFNOL")));
									Thread.sleep(500);
									// webElement1.click();
									Actions builderClick = new Actions(Automation.driver);
									Action clickAction = builderClick.moveToElement(webElement1).clickAndHold()
											.release().build();
									clickAction.perform();

									log.info("Pop-up found on the web page" + "ctrlValue");
									Thread.sleep(1000);
								} catch (Exception e) {
									// log.error(e.getMessage(), e);
									log.info("No duplicate pop up found");
								}
							} else if ((ctrlValue.equalsIgnoreCase("C"))) {
								webElement = WebHelper.wait.until(
										ExpectedConditions.elementToBeClickable(By.xpath("//button[@name='Cancel']")));
								Thread.sleep(500);
								// webElement.click();
								Actions builderClick = new Actions(Automation.driver);
								Action clickAction = builderClick.moveToElement(webElement).clickAndHold().release()
										.build();
								clickAction.perform();
								log.info("Pop-up found on the web page" + "ctrlValue");
								Thread.sleep(1000);
							} else if ((ctrlValue.equalsIgnoreCase("Ok"))) {
								webElement = WebHelper.wait.until(
										ExpectedConditions.elementToBeClickable(By.xpath("//button[@name='btnOk']")));
								Thread.sleep(500);
								// webElement.click();
								Actions builderClick = new Actions(Automation.driver);
								Action clickAction = builderClick.moveToElement(webElement).clickAndHold().release()
										.build();
								clickAction.perform();
								log.info("Pop-up found on the web page" + "ctrlValue");
								Thread.sleep(1000);
							} else if ((ctrlValue.equalsIgnoreCase("Ok1"))) {
								try {
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(7));
									// WebElement webElement1=
									// WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(By.name("btnProceedWithFNOL")));
									WebElement webElement1 = WaitForPageLoad.until(ExpectedConditions
											.elementToBeClickable(By.xpath("//button[@name='Ok' or @name='btnOk']")));
									Thread.sleep(500);
									// webElement1.click();
									Actions builderClick = new Actions(Automation.driver);
									Action clickAction = builderClick.moveToElement(webElement1).clickAndHold()
											.release().build();
									clickAction.perform();

									log.info("Pop-up found on the web page" + "ctrlValue");
								} catch (Exception e) {
									// log.error(e.getMessage(), e);
									log.info("No Pop up found");
								}
							} else if ((ctrlValue.equalsIgnoreCase("OKS"))) {// Pradeep- Added new condition to handle
																				// new Xpath for OK button.12/26/2019
								try {
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(7));
									WebElement webElement1 = WaitForPageLoad.until(ExpectedConditions
											.elementToBeClickable(By.xpath("//label[contains(text(),'OK')]")));
									Thread.sleep(500);
									// webElement1.click();
									Actions builderClick = new Actions(Automation.driver);
									Action clickAction = builderClick.moveToElement(webElement1).clickAndHold()
											.release().build();
									clickAction.perform();
									log.info("Pop-up found on the web page" + "ctrlValue");
									Thread.sleep(1000);
								} catch (Exception e) {
									// log.error(e.getMessage(), e);
									log.info("NO duplicate pop up found");
								}
							} else {
								log.info("No window Found");
							}
						}

						break;

					case NC: // Pradeep-Created NC case without
						// "ControlValue" to handle random popup
						// on screen
						try {
							WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(7));
							// WebElement webElement1=
							// WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(By.name("btnProceedWithFNOL")));
							WebElement webElement1 = WaitForPageLoad
									.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
							webElement1.click();
							log.info("Pop-up found on the web page");
						} catch (Exception e) {
							// log.error(e.getMessage(), e);//Not required as it is adding to the log
							// file.06/07/2023
							log.info("No Pop up found");
						}
						break;

					default:
						break;
					}
					break;

				case LoaderWait:// Pradeep- Added case to handle Loader wait condition.12/26/2019.
					switch (actionName) {
					case NC:
						WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(120));
						WaitForPageLoad.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
					default:
						break;
					}
					break;

				case CRMLoaderWait:// Pradeep- Added case to handle CRM Loader wait condition
					switch (actionName) {
					case NC:
						// WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
						// Duration.ofSeconds(120));
						WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(400));
						WaitForPageLoad.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(CRMloader)));
					default:
						break;
					}
					break;

				case WebTable:
					switch (actionName) {
					case Read:
					case Write:
					case NC:
					case V:
						return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
								controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
								webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo,
								operationType, cycleDate, TransactionType);

					case TableInput:
						WebHelper.findtablefound = WebHelper.currentdriver.findElements(By.xpath(controlName))
								.size() > 0;
						if (WebHelper.findtablefound == true) {
							WebElement tableFound = WebHelper.wait
									.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
							BillingProduct.TableInputAction(tableFound, controlName, logicalName, rowValues,
									WebHelper.valuesHeader, ExcelUtility.TIvaluesheetrows);
							// Thread.sleep(1000);
						} else {
							break;
						}
						break;

					case FIND:
						WebHelper.findtablefound = WebHelper.currentdriver.findElements(By.xpath(controlName))
								.size() > 0;
						if (WebHelper.findtablefound == true) {
							WebElement tableFound = WebHelper.wait
									.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
							BillingProduct.findAction(tableFound, controlName, logicalName, rowValues,
									WebHelper.valuesHeader);
							// Thread.sleep(1000);
						} else {
							break;
						}
						break;

					case I:
						WebHelper.findtablefound = WebHelper.currentdriver.findElements(By.xpath(controlName))
								.size() > 0;
						if (WebHelper.findtablefound == true) {
							WebElement tableFound = WebHelper.wait
									.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
							List<WebElement> table_Rows = tableFound.findElements(By.tagName("tr"));
							List<WebElement> table_Columns = table_Rows.get(1).findElements(By.tagName("td"));

							int ApplicationtableRowsize = table_Rows.size(); // ApplicationtableRowsize
							// =
							// no
							// of
							// rows
							// in
							// the
							// WebTable
							int Applicationtablecolumnsize = table_Columns.size(); // Applicationtablecolumnsize
							// =
							// no
							// of
							// columns
							// in
							// the
							// WebTable

							String ColumnName = ctrlValue.split(",")[0];
							String ColumnType = ctrlValue.split(",")[1];

							for (int i = 1; i <= Applicationtablecolumnsize; i++) {
								Thread.sleep(1000);
								String ApplicationColumnHeaderxapth = controlName + "/thead/tr/th[" + i + "]";
								log.info("ApplicationColumnHeader is:" + ApplicationColumnHeaderxapth);
								WebElement element = WebHelper.currentdriver
										.findElement(By.xpath(ApplicationColumnHeaderxapth));
								String ApplicationColumnHeader = element.getText();
								if ((ColumnName).equalsIgnoreCase(ApplicationColumnHeader)) {
									for (int r = 1; r <= ApplicationtableRowsize; r++) {
										if (ColumnType.equalsIgnoreCase("Webcheckbox")) {
											String XPath = controlName + "/tbody/tr[" + r + "]/td[" + i
													+ "]/div/div/input";
											WebHelper.objfound = WebHelper.currentdriver.findElements(By.xpath(XPath))
													.size() > 0;
											if (WebHelper.objfound == true) {
												WebElement newelement = WebHelper.currentdriver
														.findElement(By.xpath(XPath));
												newelement.click();
												Thread.sleep(500);
												((JavascriptExecutor) WebHelper.currentdriver)
														.executeScript("arguments[0].scrollIntoView();", newelement);
												Thread.sleep(500);
												WebHelper.objfound = false;
											}

										} else if (ColumnType.equalsIgnoreCase("WebLink")) {
											String XPath = controlName + "/tbody/tr[" + r + "]/td[" + i + "]/div/span";
											WebHelper.objfound = WebHelper.currentdriver.findElements(By.xpath(XPath))
													.size() > 0;
											if (WebHelper.objfound == true) {
												WebElement newelement = WebHelper.currentdriver
														.findElement(By.xpath(XPath));
												log.info("link xpath " + XPath);
												newelement.click();
												Thread.sleep(500);
												((JavascriptExecutor) WebHelper.currentdriver)
														.executeScript("arguments[0].scrollIntoView();", newelement);
												Thread.sleep(500);
												WebHelper.objfound = false;
											}
										} else if (ColumnType.equalsIgnoreCase("WebCheckBox")) {
											// not encountered
										}
									}
								}
							}

						}
						break;
					default:
						break;
					}
					break;

				// bhaskar capture screenshot START
				case Screenshot:
					switch (actionName) {
					case NC:
						return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
								controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
								webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo,
								operationType, cycleDate, TransactionType);
					default:
						break;
					}
					break;

				case Robot:
					if (controlName.equalsIgnoreCase("SetFilePath")) {
						StringSelection stringSelection = new StringSelection(ctrlValue);
						Toolkit.getDefaultToolkit().getSystemClipboard().setContents(stringSelection, null);
						WebHelper.robot.delay(500);
						WebHelper.robot.keyPress(KeyEvent.VK_CONTROL);
						WebHelper.robot.keyPress(KeyEvent.VK_V);
						WebHelper.robot.keyRelease(KeyEvent.VK_V);
						WebHelper.robot.keyRelease(KeyEvent.VK_CONTROL);
						WebHelper.robot.keyPress(KeyEvent.VK_ENTER);
						WebHelper.robot.keyRelease(KeyEvent.VK_ENTER);
						Thread.sleep(200);
					} else if (controlName.equalsIgnoreCase("TAB")) {
						WebHelper.robot.keyPress(KeyEvent.VK_TAB);
						WebHelper.robot.keyRelease(KeyEvent.VK_TAB);
					} else if (controlName.equalsIgnoreCase("SPACE")) {
						WebHelper.robot.keyPress(KeyEvent.VK_SPACE);
						WebHelper.robot.keyRelease(KeyEvent.VK_SPACE);
					} else if (controlName.equalsIgnoreCase("ENTER")) {
						WebHelper.robot.keyPress(KeyEvent.VK_ENTER);
						WebHelper.robot.keyRelease(KeyEvent.VK_ENTER);
						Thread.sleep(200);
					}
					break;

				case DB:
				case Database:// Minaakshi : Case added to fetch Unique
					// values
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

							if (readFromColName != null && !readFromColName.isEmpty()) {
								String tempValue = DMProduct.ReadFromExcelUsingColumnName("", readFromColName);

								if (tempValue == null) {
									/*
									 * String errorMessage = ("Column Name - " + readFromColName +
									 * " from unique no sheet has returned - " + tempValue + " value. " + "Hence" +
									 * "SQL query - " + sqlQuery + " will not be triggered");
									 */
									String errorMessage = ("Column Name - " + readFromColName
											+ " from unique no sheet has returned - " + tempValue + " value. "
											+ "Hence SQL query will not be triggered.");
									// log.info(errorMessage);
									log.info("Transaction Failed");
									report.setStatus("FAIL");
									report.setMessage(errorMessage);
									resetVariables();
									throw new Exception(errorMessage);
								}

								else {
									sqlQuery = getDataFetchingSQLQuery(sqlQuery, tempValue);
								}

							}
							ResultSet rs3;
							Connection conn2 = JDBCConnection.establishPASDBConn();
							Statement st2 = conn2.createStatement();
							rs3 = st2.executeQuery(sqlQuery);
							rs3.next();
							// ResultSet rs3 =
							// JDBCConnection.establishPASDBConn(sqlQuery);

							log.info("toBeFetchedDBColName :" + toBeFetchedDBColName);
							// ArrayList toBeFetchedDBColNameList = new ArrayList();//Sel4
							ArrayList<String> toBeFetchedDBColNameList = new ArrayList<>();
							String toBeFetchedDBColNameParts = "";

							if (toBeFetchedDBColName.contains(","))// To fetch multiple records
							{
								String[] toBeFetchedDBColNamePart = toBeFetchedDBColName.split(",");

								for (int i = 0; i < toBeFetchedDBColNamePart.length; i++) {
									toBeFetchedDBColNameParts = String
											.valueOf(rs3.getString(toBeFetchedDBColNamePart[i]));
									toBeFetchedDBColNameList.add(toBeFetchedDBColNameParts);
								}
								log.info(toBeFetchedDBColNameList);
								String toBeFetchedDBColNameMain = "";
								for (int i = 0; i < toBeFetchedDBColNameList.size(); i++) {
									toBeFetchedDBColNameMain = (String) toBeFetchedDBColNameList.get(i) + ','
											+ toBeFetchedDBColNameMain;
								}
								log.info(toBeFetchedDBColNameMain);
								ctrlValue = toBeFetchedDBColNameMain.substring(0, toBeFetchedDBColNameMain.length() - 1)
										+ "";

							} else {
								ctrlValue = String.valueOf(rs3.getString(toBeFetchedDBColName));
							}

							rs3.close();

							DMProduct.writeDataToUniqueNumberSheet(ctrlValue, writeToColName, "");
							resetVariables();
						}
						break;

					case VerifyDataFromDB: // _Verification Fetch data with DB Value

						Reporter report = new Reporter();
						report.setReport(report);
						if (logicalName.equalsIgnoreCase("SQLQuery"))
							sqlQuery = ctrlValue;
						else if (logicalName.equalsIgnoreCase("ReadFromColName") && !ctrlValue.equalsIgnoreCase(""))
							readFromColName = ctrlValue;
						else if (logicalName.equalsIgnoreCase("ColumnToBeFetchedFromDB"))
							toBeFetchedDBColName = ctrlValue;
						else if (logicalName.equalsIgnoreCase("VerifyColNameExpected")) {
							verifyColNameExpected = ctrlValue;

							// log.info("Checking expected value");
							WebHelper.ExpectedValue = DMProduct.ReadFromExcelUsingColumnName("", verifyColNameExpected);
							// log.info("WbHelper.ExpectedValue is :" + WebHelper.ExpectedValue);

							// log.info("5026");
							if (readFromColName != null && !readFromColName.isEmpty()) {
								// log.info("5027");
								String tempValue = DMProduct.ReadFromExcelUsingColumnName("", readFromColName);
								// log.info("tempValue is :"+ tempValue);
								// log.info("5029");
								sqlQuery = getDataFetchingSQLQuery(sqlQuery, tempValue);
							}
							// log.info("5030");

							ResultSet rs3;
							Connection conn2 = JDBCConnection.establishPASDBConn();
							Statement st2 = conn2.createStatement();
							rs3 = st2.executeQuery(sqlQuery);
							rs3.next();
							// log.info("5037");

							WebHelper.ActualValue = String.valueOf(rs3.getString(toBeFetchedDBColName));

							log.info("ActualValue is : " + WebHelper.ActualValue);
							log.info("ExpectedValue is : " + WebHelper.ExpectedValue);
							// if(ActualValue.equalsIgnoreCase(ExpectedValue))
							try {
								if (WebHelper.ActualValue.contains(WebHelper.ExpectedValue)) {
									report.setStatus("PASS");
									report.setStatus(report.getStatus());
									report.setMessage("Values Matched");
									// WebHelper.saveScreenShot();
								} else {
									report.setStatus("FAIL");
									report.setStatus(report.getStatus());
									report.setMessage("Values Not Matched");
									resetVariables();
								}
							} catch (Exception e) {
								String errorMessage = e.getMessage();
								String firstLine = errorMessage.split("\\r?\\n")[0];
								log.error("Error :" + firstLine + " - Expected Value from unique no sheet: "
										+ WebHelper.ExpectedValue + " does not match the Actual Value from DB :"
										+ WebHelper.ActualValue);
								// Concatenate the custom message with the first line of the exception
								// String customMessage = "Expected Value from unique no sheet: "+
								// WebHelper.ExpectedValue + " does not match the Actual Value from DB :" +
								// WebHelper.ActualValue;
								/*
								 * String customMessage = "Expected Value from unique no sheet: " +
								 * (WebHelper.ExpectedValue != null ? WebHelper.ExpectedValue : "null") +
								 * " does not match the Actual Value from DB: " + (WebHelper.ActualValue != null
								 * ? WebHelper.ActualValue : "null");
								 */
								String customMessage = "Expected Value from unique no sheet: "
										+ (WebHelper.ExpectedValue != null ? WebHelper.ExpectedValue.toString()
												: "null")
										+ " does not match the Actual Value from DB : "
										+ (WebHelper.ActualValue != null ? WebHelper.ActualValue.toString() : "null");
								resetVariables();
								// Throw a new exception with the custom message
								throw new Exception(customMessage);
							}

							WebHelper.columns.add("");
							WebHelper.columnsData.add(WebHelper.columns);
							int temprowcount = 0;
							int tempcolcount = 0;

							ExcelUtility.WriteToCompareDetailResults(testCase, TransactionType, WebHelper.columns,
									WebHelper.columnsData, temprowcount, tempcolcount, report, WebHelper.ExpectedValue,
									WebHelper.ActualValue, logicalName, operationType, cycleDate);
							resetVariables();
						}

						break;
					case UpdateDB: // _ Update Query DB
						if (logicalName.equalsIgnoreCase("SQLQuery"))
							sqlQuery = ctrlValue;
						else if (logicalName.equalsIgnoreCase("ReadFromColName"))
							readFromColName = ctrlValue;
						else if (logicalName.equalsIgnoreCase("ColumnToBeFetchedFromDB"))
							toBeFetchedDBColName = ctrlValue;
						else if (logicalName.equalsIgnoreCase("WriteToColName")) {
							writeToColName = ctrlValue;

							if (sqlQuery.contains("'$datevalue'")) { // Add todays_date
								sqlQuery = AddTodayDate(sqlQuery);
							}

							if (sqlQuery.contains("'$randomvalue'")) { // Add Random Check number
								sqlQuery = AddRandomCodeGenerator(sqlQuery);
							}

							String tempValue = DMProduct.ReadFromExcelUsingColumnName("", readFromColName);
							sqlQuery = getDataFetchingSQLQuery(sqlQuery, tempValue);

							// ResultSet rs3;
							Connection conn2 = JDBCConnection.establishPASDBConn();
							Statement st2 = conn2.createStatement();

							st2.executeQuery(sqlQuery);
							if (Config.databaseType.equalsIgnoreCase("ORACLE")) {
								st2.execute("commit");
							}
							st2.close();
							JDBCConnection.closeConnection(conn2);
						}

						break;
					default:
						break;
					}
					break;

				case WaitForEC:
				case SikuliScreen:
				case SikuliType:
				case SikuliButton:
				case Slider:
				case Date:
					return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
							controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
							webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo, operationType,
							cycleDate, TransactionType);

				case ClearCache:
					return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
							controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
							webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo, operationType,
							cycleDate, TransactionType);

				case MaskedInputDate:
					if (!ctrlValue.equalsIgnoreCase("null")) {
						webElement.clear();
						webElement.click();
						webElement.sendKeys(ctrlValue);

					} else {
						webElement.clear();
					}
					break;
				// bhaskar

				case FileUpload:
					webElement.sendKeys(ctrlValue);
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
					default:
						break;
					}
					break;

				case ScrollTo:
					/*
					 * Locatable element = (Locatable) webElement; Point p=
					 * element.getCoordinates().onScreen(); JavascriptExecutor js =
					 * (JavascriptExecutor) currentdriver; js.executeScript("window.scrollTo(" +
					 * p.getX() + "," + (p.getY()+150) + ");");
					 */

				case NC: // ScrollUp_
					if (!ctrlValue.equalsIgnoreCase("")) {
						claimLoaderWait();
						if (ctrlValue.equalsIgnoreCase("Up")) {
							JavascriptExecutor jse = (JavascriptExecutor) WebHelper.currentdriver;
							jse.executeScript("window.scrollBy(0,-250)", "");
							Thread.sleep(100);
							break;
						} else if (ctrlValue.equalsIgnoreCase("Down")) {
							JavascriptExecutor jse = (JavascriptExecutor) WebHelper.currentdriver;
							jse.executeScript("window.scrollBy(0,250)", "");
							Thread.sleep(100);
							break;
						} else if (ctrlValue.equalsIgnoreCase("DoubleDown")) {
							JavascriptExecutor jse = (JavascriptExecutor) WebHelper.currentdriver;
							jse.executeScript("window.scrollBy(0,775)", "");
							Thread.sleep(100);
							break;
						} else if (ctrlValue.equalsIgnoreCase("DoubleUp")) {
							JavascriptExecutor jse = (JavascriptExecutor) WebHelper.currentdriver;
							jse.executeScript("window.scrollBy(0,-775)", "");
							Thread.sleep(100);
							break;
						} else {
							log.info("Invalid contol value");
						}
					} else {
						log.info("no scroll");
					}
					break;

				case I:
					JavascriptExecutor jsc = (JavascriptExecutor) WebHelper.currentdriver;
					jsc.executeScript("window.scrollBy(0,250)", "");
					/*
					 * try{ //Asif:Added 31-May:Start if(webElement.isDisplayed()){ Coordinates
					 * coordinate = ((Locatable)webElement).getCoordinates(); coordinate.onPage();
					 * coordinate.inViewPort(); } }catch(Exception e){
					 * log.info("Element does not exist");
					 * 
					 * }//Asif:Added 31-May:End
					 */
					break;

				case ScrollToElement2:
					switch (actionName) {
					case I:
						if (!ctrlValue.equalsIgnoreCase("")) {
							((JavascriptExecutor) WebHelper.currentdriver)
									.executeScript("arguments[0].scrollIntoView();", webElement);
							Thread.sleep(300);
							((JavascriptExecutor) WebHelper.currentdriver).executeScript("window.scrollBy(0,-40)", "");
						}
						break;
					default:
						break;
					}
					break;

				case ScrollToElement: // Pradeep-Added case from PAS
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

					case I_$Value: // Pradeep- Added case to handle Scroll to Element as per dynamic Xpath contains
									// Value and ClaimNo.12/26/2019.

						if (controlId.equalsIgnoreCase("XPath$Value$ClaimNo$") && !ctrlValue.equalsIgnoreCase("")) {
							String tempCtrlName = controlName;
							String uniqueNumber1 = ReadFromExcel(ctrlValue);
							String tempReplaceString = tempCtrlName.replace("$ClaimNo$", uniqueNumber1);

							String ExpectedValue = tempReplaceString.replace("$Value", ctrlValue);
							// WebElement DynamicElement = WebHelper.wait
							// .until(ExpectedConditions.elementToBeClickable(By.xpath(ExpectedValue)));
							WebElement webElement1 = WebHelper.currentdriver.findElement(By.xpath(ExpectedValue));
							// Thread.sleep(1000);
							// DynamicElement.click();

							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
									webElement1);
							break;

						} else {
							log.info("No Value Found");
						}
					default:
						break;

					}
					break;

				case DownloadDocument: // Pradeep- Added case of PDF
					// Comparison from Helper.PAS
				case PDFDocumentCompare:
				case IgnoreString:
				case RenameDocument:
				case MoveXMLDocument:
				case MoveDocument:
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case RenameFile:
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case WaitTillFileDownload:
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case PDFComparisonMode:
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case WebService:
				case WebService1:
				case WebService2:
				case WebService3:
				case WebServiceV:
				case WebServiceC:
				case WebServiceRP:
				case WebServiceV1:
				case WebServiceV2:
				case WebServiceVAG:
				case WebServiceV3:
				case OpenAPI:
				case CopyInterfaceFile:// 17/7/23 - For Claims Flat File Requirement.
					return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
							controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
							webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo, operationType,
							cycleDate, TransactionType);

				case ReplaceDefault:
					switch (actionName) {
					case I:
						if (ctrlValue == null || ctrlValue.equalsIgnoreCase("")) {
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
								Thread.sleep(500);
								temp = webElement.getAttribute("value");
								if (temp.equals("")) {
									temp = webElement.getText();
								}
								counter++;
								if (counter == 2) {
									log.info("Could not set correct value");
									break;
								}
							}
						} catch (StaleElementReferenceException e) {
							e.toString();
							log.info("Trying to recover from a stale element :" + e.getMessage());
						}
						break;
					default:
						break;
					}
					break;

				default:
					log.info("U r in Default for missing controltype: " + controlTypeEnum);
					break;
				}
				// To handle processing icon in Claims
				try {
					Thread.sleep(200);
					for (int i = 1; i <= 340; i++) {
						List<WebElement> loader = Automation.driver
								.findElements(By.xpath("//body[contains(@class,'waiting')]"));
						int loaderSize = loader.size();
						// log.info("loaderSize is-->" + loaderSize);
						if (loaderSize >= 1)
							Thread.sleep(400);
						else
							break;
					}
				} catch (Exception e) {
				}

				// To handle processing bar in policy 2015 | Added specifically for Suite\SIT
				// automation
				try {
					// Thread.sleep(200);
					for (int i = 1; i <= 340; i++) {
						List<WebElement> loader = Automation.driver.findElements(By.xpath(
								"//div[@class='v-loading-indicator first' and @style='position: absolute; display: block;']"));
						int loaderSize = loader.size();
						// log.info("loaderSize is-->" + loaderSize);
						if (loaderSize >= 1)
							Thread.sleep(300);
						else
							break;
					}
				} catch (Exception e) {
				}

			} catch (WebDriverException we) {
				try {
					if (StringUtils.contains(we.getMessage().toLowerCase(), ("not clickable at point"))) {
						log.info("Click event failed. Tried JS scroll and click.");
						((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].scrollIntoView();",
								webElement);
						webElement.click();
					} else
						throw new Exception("WebDriverException Occurred from Do Action : " + we.getMessage());
				} catch (Exception e) {
					throw new Exception("WebDriverException Occurred from Do Action : " + we.getMessage());
				}
			} catch (FileNotFoundException e) {
				// Specific handling for FileNotFoundException
				// Log the exception or perform any other necessary actions
				throw e; // Re-throwing the exception to be handled by the calling code
			} catch (IOException ioe) {
				// log.error(ioe.getMessage(), ioe);
				throw new Exception("IOException Occurred from Do Action : " + ioe.getMessage());
			}

			catch (Exception e) {
				String errorMessage = e.getMessage();
				String firstLine = errorMessage.split("\\r?\\n")[0];
				log.error("Error :" + firstLine);
				// log.error(e.getMessage(), e);
				throw new Exception("Error : " + e.getMessage());
			}
		}
		if ((action.toString().equalsIgnoreCase("V") || action.toString().equalsIgnoreCase("FormVerificationDetails")
				|| action.toString().equalsIgnoreCase("VFormCount")
				|| action.toString().equalsIgnoreCase("VFormPresence")
				|| action.toString().equalsIgnoreCase("V_availabilityStatus")
				|| action.toString().equalsIgnoreCase("LookupVerificationDetails")
				|| action.toString().equalsIgnoreCase("VLookupValuesCount")
				|| action.toString().equalsIgnoreCase("VLookupValuesPresence")
				|| action.toString().equalsIgnoreCase("V_date") || action.toString().equalsIgnoreCase("V_text")
				|| action.toString().equalsIgnoreCase("V_attributeValue") || action.toString().equalsIgnoreCase("F")
				|| action.toString().equalsIgnoreCase("VA") || action.toString().equalsIgnoreCase("V_Disable")
				|| action.toString().equalsIgnoreCase("V_List")
				|| action.toString().equalsIgnoreCase("V_EnableDisable")) && !ctrlValue.equalsIgnoreCase("")) {
			if (Results == true) {
				webDriver.setReport(WebHelperUtil.WriteToDetailResults(ctrlValue, currentValue, logicalName));
			}
		}
		return currentValue;
	}

	public static String getMonth() {
		return WebHelper.month;
	}

	public static void scroll(String XPath, WebElement we) {
		try {
			// ((JavascriptExecutor)
			// WebHelper.currentdriver).executeScript("arguments[0].scrollIntoView();", we);
			((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].scrollIntoViewIfNeeded(true);",
					we);
			// ((JavascriptExecutor)
			// WebHelper.currentdriver).executeScript("window.scrollBy(0,-25)", "");
		} catch (Exception e) {
		}
	}

	// Mandar Test
	/*
	 * public static void scrollAndHighlight(String XPath, WebElement we) { try { //
	 * Scroll to the element ((JavascriptExecutor)
	 * WebHelper.currentdriver).executeScript(
	 * "arguments[0].scrollIntoViewIfNeeded(true);", we);
	 * 
	 * // Highlight the element by changing its border or background color
	 * ((JavascriptExecutor) WebHelper.currentdriver).
	 * executeScript("arguments[0].style.border = '2px solid red';", we);
	 * 
	 * // Wait for a moment to keep the highlight effect visible (you can adjust the
	 * // duration) Thread.sleep(1000); // 1 second
	 * log.info("Claim no highlighted"); WebHelperUtil.saveScreenShot();// 20/10/23
	 * 
	 * // Reset the element's style to remove the highlight ((JavascriptExecutor)
	 * WebHelper.currentdriver).executeScript("arguments[0].style.border = '';",
	 * we); } catch (Exception e) { // Handle any exceptions } }
	 */

	public static void scrollAndHighlight(String XPath, WebElement we) {
		try {

			((JavascriptExecutor) WebHelper.currentdriver)
					.executeScript("arguments[0].scrollIntoView({block: 'center'});", we);

			// Highlight the element by changing its border or background color
			((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].style.border = '2px solid red';",
					we);
			Thread.sleep(1000); // 1 second
			log.info("Claim no highlighted");
			WebHelperUtil.saveScreenShot();// 20/10/23
			// Reset the element's style to remove the highlight
			((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].style.border = '';", we);
		} catch (ElementNotInteractableException e) {
			// Handle ElementNotInteractableException (or other specific exceptions) here
			log.error("Error scrolling and highlighting: " + e.getMessage());
			// You can take further action, like capturing a screenshot
			WebHelperUtil.saveScreenShot();
		} catch (Exception e) {
			// Handle any other exceptions here
			log.error("Error in scrollAndHighlight: " + e.getMessage());
		}
	}

	public static String getDataFetchingSQLQuery(String sqlQuery, String ctrlValue) throws IOException {

		if (ctrlValue == null) {
			// Handle null ctrlValue here, for example, you can set it to an empty string or
			// throw an exception
			// For now, let's set it to an empty string
			ctrlValue = "";
			return ctrlValue;
		}
		if (sqlQuery.contains("$value2")) {
			String[] arrTempVal = ctrlValue.split(",");
			sqlQuery = sqlQuery.replace("$value1", arrTempVal[0].toString());
			sqlQuery = sqlQuery.replace("$value2", arrTempVal[1].toString());
		} else {
			sqlQuery = sqlQuery.replace("$value", ctrlValue);
		}
		return sqlQuery;
	}

	public static String AddTodayDate(String sqlQuery) { // Generate and adds Todays date to SQL Query
		DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
		Date date = new Date();
		String todayDate = dateFormat.format(date);
		log.info(todayDate);
		sqlQuery = sqlQuery.replace("$datevalue", todayDate);
		return sqlQuery;
	}

	public static String AddRandomCodeGenerator(String sqlQuery) { // Generate and adds Random 7 digit String to SQL
																	// Query
		String SALTCHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890";
		StringBuilder salt = new StringBuilder();
		Random rnd = new Random();
		while (salt.length() < 7) { // length of
			// the
			// random
			// string.
			int index = (int) (rnd.nextFloat() * SALTCHARS.length());
			salt.append(SALTCHARS.charAt(index));
		}
		String saltStr = salt.toString();
		sqlQuery = sqlQuery.replace("$randomvalue", saltStr);
		return sqlQuery;
	}

	static String ReadFromExcel(String controlValue) throws IOException {
		return WebHelperUtil.ReadFromExcel(controlValue, WebHelper.columnName);
	}

	public static void claimLoaderWait() throws InterruptedException {
		// Thread.sleep(1500);
		WebHelper.wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
		// Thread.sleep(500);
	}

	public static void GetCellInfo(String FilePath, Row rowValues, int valuesRowIndex, int valuesRowCount) {
		String logicalName = "";
		String controltype = "";
		boolean reportWrittenAfterException = false;
		try {
			WebHelper.TIFilePath = FilePath;
			if (webDriver.getReport().getDriver() != null) {
				WebHelper.currentdriver = webDriver.getReport().getDriver();
			}
			WebHelperClaims.implementWait();

			// WebHelper.frmDate = new Date();
			log.info("********************************" + controller.controllerTestCaseID.toString()
					+ " start DateTime is " + WebHelper.frmDate);
			WebHelper.isDynamicNumFound = true;

			InputStream myXls = new FileInputStream(FilePath);
			Workbook workBook = WebHelperUtil.createWorkbook(myXls, FilePath);
			WebHelper.format = workBook.createDataFormat();

			Sheet sheetStructure = null;
			sheetStructure = workBook.getSheet("Structure");

			int rowCount1 = sheetStructure.getLastRowNum() + 1;
			Sheet headerValues = ExcelUtility.GetSheet(FilePath, "Values");
			// log.info(Automation.dtFormat.format(frmDate));
			// String fromDate = Config.dtFormat.format(WebHelper.frmDate);
			WebHelper.Config_endComparison = Config.endComparison;
			// webDriver.getReport().setFromDate(fromDate);
			WebHelper.structureHeader = WebHelperUtil.getValueFromHashMap(sheetStructure);
			WebHelper.columnName = null;
			int dynamicIndexNumber;// Added for Action Loop
			String imageType, indexVal, controlName, executeFlag, action, controlID, dynamicIndex, newDynamicIndex,
					rowNo, colNo;// newly
			webDriver.getReport().setMessage("");
			webDriver.getReport().setStatus("PASS");
			// log.info("in method GetCellInfo 2");
			// log.info("MainController.pauseExecution:"+MainController.pauseExecution);
			for (int rowIndex1 = 1; rowIndex1 < rowCount1 && !controller.pauseExecution; rowIndex1++) {
				// structureRow = sheetStructure.getRow(rowIndex);
				controlName = WebHelperUtil.getCellData("ControlName", sheetStructure, rowIndex1,
						WebHelper.structureHeader);// structureRow.getCell(Duration.ofSeconds(3));
				executeFlag = WebHelperUtil.getCellData("ExecuteFlag", sheetStructure, rowIndex1,
						WebHelper.structureHeader);// structureRow.getCell(0);

				if (executeFlag.toString().equals("Y")) {
					imageType = WebHelperUtil.getCellData("ImageType", sheetStructure, rowIndex1,
							WebHelper.structureHeader);
					action = WebHelperUtil.getCellData("Action", sheetStructure, rowIndex1, WebHelper.structureHeader);// structureRow.getCell(1);
					logicalName = WebHelperUtil.getCellData("LogicalName", sheetStructure, rowIndex1,
							WebHelper.structureHeader);// structureRow.getCell(Duration.ofSeconds(2));
					controltype = WebHelperUtil.getCellData("ControlType", sheetStructure, rowIndex1,
							WebHelper.structureHeader);// structureRow.getCell(4);
					controlID = WebHelperUtil.getCellData("ControlID", sheetStructure, rowIndex1,
							WebHelper.structureHeader);// structureRow.getCell(6);
					indexVal = WebHelperUtil.getCellData("Index", sheetStructure, rowIndex1, WebHelper.structureHeader);// structureRow.getCell(Duration.ofSeconds(7));
					WebHelper.columnName = WebHelperUtil.getCellData("ColumnName", sheetStructure, rowIndex1,
							WebHelper.structureHeader);
					rowNo = WebHelperUtil.getCellData("RowNo", sheetStructure, rowIndex1, WebHelper.structureHeader);
					colNo = WebHelperUtil.getCellData("ColumnNo", sheetStructure, rowIndex1, WebHelper.structureHeader);
					dynamicIndex = WebHelperUtil.getCellData("DynamicIndex", sheetStructure, rowIndex1,
							WebHelper.structureHeader);// Added

					// *** Mandar -- For the below given code in GPbilling its
					// "FinalSubmitButton" -- need to verify ***
					if (logicalName.equalsIgnoreCase("PolicyNo")) {
						log.info("stop");
					}
					if (logicalName.equalsIgnoreCase("FinalSubmitButton")) {
						log.info("stop");
					}
					if (action.equalsIgnoreCase("LOOP")) {
						WebHelper.loopRow = rowIndex1 + 1;
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
					calldoAction(headerValues, logicalName, rowValues, controller.controllerTransactionType.toString(),
							valuesRowIndex, action, controlName, sheetStructure, rowIndex1,
							controller.controllerTestCaseID.toString(), controltype, controlID, indexVal, imageType,
							FilePath, rowCount1, rowNo, colNo, TransactionMapping.operationType);

					if (action == "END_LOOP" && (valuesRowCount != valuesRowIndex)) {
						WebHelper.loopRow = 1;
						break;
					}
				} else {
					log.info("ExecuteFlag is N");
				}
			}
			Date toDate = new Date();
			webDriver.getReport().setFromDate(Config.dtFormat.format(WebHelper.frmDate));
			webDriver.getReport().setIteration(Config.cycleNumber);
			webDriver.getReport().setTestcaseId(controller.controllerTestCaseID.toString());
			webDriver.getReport().setGroupName(controller.controllerGroupName.toString());
			webDriver.getReport().setTrasactionType(controller.controllerTransactionType.toString());
			webDriver.getReport().setToDate(Config.dtFormat.format(toDate));

			// Setting status for field verification failures
			if (WebHelper.fieldVerFailCount > 0) {
				webDriver.getReport().setMessage("Check Detailed Results");
				webDriver.getReport().setStatus("FAIL");
			}

			log.info("Completed Transaction : " + WebHelper.transactionType.toString());
		} catch (Exception e) {
			if (!controltype.equalsIgnoreCase("OpenAPI")) {

				String errorMessage = e.getMessage();
				String firstLine = errorMessage.split("\\r?\\n")[0];
				log.error("Error :" + firstLine);
				// log.error(e.getMessage(), e);
				controller.pauseFun("Error : LogicalName: " + logicalName + ".\n" + e.getMessage());
			}
			// log.error(e.getMessage(), e);//Sel4
			// controller.pauseFun( logicalName + ".\n" + e.getMessage());//Sel4
			// controller.pauseFun("FieldName Or LogicalName: " + logicalName + ".\n" +
			// e.getMessage());//Sel4
			webDriver.report.setMessage(e.getMessage());
			webDriver.report.setStatus("FAIL");
			reportWrittenAfterException = true;
		} finally {
			WebHelper.structureHeader.clear();
			WebHelper.valuesHeader.clear();

			if (webDriver.getReport().getStatus().equalsIgnoreCase("FAIL")
					&& StringUtils.equalsIgnoreCase(controltype, "OpenAPI")) {

				controller.pauseFun(webDriver.getReport().getMessage());

				// callRecoveryHandler();

				/*
				 * try { ExcelUtility.writeReportPAS(webDriver.getReport()); } catch
				 * (IOException e) { // TODO Auto-generated catch block e.printStackTrace(); }
				 */

				reportWrittenAfterException = true;
			}
			if (StringUtils.equalsIgnoreCase(Config.verificationFailureHandling, "false")
					&& !reportWrittenAfterException && webDriver.getReport().getStatus().equalsIgnoreCase("FAIL")) {
				controller.pauseFun(webDriver.getReport().getMessage());
				reportWrittenAfterException = true;
			}
			// Below given code added to support Claims VV failure - 12/05/2023
			if (StringUtils.equalsIgnoreCase(Config.verificationFailureHandling, "true") && !reportWrittenAfterException
					&& webDriver.getReport().getStatus().equalsIgnoreCase("FAIL")) {
				// controller.pauseFun(webDriver.getReport().getMessage());
				reportWrittenAfterException = true;
			}

			if (!reportWrittenAfterException) {
				try {
					// ExcelUtility.writeReport(webDriver.getReport());
					ExcelUtility.writeReportPAS(webDriver.getReport());
				} catch (Exception e) {
					log.error("Failed while executing ExcelUtility.writeReportPAS in getCellInfo <-|-> LocalizeMessage "
							+ e.getLocalizedMessage() + " <-|-> Message " + e.getMessage() + " <-|-> Cause "
							+ e.getCause(), e);
				}
			}
			WebHelper.fieldVerFailCount = 0;
		}

	}

	public static void callRecoveryHandler() {
		// Create an instance of Maincontrollerclaims
		MainControllerClaims mainController = new MainControllerClaims();

		// Call the recoveryhandler method on the instance
		mainController.recoveryhandler();
	}

	private static void resetVariables() {
		// Resetting or releasing variables
		// VerifyDataFromDB
		sqlQuery = "";
		readFromColName = "";
		toBeFetchedDBColName = "";
		verifyColNameExpected = "";
		// Database/UpdateDB
		writeToColName = "";
		log.info("All Variables reset sucessfully");
	}

	/*
	 * // Method to check if an element is present private static boolean
	 * isElementPresent(WebElement element) { try { element.getTagName(); // Check
	 * if the element is still present return true; } catch
	 * (StaleElementReferenceException | NoSuchElementException e) { return false; }
	 * catch (Exception e) { log.error("Error checking element presence: " +
	 * e.getMessage()); return false; } }
	 */
	
}	
