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
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Random;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.swing.event.PopupMenuEvent;
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
import org.openqa.selenium.WebDriver;
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
import com.majesco.itaf.util.DSProduct;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.util.JDBCConnection;
import com.majesco.itaf.verification.WebVerification;
import com.majesco.itaf.vo.Reporter;

@SuppressWarnings("unused")
public class WebHelperDS {

	private static final String DATA_SHEET = "DataSheet";
	private static final String RESULT_SHEET = "D:\\Workspace\\iTAF-1.20.11.0\\DevStudio\\RegressionTesting\\Resources\\Results\\ResultsSample.xlsx";
	private static final String WRITE_TO_COL_NAME = "WriteToColName";
	private static final String VERIFY_DS_DB = "Verify_DS_DB";
	public static String claimsloader = "//div[@class='overlay']/div[@class='logo-wrapper']/div";
	public static String CRMloader = "//div[@class='spinner-three-bounce full-screen']";
	public static Reporter report = new Reporter();
	private final static Logger log = LogManager.getLogger(WebHelperDS.class.getName());
	public static Cell WebservicecycleDate = null;
	public static String wscycledate;
	public static Boolean blank = false;
	private static HashMap<String, Object> maincontrolsheet = new HashMap<String, Object>();
	private static Row currRow = null;
	public static boolean recovery_done = false;
	private static int current_SC_NO = 0;
	private static int currentscenario_num = 0;
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
	static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	static MainController controller = ObjectFactory.getMainController();
	static void implementWait() throws InterruptedException {
		WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(90));

		FluentWait<WebDriver> wait = new FluentWait<>(WebHelper.currentdriver)
				.withTimeout(Duration.ofSeconds(Integer.parseInt(Config.timeOut)))
				.pollingEvery(Duration.ofSeconds(5))
				.ignoring(NoSuchElementException.class)
				.ignoring(ElementNotInteractableException.class)
				.ignoring(ElementNotInteractableException.class);
		
				//.ignoring(ElementNotFoundException.class); Sel4
	}

	public static void calldoAction2(Sheet headerValues, String logicalName, Row rowValues, String transactionType,
			int valuesRowIndex, String action, String controlName, Sheet sheetStructure, int rowIndex,
			String TestCaseID, String controltype, String controlID, String indexVal, String imageType, String FilePath,
			int rowCount, String rowNo, String colNo, String operationType, final VerifyDB verifyDb,
			boolean executeFlag) throws Exception {
		if (WebHelper.valuesHeader.isEmpty() == true) {
			WebHelper.valuesHeader = WebHelperUtil.getValueFromHashMap(headerValues);
			WebHelper.transactionType = rowValues
					.getCell(Integer.parseInt(WebHelper.valuesHeader.get("TransactionType").toString()));
		}

		if (VERIFY_DS_DB.equals(transactionType)) {

			populateVerifyDb(verifyDb, logicalName, headerValues, valuesRowIndex);
			verifyDb.setTestCaseId(TestCaseID);
			if (WRITE_TO_COL_NAME.equalsIgnoreCase(logicalName)) {
				// Execute Query on the database
				final List<QueryData> myData = WebHelperDS.queryDatabase(verifyDb);
				// Write the results of the Query to an excel file.
				writeResultToSheet(myData, valuesRowIndex);
				verifyDb.initialise();
			}

		} else {
			calldoAction(headerValues, logicalName, rowValues, controller.controllerTransactionType.toString(),
					valuesRowIndex, action, controlName, sheetStructure, rowIndex,
					controller.controllerTestCaseID.toString(), controltype, controlID, indexVal, imageType, FilePath,
					rowCount, rowNo, colNo, TransactionMapping.operationType);
		}
	}

	private static void writeResultToSheet(final List<QueryData> myData, final int rowIndex) {
		// TODO Auto-generated method stub
		FileInputStream in = null;
		FileOutputStream out = null;
		Workbook wk = null;
		Row rowHeader = null, rowValue = null;
		String filePath = RESULT_SHEET;

		try {
			int rowNum = ExcelUtility.getUsedRowCount(RESULT_SHEET, DATA_SHEET);
			String[] values = new String[myData.size()];
			for (int i = 0; i < myData.size(); i++) {
				values[i] = myData.get(i).getColName();
			}
			// To write header
			ExcelUtility.writeRowToExcel(RESULT_SHEET, DATA_SHEET, rowNum + 1, 0, values);
			for (int i = 0; i < myData.size(); i++) {
				values[i] = myData.get(i).getColVal();
			}
			// To write values.
			ExcelUtility.writeRowToExcel(RESULT_SHEET, DATA_SHEET, rowNum + 2, 0, values);

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	/**
	 * @param verifyDb
	 * @throws Exception
	 * @throws ClassNotFoundException
	 * @throws SQLException
	 */
	private static List<QueryData> queryDatabase(final VerifyDB verifyDb) throws ClassNotFoundException, SQLException {
		String colVal;
		List<QueryData> myData = new ArrayList<>();
		myData.add(QueryData.newInstance(VerifyDB.TEST_CASE_ID, verifyDb.getTestCaseId()));

		try (Connection conn = JDBCConnection.establishDSConnection();
				Statement st = conn.createStatement();
				ResultSet rs = st.executeQuery(verifyDb.getSqlQuery())) {
			if (rs.next()) {
				Object objVal = null;
				ResultSetMetaData metaData = rs.getMetaData();
				int colCount = metaData.getColumnCount();

				String val = "";
				for (int i = 1; i < colCount + 1; i++) {
					objVal = rs.getObject(i);
					if (null != objVal) {
						// val = o.toString();
						// DSProduct.writeDataToUniqueNumberSheet(val,metaData.getColumnName(i));
						// val=val + " for column " + metaData.getColumnName(i);
						System.out.println("Value From DSDB Column = " + objVal.toString());
						myData.add(QueryData.newInstance(metaData.getColumnName(i), objVal.toString()));
						// writeResultToSheet();

					}
				}
			}

		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return myData;

	}

	/* DK */
	/**
	 * Will read value from Verify_DB.xls and populate VerifyDB instance.
	 * 
	 * @param verifyDb
	 * @param logicalName
	 * @param headerValues
	 * @param valuesRowIndex
	 * @throws IOException
	 */
	private static void populateVerifyDb(final VerifyDB verifyDb, String logicalName, Sheet headerValues,
			int valuesRowIndex) throws IOException {
		String ctrlValue = WebHelperUtil.getCellData(logicalName, headerValues, valuesRowIndex, WebHelper.valuesHeader);
		verifyDb.populateAttribute(logicalName, ctrlValue);
	}

	public static void calldoAction(Sheet headerValues, String logicalName, Row rowValues, String TransactionType,
			int valuesRowIndex, String action, String controlName, Sheet sheetStructure, int rowIndex,
			String TestCaseID, String controltype, String controlID, String indexVal, String imageType, String FilePath,
			int rowCount, String rowNo, String colNo, String operationType) throws Exception {
		String ctrlValue1 = null;
		String ctrlValue2 = null;
		String ctrlValue = null;
		WebElement webElement = null;
		List<WebElement> controlList = null;

		colnotfound = false;
		if (WebHelper.valuesHeader.isEmpty() == true) {
			WebHelper.valuesHeader = WebHelperUtil.getValueFromHashMap(headerValues);
		}
		Object actualValue = null;

		if (logicalName != null) {
			actualValue = WebHelper.valuesHeader.get(logicalName.toString());
		}

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
		String stransactionType = TransactionType.toString();

		if (colnotfound == false) {
			ctrlValue = WebHelperUtil.getCellData(logicalName, headerValues, valuesRowIndex, WebHelper.valuesHeader);
		} else {
			ctrlValue = "";
		}

		// bhaskar remove leading and trailing whitespaces from values sheet
		// data START
		Pattern trimregex = Pattern.compile("^\\s+|\\s+$");
		Matcher match = trimregex.matcher(ctrlValue);
		StringBuffer ctrlValue_output = new StringBuffer();
		while (match.find())
			match.appendReplacement(ctrlValue_output, "");
		match.appendTail(ctrlValue_output);
		// bhaskar remove leading and trailing whitespaces from values sheet
		// data END

		// bhaskar Action CAPTURE keyword START
		if (action.equalsIgnoreCase("Capture")) {
			// Reporter report = new Reporter();
			log.info("Inside Capture Case");
			controlName = WebHelperUtil.getCellData("ControlName", sheetStructure, rowIndex, WebHelper.structureHeader);
			logicalName = WebHelperUtil.getCellData("LogicalName", sheetStructure, rowIndex, WebHelper.structureHeader);
			if (ctrlValue.equalsIgnoreCase("Y")) {
				TransactionMapping.TransactionCaptureData("NA", TestCaseID, controlName,
						Config.transactionInputFilePath);
			}
		}
		// bhaskar Action CAPTURE keyword END

		if (((action.equals("I") && !StringUtils.isEmpty(ctrlValue))
				|| (action.equals("V_attributeValue") && !StringUtils.isEmpty(ctrlValue))
				|| (action.equals("T") && !StringUtils.isEmpty(ctrlValue))
				|| (action.equals("V") && !StringUtils.isEmpty(ctrlValue)) || !action.equals("I"))
				&& !action.equalsIgnoreCase("Capture") && !action.equalsIgnoreCase("FIND")
				&& !action.equalsIgnoreCase("TABLEINPUT")) {
			if (logicalName.equalsIgnoreCase("WAIT")) {
				log.info(action + " " + logicalName + " For " + controlName);
			} else {
				log.info(action + " On " + controltype + " " + logicalName);
			}

			if (logicalName.equalsIgnoreCase("CreateBatch")) {
				System.out.println("wait");
			}

			if (!controltype.startsWith("Sikuli")) {
				if (!action.equalsIgnoreCase("V") && !action.equals("V_attributeValue")
						&& !action.equalsIgnoreCase("LOOP") && !controltype.equalsIgnoreCase("Wait")
						&& !controltype.equalsIgnoreCase("Wait_IfValue") && !action.equalsIgnoreCase("END_LOOP")
						&& !controltype.equalsIgnoreCase("Browser") && !controltype.equalsIgnoreCase("NewBrowser")
						&& !controltype.equalsIgnoreCase("CloseBrowser")
						&& !controltype.equalsIgnoreCase("CloseAndLaunchNewBrowser")
						&& !controltype.equalsIgnoreCase("Window") && !controltype.equalsIgnoreCase("Alert")
						&& !controltype.equalsIgnoreCase("URL") && !controltype.equalsIgnoreCase("WaitForJS")
						&& !controltype.contains("Robot") && !controltype.equalsIgnoreCase("Calendar")
						&& !controltype.equalsIgnoreCase("CalendarNew") && !controltype.equalsIgnoreCase("CalendarIPF")
						&& !controltype.equalsIgnoreCase("CalendarEBP")
						&& (!action.equalsIgnoreCase("Read")
								|| ((action.equalsIgnoreCase("Read") && !StringUtils.isEmpty(controlName))))
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
							webElement = getElementByType(controlID, controlName, WebHelper.control, imageType,
									ctrlValue);
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

						default:
							break;
						}

						if (elementexists == true) {
							try {
								webElement = getElementByType(controlID, controlName, WebHelper.control, imageType,
										ctrlValue);
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
										controlName, WebHelper.control, ctrlValue); // ,
																					// ISelenium
																					// selenium)
							} catch (NoSuchElementException nse) {
								expMessage = "Failed to find Elements using FindBy with index. " + nse.getMessage();
								StartRecoveryClaims.initiateRecovery(expMessage);
								throw new Exception(expMessage);
							} catch (StaleElementReferenceException sere) {
								expMessage = "Element is no longer appearing on the DOM page with index. "
										+ sere.getMessage();
								StartRecoveryClaims.initiateRecovery(expMessage);
								throw new Exception(expMessage);
							} // catch (ElementNotVisibleException env) //Sel4
							catch (ElementNotInteractableException env) {
								expMessage = "Element is not visible with index. " + env.getMessage();
								StartRecoveryClaims.initiateRecovery(expMessage);
								throw new Exception(expMessage);
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

		/*** Perform action on the identified control ***/
		// log.info("go to method doAction");
		if (!action.equalsIgnoreCase("Capture")) {
			doAction(FilePath, rowValues, testCase, imageType, controltype, controlID, controlName, ctrlValue,
					ctrlValue1, ctrlValue2, wscycledate, logicalName, action, webElement, true, sheetStructure,
					headerValues, rowIndex, rowCount, rowNo, colNo, operationType, "NA", TransactionType);
		}
	}

	// Pradeep- Adding "WriteToDetailResults" for PDF Comparision from
	// WebHelperPAS
	public static Reporter WriteToDetailResults(String expectedValue, String actualValue, String columnName)
			throws IOException

	{
		if (WebHelper.file.exists() == true && DResult == 1) {
			WebHelper.file.delete();
		}
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
			System.out.println(TotalpassCount);
		} else {
			report.setActualValue("FAIL|" + actualValue + "|" + expectedValue);
			System.out.println("ab");
			report.setExpectedValue(expectedValue);
			report.setStatus("FAIL");
			report.setToDate(Config.dtFormat.format(WebHelper.frmDate));
			failCount = "1";
			passCount = "0";
			WebHelper.fieldVerFailCount += 1;
			TotalfailCount += 1;
			System.out.println(TotalfailCount);
		}

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

	// *****END**********

	/** Locating Web Element **/
	public static WebElement getElementByType(String controlId, String controlName, String controlType,
			String imageType, String controlValue) throws Exception {

		WebElement controlList = null;
		String expMessage;
		try {

			if (controlId.equalsIgnoreCase("Id") || controlId.equalsIgnoreCase("HTMLID")) {
				controlList = WebHelper.wait.until(ExpectedConditions.elementToBeClickable(By.id(controlName)));
			} else if (controlId.equalsIgnoreCase("XPath")) {
				controlList = WebHelper.wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
			} else if (controlId.equalsIgnoreCase("XPath_WithBlankValue")) // Pradeep-
																			// Added
																			// condition
																			// to
																			// skip
																			// element
																			// having
																			// blank
																			// in
																			// value
																			// sheet
			{
				if (!controlValue.equalsIgnoreCase("")) {
					controlList = WebHelper.wait
							.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
				} else {
					System.out.println("Blank value found");
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
			} else if (controlId.equalsIgnoreCase("AjaxPath_Dynamic")) {
				controlList = WebHelper.wait.until(ExpectedConditions
						.elementToBeClickable(By.xpath(controlName + "[contains(text(),'" + controlValue)));
			} else if (controlId.equalsIgnoreCase("Id_p") || controlId.equalsIgnoreCase("HTMLID_p")) {
				controlList = WebHelper.wait.until(ExpectedConditions.presenceOfElementLocated(By.id(controlName)));
			} else if (controlId.equalsIgnoreCase("XPath_p")) {
				controlList = WebHelper.wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
			}

			else if (controlId.equalsIgnoreCase("Xpath_Dynamic"))// Pradeep:
			// Added controlId to handle dynamic Xpath
			{
				controlList = WebHelper.wait.until(ExpectedConditions.elementToBeClickable(
						By.xpath("//div[contains(text(),'" + controlValue + "')]" + controlName)));

			}

			else if (controlId.equalsIgnoreCase("XPathValue")) // Pradeep: Added

			// condition to replace control value in xpath.
			{
				if (!controlValue.equalsIgnoreCase("")) {
					if (controlValue.trim().equals("pickFromUniqueNumbers")) {
						controlValue = WebHelperUtil.ReadFromExcel(controlValue, WebHelper.columnName);
					}
					String tempCtrlName = controlName;
					String tempReplaceString = tempCtrlName.replace("$Value", controlValue);
					controlList = WebHelper.wait
							.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString)));
				} else {
					System.out.println("No Value to select element");
				}
			}

			else if (controlId.equalsIgnoreCase("XPath2$Value")) // Pradeep:
																	// Added
																	// condition
																	// to
																	// replace
																	// TWO
																	// control
																	// value in
																	// xpath.12/26/2019
			{
				if (!controlValue.equalsIgnoreCase("")) {
					if (controlValue.contains(";")) {
						String[] MultipleControlValue = controlValue.split(";");

						String tempControlname1 = controlName.replace("$Value1", MultipleControlValue[0]);
						String tempControlname2 = tempControlname1.replace("$Value2", MultipleControlValue[1]);
						WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(60));
						WaitForPageLoad.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
						controlList = WebHelper.wait
								.until(ExpectedConditions.elementToBeClickable(By.xpath(tempControlname2)));
					}

				} else {
					System.out.println("No Value to select element");
				}
			}

			// [DivyaK] Commenting the claims code
			else if (controlId.equalsIgnoreCase("XPathClaimNo")) // Pradeep:
			// Added condition to replace ClaimNo in xpath.
			{
				String tempCtrlName = controlName;
				String uniqueNumber = ReadFromExcel(controlValue);
				String tempReplaceString = tempCtrlName.replace("$Value", uniqueNumber);
				controlList = WebHelper.wait
						.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString)));
			}

			else if (controlId.equalsIgnoreCase("XPath_R"))// Asif : Added
			// 31-May :Start
			{
				if (!controlValue.isEmpty()) {

					try {

						controlList = WebHelper.wait
								.until(ExpectedConditions.elementToBeClickable(By.xpath(controlValue)));

					} catch (Exception e) {
						log.error(e.getMessage(), e);
						controlList = null;
					}
				}
			} // Asif : Added 31-May :End

			else if (controlId.equalsIgnoreCase("XPath_if"))// Asif : Added
			// 31-May :Start
			{
				WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(3));
				try {

					if (wait1.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName))).isDisplayed()) {
						controlList = wait1.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
					}
				} catch (Exception e) {
					log.error(e.getMessage(), e);
					controlList = null;
				}
			} // Asif : Added 31-May :End

			else if (controlId.equalsIgnoreCase("XPath_value"))// Asif : Added
			// 31-May :Start
			{
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
			} // Asif : Added 31-May :End

			// bhaskar removing enum constants as suggested by dharmendra END
			return controlList;
		} catch (NoSuchElementException nse) {
			expMessage = "Failed to find Element(s). " + nse.getMessage();
			StartRecoveryClaims.initiateRecovery(expMessage);
			throw new Exception(expMessage);
		} catch (StaleElementReferenceException sere) {
			expMessage = "Element is no longer appearing on the DOM page. " + sere.getMessage();
			StartRecoveryClaims.initiateRecovery(expMessage);
			throw new Exception(expMessage);
		} // catch (ElementNotVisibleException env) //Sel4
		catch (ElementNotInteractableException env) {
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

		String clocation;
		// log.info("In method doaction");
		List<WebElement> WebElementList = null;
		String currentValue = null;
		String uniqueNumber = "";
		WebVerification.isFromVerification = false;
		Constants.ControlTypeEnum controlTypeEnum = Constants.ControlTypeEnum.valueOf(controlType);
		Constants.ControlTypeEnum actionName = Constants.ControlTypeEnum.valueOf(action.toString());
		// bhaskar
		WebHelper.sikscreen = Config.SikuliScr;
		// log.info(sikscreen);
		// bhaskar
		if (controlType.contains("Robot") && !WebHelper.isIntialized) {
			log.info("In method doaction debug1");
			WebHelper.robot = new Robot();
			WebHelper.isIntialized = true;
		}

		if (!WebHelperUtil.stringIn(action, new String[] { "I", "V", "F", "VA", "VV" })
				|| !ctrlValue.equalsIgnoreCase("")) {

			try {
				switch (controlTypeEnum) {

				case WebEdit_NoScroll:
					webEditNoScroll(ctrlValue, webElement, actionName);
					break;
				case WebEdit:
					switch (actionName) {
					case Read:
						webEdit_Read(ctrlValue, webElement);
						break;
					case Write:
						webEdit_Write(controlType, controlId, controlName, ctrlValue, webElement, rowNo, colNo);
						break;

					case I:
						ctrlValue = webEdit_I(controlType, controlId, controlName, ctrlValue, logicalName, webElement,
								rowNo, colNo, TransactionType);

						break;
					case V:

						break;

					case V_TextImage:
						currentValue = webEdit_V_TextImage(controlName, ctrlValue, currentValue);
						break;

					case V_attributeValue:
						if (StringUtils.equalsIgnoreCase(ctrlValue, null) || ctrlValue.trim().equals("")
								|| ctrlValue.equalsIgnoreCase("IGNORE")) {
							break;
						}
						if (ctrlValue.equalsIgnoreCase("pickFromUniqueNumbers")) {
							ctrlValue = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
						}
						currentValue = webEdit_V_attributeValue(imageType, ctrlValue, webElement, currentValue);
						break;
					case VF:
						break;

					case VV:
						break;
					default:
						break;
					}
					break;

				case WebButton:
					switch (actionName) {
					case I:
						webButton_Case_I(controlName, ctrlValue, logicalName, webElement, TransactionType);
						break;

					case NC:
						webButton_Case_NC(ctrlValue, webElement);
						break;

					case C: // Pradeep - Direct click on button without
						// required any Input from InputSheet

						webButton_Case_C(webElement);

						break;

					case ReadD: // Pradeep-Added case to handle dynamic
						// Xpath replace with UniqueNumber
						if (ctrlValue.equalsIgnoreCase("Y")) {
							uniqueNumber = ReadFromExcel(ctrlValue);
							String controlNameNew = controlName.replace("$Value", uniqueNumber);
							WebElement webElementNew = WebHelper.wait
									.until(ExpectedConditions.elementToBeClickable(By.xpath(controlNameNew)));
							System.out.println(webElementNew);
							WebDriverWait WaitForPageLoadNew = new WebDriverWait(Automation.driver,
									Duration.ofSeconds(120));
							WaitForPageLoadNew
									.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
							webElementNew.click();
							Thread.sleep(1000);
						} else {
							System.out.println("Control Value found Blank or N ");
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

					case V_Disable: // Pradeep- added condition "V_Disable" to
									// verify Enable/Disable of element.
						WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(5));
						WaitForPageLoad.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
						Verifyelement = null;
						try {
							WebElementList = WebHelper.currentdriver.findElements(By.xpath(controlName));
							if (!ctrlValue.equalsIgnoreCase("")) {
								if (WebElementList.isEmpty()) {

									currentValue = "Disable";
								}
							}

						} catch (Exception e) {
							log.error(e.getMessage(), e);
							currentValue = "Enable";
						}
						break;

					case Read: // Mayur_Claims (WorkBench)
						Reporter report = new Reporter();
						report.setReport(report);

						WebHelper.ActualValue = webElement.getText();
						uniqueNumber = ReadFromExcel(ctrlValue);
						WebHelper.ExpectedValue = uniqueNumber;

						log.info("ActualValue is : " + WebHelper.ActualValue);
						log.info("ExpectedValue is : " + WebHelper.ExpectedValue);
						if (WebHelper.ActualValue.contains(WebHelper.ExpectedValue)) {
							report.setStatus("PASS");
							report.setStatus(report.getStatus());
							report.setMessage("Values Matched");
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

						break;
					case T: // Mayur
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
							// Thread.sleep(1200);
							WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(90));
							// [Divya K] claimLoaderWait();
							// WaitForPageLoad.until(ExpectedConditions.invisibilityOfElementLocated(By.className("overlay
							// hide")));
							WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
							WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
							webElement.click();
						} else {
							System.out.println("Element is not clicked");
						}
						break;
					case NCIF: // to be used with XPath_if controlID
						try {
							if (webElement.isDisplayed()) {
								webElement.click();
								Thread.sleep(500);
							}

						} catch (Exception e) {
							log.error(e.getMessage(), e);
							log.info("Element Not Found on Page in NCIF case");
						}
						break;
					default:
						break;
					}
					break;

				case AjaxWebButton: // Mayur-To Handle AJAX buttons
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
					case V: // Govind- To verification of text-field
						// data.
						if (!(ctrlValue.equalsIgnoreCase(""))) {
							claimLoaderWait();
							WebElement Verifyelement = null;

							Verifyelement = WebHelper.wait
									.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
							// break;

							currentValue = Verifyelement.getAttribute("value");
							Verifyelement = null;
						} else {
							System.out.println("Element not Present");
						}
						break;
					default:
						break;
					}
					break;

				case WebElement:
					// bhaskar
					// log.info("Inside webelement scenario");
					// WebVerification.isFromVerification = true;
					// bhaskar
					switch (actionName) {
					case NC: // Scroll and Click on JAvaScript
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
							uniqueNumber = ReadFromExcel(ctrlValue);

							scroll(controlName, webElement);

							webElement.clear();
							webElement.sendKeys(uniqueNumber);
							/*
							 * ((JavascriptExecutor)currentdriver). executeScript(
							 * "arguments[0].scrollIntoView();", webElement); Thread.sleep(1000);
							 * ((JavascriptExecutor)currentdriver ).executeScript("arguments[0].click();",
							 * webElement);
							 */
						} else {
							System.out.println("Element is not Clicked");
						}
						break;
					case I: // Mayur_Claims
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
							WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
									Duration.ofSeconds(90));
							claimLoaderWait();
							scroll(controlName, webElement);
							WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
							WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));

							Thread.sleep(500);
							webElement.click();
						} else {
							System.out.println("Element is not clicked");
						}
						break;

					case C:
						if (controlId.equalsIgnoreCase("XPath$Value$ClaimNo$") && ctrlValue.equalsIgnoreCase("")) // Pradeep:
						// Added
						// condition
						// to
						// replace
						// ClaimNo
						// and
						// Value
						// in
						// xpath.
						{
							System.out.println("Blank Value Found");
						} else // Pradeep: Added condition to replace
								// ClaimNo in xpath.
						{
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

					case Read: // Mayur_Claims
						uniqueNumber = ReadFromExcel(ctrlValue);
						/*
						 * try{ webElement= wait.until(ExpectedConditions .elementToBeClickable
						 * (By.xpath("//li[@class='last']//a"))); ((JavascriptExecutor
						 * )currentdriver).executeScript( "arguments[0].scrollIntoView();", webElement);
						 * WebDriverWait WaitForPageLoad=new
						 * WebDriverWait(Automation.driver,Duration.ofSeconds(120)); WaitForPageLoad.
						 * until(ExpectedConditions. invisibilityOfElementLocated
						 * (By.xpath(claimsloader))); WaitForPageLoad.until(ExpectedConditions
						 * .visibilityOf(webElement)); WaitForPageLoad.until(ExpectedConditions
						 * .elementToBeClickable(webElement)); webElement.click(); WebElement
						 * webElement2= wait.until (ExpectedConditions.elementToBeClickable(By.
						 * xpath(controlName +"[contains(text(),'"+uniqueNumber+"')]")));
						 * System.out.println(webElement2); WaitForPageLoad.until
						 * (ExpectedConditions.invisibilityOfElementLocated (By.xpath(claimsloader)));
						 * webElement2.click(); Thread.sleep(1000); } catch(Exception e){
						 */
						if (ctrlValue.trim().equalsIgnoreCase("")) { // Mayur_Avoid
																		// Claim
																		// Click
							WebElement webElement2 = WebHelper.wait.until(ExpectedConditions.presenceOfElementLocated(
									By.xpath(controlName + "[contains(text(),'" + uniqueNumber + "')]")));
							System.out.println(webElement2);
							claimLoaderWait();

							scroll(controlName, webElement);

							webElement2.click();
						} else {
							System.out.println("Not needed");
						}

						break;
					case Write:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
							WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName,
									rowNo, colNo);
						} else {
							System.out.println("Not needed");
						}
						break;
					case V:// Mayur_ Verification of Text-fields
						if (!(ctrlValue.equalsIgnoreCase(""))) {

							if (!(ctrlValue.equalsIgnoreCase(""))) {

								if (ctrlValue.contains("<>")) // Pradeep: Added
																// code to get
																// Dynamic XPath
																// value from
																// Inputdata to
																// verify table
																// data.12/26/2019
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
									// currentValue =
									// Verifyelement.getAttribute(tempControlname1.toString());
									currentValue = Verifyelement.getText();
									log.info("ActualValue is : " + currentValue);
									log.info("ExpectedValue is : " + MultipleControlValue[1]);
									// Verifyelement = null;
									ctrlValue = MultipleControlValue[1];

								}

								// Pradeep - Added case to verify randaom value
								// starting with "M" , "Null" and Random
								// Numbers, TextVerify, Todays date, Count
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
										System.out.println(ExpectedDate);

										WebElement AutualElement = WebHelper.wait
												.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
										// String ActualDate =
										// AutualElement.getAttribute("value");
										String ActualDate = AutualElement.getText();

										if (ExpectedDate.equalsIgnoreCase(ActualDate)) {

											System.out.println("Date Matched");
											currentValue = "Today";
										} else {
											System.out.println("Date Not Matched");
											currentValue = ActualDate;
										}
										break;

									}

									else if (logicalName.contains("TextVerify")) {
										WebElement ActualElement = WebHelper.currentdriver
												.findElement(By.xpath(controlName));
										// Modified by Shubhashree-Uncommented
										// to fetch the attribute property
										String ActualValue1 = ActualElement.getAttribute("value");
										String ActualValue = ActualElement.getText();
										if (ActualValue.contains(ctrlValue)) {
											System.out.println("Text Matched");
											currentValue = ctrlValue;
										}

										else if (ActualValue1.equals(ctrlValue)) {
											System.out.println("The text matches with the orginal value");
											currentValue = ctrlValue;
										} else {
											System.out.println("Text Not Matched");
											currentValue = ActualValue;
										}
										break;
									}

									else if (logicalName.contains("Count"))

									{
										List<WebElement> CurrentElement = WebHelper.currentdriver
												.findElements(By.xpath(controlName));
										// String currentValue1 =
										// CurrentElement.getAttribute("Value");
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
								System.out.println("Element not Present");
							}
						}
						break;

					/*
					 * case V_ReadOnly:// Pradeep-Added case to verify read only
					 * 
					 * try { if (ctrlValue.equalsIgnoreCase("ReadOnly")) { WebElement webElement1 =
					 * WebHelper.currentdriver.findElement(By.xpath(controlName)); String
					 * Verifyelement1 = webElement1.getAttribute("readonly"); if
					 * (Verifyelement1.equals(true)) { currentValue = "ReadOnly"; } else {
					 * currentValue = "Enable"; } } else {
					 * System.out.println("No control value to be verified"); } } catch (Exception
					 * e) { log.error(e.getMessage(), e); currentValue = "False"; } break;
					 *///Sel4
						
					case V_ReadOnly:// Pradeep-Added case to verify read only
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
					            System.out.println("No control value to be verified");
					        }
					    } catch (Exception e) {
					        log.error(e.getMessage(), e);
					        currentValue = "False";
					    }
					    break;	

					// bhaskar Action Verify Values VV START
					case T: // Mayur_Claims (Payment I Icon
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
					case V_NotBlank: // Pradeep - Added case to verify empty
										// field or field with value.12/26/2019.
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
								System.out.println("No control value to be verified");
							}
						} catch (Exception e) {
							log.error(e.getMessage(), e);
							currentValue = "False";
						}
						break;

					case V_ClaimNo: // Pradeep - Added case to verify
									// ClaimNo.12/26/2019.
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
								System.out.println("No value to be verified");
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

				case JSScript:
					((JavascriptExecutor) WebHelper.currentdriver).executeScript(controlName, ctrlValue);
					break;

				case WaitForPageToLoad:// Asif: Added 31-May 2017:Start
					switch (actionName) {
					case I: // Mayur_Claims
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !ctrlValue.trim().equals("")) {
							// Automation.driver.manage().timeouts().pageLoadTimeout(120,
							// TimeUnit.SECONDS);

							WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
									Duration.ofSeconds(300));
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
						// TimeUnit.SECONDS);// Sel4
						Automation.driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(120));

						WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(300));

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
					default:// Added Sel4
						break;
					}
					break;

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
								// Thread.sleep(1000);
								webElement.click(); // Select Checkbox,
								// if it is
								// not selectd
							}
						}

						else if (ctrlValue.equalsIgnoreCase("N") || ctrlValue.equalsIgnoreCase("No")) {
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
								// Thread.sleep(1000);
								webElement.click(); // Deselect
								// Checkbox, if it
								// is selectd
							}
						}
						break;
					case V:
						/*
						 * if (webElement.isSelected()) { //
						 * log.info("logical Name is not Selected:"+logicalName) ;
						 * System.out.println(logicalName + "Is Selected"); // currentValue = //
						 * webElement.getAttribute(controlName.toString()); } else {
						 * System.out.println(logicalName + "Is not Selected"); }
						 */
						if (ctrlValue.contains("FALSE") || ctrlValue.contains("TRUE")) { // Mayur-
																							// Verify
																							// Checkbox
																							// is
																							// selcted
																							// or
																							// Not
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

					case V_EnableDisable:// Pradeep- Added case to verify
											// EnableDisable based on element
											// attribute as "disabled"
						if (ctrlValue.equalsIgnoreCase("") || StringUtils.isEmpty(ctrlValue)) {
							System.out.println("Blank Value Found");
							break;
						} else {
							String ActualValue = webElement.getAttribute("disabled");

							if (ActualValue.equalsIgnoreCase("true")) {
								// String ActualValue =
								// webElement.getAttribute("disabled");
								System.out.println(logicalName + "  Is Disabled");
								System.out.println(ActualValue);
								currentValue = "Disabled";
							} else {
								System.out.println(logicalName + "Element is Enabled");
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
					case Read: // Mayur_Claims for delete/void bulk
						uniqueNumber = ReadFromExcel(ctrlValue);
						try {
							WebElement webElement2 = WebHelper.wait.until(
									ExpectedConditions.elementToBeClickable(By.xpath(controlName + "[contains(text(),'"
											+ uniqueNumber + "')]//following::td[@data-colid='Col_7']/div/div/input")));
							System.out.println(webElement2);
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
							System.out.println(webElement2);
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
					case I: // Mayur_Claims for release bulk
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
									System.out.println(webElement2);
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
									System.out.println(webElement2);
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
									System.out.println(webElement2);
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
									System.out.println(webElement2);
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
									System.out.println(webElement2);
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
									System.out.println(webElement2);
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

					case NC: // Mayur- To verify radio or check box
						// selected or
						// not
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
							if (!webElement.isSelected()) {
								System.out.println("IS NOT SELECTED");
								log.info(logicalName + " is not Selected ");
								// webElement.click();
							} else {
								System.out.println("IS SELECTED");
								log.info(logicalName + " is Selected ");
							}
						} else {
							System.out.println("No action CHECKBOX/ RADIO Button");
						}
						break;
					case T: // Mayur_Select depending on the
						// requirement
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
							if (!webElement.isSelected()) {
								webElement.click();
							} else {
								System.out.println("No radio found");

							}
						} else {
							System.out.println("NO for this Scenario");
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
					default:// Sel4
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
						WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(30));
						claimLoaderWait();
						WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
						WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));

						Thread.sleep(500);
						webElement.click();

						break;
					case V_List:// Pradeep- Added case to verify

						currentValue = new String();
						String[] ExpectedValue;
						String string = ctrlValue;
						String[] ExpectedValue1 = string.split(",");
						for (int i = 1; i < ExpectedValue1.length + 5; i++) {
							String webElement1 = WebHelper.currentdriver.findElements(By.xpath(
									"//span[starts-with(@class,'select2-container select2-container--default select2-container--open')]//span[@class='select2-results']//ul[starts-with(@id,'select2-cmbVendorType-') and contains(@id,'results')]//li"))
									.get(i).getText();

							if (webElement1.equals(ExpectedValue1[i])) {
								System.out.println("MATCHED");
							}

							else {
								System.out.println("Value Not MATCHED");
							}
						}

						break;

					case V_Select:// Pradeep- Added case to verify

						if (!ctrlValue.equalsIgnoreCase("")) {
							String ExpectedValue_String = ctrlValue;
							String[] ExpectedValue2 = ExpectedValue_String.split(",");
							for (int i = 0; i < ExpectedValue2.length; i++) {

								WebElement VendorSearchClick = WebHelper.currentdriver.findElement(By.xpath(
										"//span[starts-with(@aria-labelledby,'select2-cmbVendorType-') and contains(@aria-labelledby,'container')]//span[@class='select2-selection__arrow']/b"));
								VendorSearchClick.click();
								Thread.sleep(1000);
								WebElement VendorSearchSelect = WebHelper.currentdriver.findElement(
										By.xpath(controlName + "[contains(text(),'" + ExpectedValue2[i] + "')]"));
								VendorSearchSelect.click();

								Thread.sleep(1000);
								WebElement webElement1 = WebHelper.currentdriver.findElement(By.xpath(
										"//div[@class='col-md-3']/div[@style='display: block;']//span[starts-with(@id,'select2-cmbAssociatedOrganization-') and contains(@id,'container')]"));

								boolean AssoOrgDropDown = webElement1.isDisplayed();
								if (AssoOrgDropDown == true) {
									System.out
											.println("Associated Organization dropDown found for " + ExpectedValue2[i]);
									String Actualvalue = webElement1.getText();

									if (Actualvalue.equals("Select")) {
										System.out.println("Value Select found");
										currentValue = Actualvalue;
										ctrlValue = ExpectedValue2[i];
									}

									else {
										System.out.println("Select not found");
										currentValue = Actualvalue;
										ctrlValue = ExpectedValue2[i];
									}

								}

								else {
									System.out.println(
											"Associated Organization dropDown not found for " + ExpectedValue2[i]);
								}
							}
						}
						break;

					case V_Disable:// Pradeep- Added case to verify
						// Dropdown value with expected list
						// Thread.sleep(2000);
						if (!ctrlValue.equalsIgnoreCase("")) {
							String ExpectedValue1_String = ctrlValue;
							// Start-Change the logic to take value from
							// controlValue
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
									System.out.println(
											"Associated Organization dropDown not found for " + ExpectedValue3[i]);
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

					case V: // Pradeep: Added case to verify selected value from
							// drop down
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
						System.out.println("Element Entered");
						// webElement.sendKeys(Keys.ARROW_DOWN);
						Thread.sleep(200);
						webElement.sendKeys(Keys.ENTER);
						System.out.println("Enter key pressed");
						Thread.sleep(400);
						break;
					default:// Sel4
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

					case Select:// Pradeep: Added case to select drop down value

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

					case V_List: // Pradeep-Added case to verify DropDown value
									// through select.getOptions() method.
						if (ctrlValue == null || ctrlValue.trim().equals("")) {
							break;
						}

						String[] ExpectedValue1 = ctrlValue.split("<>");
						List<WebElement> AllWebElement = WebHelper.currentdriver.findElements(By.xpath(controlName));
						WebElement dropdown1 = WebHelper.currentdriver.findElement(By.xpath(controlName));

						Select select = new Select(dropdown1);

						List<WebElement> options = select.getOptions();
						int count = 0;
						for (WebElement we : options) {
							for (int i = 0; i < ExpectedValue1.length; i++) {

								if (we.getText().equals(ExpectedValue1[i])) {
									count++;
									// String addArray[]= (String)
									// ExpectedValue1[i];
									System.out.println("Matched");
									System.out.println("Actual Value: " + we.getText());
									System.out.println("Expected Value: " + ExpectedValue1[i]);
								} else {
									System.out.println("Actual Fail Value: " + we.getText());
								}
							}
						}
						if (count == ExpectedValue1.length) {
							System.out.println("Expected Count: " + ExpectedValue1.length);
							System.out.println("Acutal Count: " + count);
							currentValue = "Matched";
							ctrlValue = "Matched";
							break;
						} else {
							System.out.println("Expected Count: " + ExpectedValue1.length);
							System.out.println("Acutal Count: " + count);
							currentValue = "Value Not Matched";
							ctrlValue = "Matched";
						}
						// break;
					default:// Sel4
						break;

					}
					break;

				case TAB: // Pradeep- To handle TAB to focus out from
					// element
					switch (actionName) {
					case NC:
						System.out.println("Webelement is: " + webElement);

						webElement.sendKeys(Keys.TAB);
						System.out.println("Performed TAB out");
						break;

					case I: // Mayur_ TAB optional
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
							System.out.println("Webelement is: " + webElement);

							webElement.sendKeys(Keys.TAB);
							System.out.println("Performed TAB out");

						}
						break;
					default:// Sel4
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
							// condition
							// to
							// take
							// user
							// Input
							// flag
							// as
							// "Y"
							System.out.println("In method doaction debug4");
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

						System.out.println("In method doaction debug5");
						break;

					case NC:
						System.out.println("In method doaction debug4");
						Thread.sleep(500);

						if (controlName.startsWith("//iframe")) {

							WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(700));

							wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

							Automation.driver.switchTo().frame(Automation.driver.findElement(By.xpath(controlName)));

						}

						else {

							Automation.driver.switchTo().frame(controlName);

						}

						System.out.println("In method doaction debug5");
						break;

					case Back: // Pradeep-Added case to handle
						// Switching back to parent window
						System.out.println("Switching back to Parent window");
						Thread.sleep(500);

						Automation.driver.switchTo().defaultContent();

						System.out.println("Done the Switching back to Parent window");
						break;
					default:// Sel4
						break;

					}
					break;

				case Multiselect: // Mayur_ MultiSelect/Deselct for JS
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
							System.out.println("No Data Entered");
						}
						break;

					case T: // Mayur - Deselect multi choice
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
							System.out.println("No Data Entered");
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

				case NewBrowser:// Mayur: Added 31-May :Start
					switch (actionName) {
					case I: // Mayur_Extra verification...
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
					default:// Sel4
						break;
					}
					break;// case NewBrowser://Mayur

				case CloseBrowser:// Mayur
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
						WebHelperDS.implementWait();
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
						WebHelperDS.implementWait();
						break;
					default:// Sel4
						break;
					}
					break;

				case URL:
					switch (actionName) {
					case I:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
							if (ITAFWebDriver.isClaimsApplication()) {
								try {
									WebHelper.currentdriver.navigate().to(controlName);
									WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
											Duration.ofSeconds(5));
									WebElement webElement1 = WaitForPageLoad
											.until(ExpectedConditions.elementToBeClickable(
													By.xpath("//div[contains(@class,'loginPage-majescologo')]")));
								} catch (Exception e) {
									try {
										WebHelper.currentdriver.navigate().refresh();
										WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
												Duration.ofSeconds(5));
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
							if (ITAFWebDriver.isClaimsApplication()) // For CS
							// Enviornment
							{
								WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
										Duration.ofSeconds(7));
								try {
									Automation.driver.get(controlName);
									Thread.sleep(200);
									claimLoaderWait();
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
							System.out.println("Alert found on the web page");
							System.out.println(currentValue);
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
					default:// Sel4
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
								System.out.println("Element not exist");
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
					default:// Sel4
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
					case I:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
							WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
									Duration.ofSeconds(30));
							claimLoaderWait();
							WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
							WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
							Actions builderMouserOver = new Actions(WebHelper.currentdriver);
							// builderMouserOver.moveToElement(webElement).build().perform();
							builderMouserOver.moveToElement(webElement).click().build().perform(); // Pradeep-Click
							// after
							// Mouseover
							// Thread.sleep(1000);
							// Action mouseOverAction =
							// builderMouserOver.moveToElement(webElement).build();
							// mouseOverAction.perform();
						}
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
					case I: // Mayur_ Claims_handle AJAX pop-up
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("WK")
								|| ctrlValue.equalsIgnoreCase("D") || ctrlValue.equalsIgnoreCase("DP")
								|| ctrlValue.equalsIgnoreCase("C") || ctrlValue.equalsIgnoreCase("Ok")
								|| ctrlValue.equalsIgnoreCase("Ok1") || !(ctrlValue.trim().equalsIgnoreCase(""))) {
							claimLoaderWait();
							if (ctrlValue.equalsIgnoreCase("Y")) {
								try {
									webElement = WebHelper.wait.until(ExpectedConditions.elementToBeClickable(
											By.xpath("//button[@name='Ok' or @name='btnOk' or @name='OK']")));// Added
																												// "Ok"
																												// for
																												// updated
																												// popup
									Thread.sleep(500);
									// webElement.click();
									Actions builderClick = new Actions(Automation.driver);
									Action clickAction = builderClick.moveToElement(webElement).clickAndHold().release()
											.build();
									clickAction.perform();
									log.info("Pop-up found on the web page" + "ctrlValue");
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

									log.info("Pop-up found on the web page" + "ctrlValue");
									Thread.sleep(1000);
								} catch (Exception e) {
									log.error(e.getMessage(), e);
									System.out.println("NO duplicate pop up found");
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
									log.error(e.getMessage(), e);
									System.out.println("NO duplicate pop up found");
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
									log.error(e.getMessage(), e);
									System.out.println("No Pop up found");
								}
							} else if ((ctrlValue.equalsIgnoreCase("OKS"))) {// Pradeep-
																				// Added
																				// new
																				// condition
																				// to
																				// handle
																				// new
																				// Xpath
																				// for
																				// OK
																				// button.12/26/2019
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
									log.error(e.getMessage(), e);
									System.out.println("NO duplicate pop up found");
								}
							} else {
								System.out.println("No window Found");
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
							log.error(e.getMessage(), e);
							System.out.println("NO Pop up found");
						}
						break;

					default:
						break;
					}
					break;

				case LoaderWait:// Pradeep- Added case to handle Loader wait
								// condition.12/26/2019.
					switch (actionName) {
					case NC:
						WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(120));
						WaitForPageLoad.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
					default:
						break;
					}
					break;

				case CRMLoaderWait:// Pradeep- Added case to handle CRM Loader
									// wait condition
					switch (actionName) {
					case NC:
						// WebDriverWait WaitForPageLoad = new
						// WebDriverWait(Automation.driver, Duration.ofSeconds(120));
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
				case Database:// Case added to fetch Unique values from Database
					switch (actionName) {
					case CaptureDataFromDB:
						if (logicalName.equalsIgnoreCase("SQLQuery"))
							sqlQuery = ctrlValue;
						else if (logicalName.equalsIgnoreCase("ReadFromColName"))
							readFromColName = ctrlValue;
						else if (logicalName.equalsIgnoreCase("ColumnToBeFetchedFromDB"))
							toBeFetchedDBColName = ctrlValue;
						else if (logicalName.equalsIgnoreCase(WRITE_TO_COL_NAME)) {
							writeToColName = ctrlValue;

							if (readFromColName != "") {
								String tempValue = DSProduct.ReadFromExcelUsingColumnName("", readFromColName);
								sqlQuery = getDataFetchingSQLQuery(sqlQuery, tempValue);
							}

							ResultSet rs3;
							Connection conn2 = JDBCConnection.establishPASDBConn();
							Statement st2 = conn2.createStatement();
							rs3 = st2.executeQuery(sqlQuery);
							rs3.next();
							System.out.println(toBeFetchedDBColName);
							// ArrayList toBeFetchedDBColNameList = new ArrayList();//Sel4
							ArrayList<String> toBeFetchedDBColNameList = new ArrayList<String>();// Sel4
							String toBeFetchedDBColNameParts = "";

							if (toBeFetchedDBColName.contains(","))// ToFetchmultiplerecords
							{
								String[] toBeFetchedDBColNamePart = toBeFetchedDBColName.split(",");

								for (int i = 0; i < toBeFetchedDBColNamePart.length; i++) {
									toBeFetchedDBColNameParts = String
											.valueOf(rs3.getString(toBeFetchedDBColNamePart[i]));
									toBeFetchedDBColNameList.add(toBeFetchedDBColNameParts);
								}
								System.out.println(toBeFetchedDBColNameList);
								String toBeFetchedDBColNameMain = "";
								for (int i = 0; i < toBeFetchedDBColNameList.size(); i++) {
									toBeFetchedDBColNameMain = (String) toBeFetchedDBColNameList.get(i) + ','
											+ toBeFetchedDBColNameMain;
								}
								System.out.println(toBeFetchedDBColNameMain);
								ctrlValue = toBeFetchedDBColNameMain.substring(0, toBeFetchedDBColNameMain.length() - 1)
										+ "";

							} else {
								ctrlValue = String.valueOf(rs3.getString(toBeFetchedDBColName));
							}

							rs3.close();

							DSProduct.writeDataToUniqueNumberSheet(ctrlValue, writeToColName);
						}
						break;

					case VerifyDataFromDB: // Mayur_Verification Fetch data with
											// DB Value

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

							WebHelper.ExpectedValue = DSProduct.ReadFromExcelUsingColumnName("", verifyColNameExpected);

							if (!readFromColName.equalsIgnoreCase("")) {
								String tempValue = DSProduct.ReadFromExcelUsingColumnName("", readFromColName);

								sqlQuery = getDataFetchingSQLQuery(sqlQuery, tempValue);
							}

							ResultSet rs3;
							Connection conn2 = JDBCConnection.establishPASDBConn();
							Statement st2 = conn2.createStatement();
							rs3 = st2.executeQuery(sqlQuery);
							rs3.next();

							WebHelper.ActualValue = String.valueOf(rs3.getString(toBeFetchedDBColName));

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
						}
						break;
					case UpdateDB: // Mayur_ Update Query DB
						if (logicalName.equalsIgnoreCase("SQLQuery"))
							sqlQuery = ctrlValue;
						else if (logicalName.equalsIgnoreCase("ReadFromColName"))
							readFromColName = ctrlValue;
						else if (logicalName.equalsIgnoreCase("ColumnToBeFetchedFromDB"))
							toBeFetchedDBColName = ctrlValue;
						else if (logicalName.equalsIgnoreCase(WRITE_TO_COL_NAME)) {
							writeToColName = ctrlValue;

							if (sqlQuery.contains("'$datevalue'")) { // Add
																		// todays_date
								sqlQuery = AddTodayDate(sqlQuery);
							}

							if (sqlQuery.contains("'$randomvalue'")) { // Add
																		// Random
																		// Check
																		// number
								sqlQuery = AddRandomCodeGenerator(sqlQuery);
							}

							String tempValue = DSProduct.ReadFromExcelUsingColumnName("", readFromColName);
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

					// Below given missing conditions added to avoid warning message - Sel4
					case connect:
						break;
					case XMLUpload:
						break;
					case filedownload:
						break;
					case AP_Outbound:
						break;
					case ActionClick:
						break;
					case ActionClickandEsc:
						break;
					case ActionDoubleClick:
						break;
					case ActionMouseOver:
						break;
					case AjaxPath:
						break;
					case AjaxWebButton:
						break;
					case AjaxWebList:
						break;
					case Alert:
						break;
					case Apex_Archive:
						break;
					case April:
						break;
					case ArchiveFiles:
						break;
					case AttributeisPresent:
						break;
					case August:
						break;
					case Back:
						break;
					case Back_Date:
						break;
					case BatchStatus:
						break;
					case Billing:
						break;
					case Browser:
						break;
					case BrowserSwitchUsingURL:
						break;
					case C:
						break;
					case CA:
						break;
					case CRMLoaderWait:
						break;
					case CV:
						break;
					case Calendar:
						break;
					case CalendarEBP:
						break;
					case CalendarIPF:
						break;
					case CalendarNew:
						break;
					case Calendar_DM:
						break;
					case Capture:
						break;
					case CheckBatchStatus:
						break;
					case CheckBox:
						break;
					case CheckBoxOff:
						break;
					case Clear:
						break;
					case ClearUniqueNumberFileContent:
						break;
					case CloseAndLaunchNewBrowser:
						break;
					case CloseBrowser:
						break;
					case CloseWindow:
						break;
					case ComboList:
						break;
					case CompleteFormVariable:
						break;
					case ConvertToXml:
						break;
					case CopyFiles:
						break;
					case CopyFilesToLocal:
						break;
					case CopyFilesToServer:
						break;
					case CopyFilesToWindowsServer:
						break;
					case CopyFlatFile:
						break;
					case CopyFlatFileLinux:
						break;
					case CopyFlatFile_Text:
						break;
					case CopyFromServerLocation:
						break;
					case CopyInterfaceFile:
						break;
					case CreateDynamicData:
						break;
					case CreateFile:
						break;
					case CreateJSON_Claims:
						break;
					case CreateJsonFile:
						break;
					case CurrentDate:
						break;
					case DB:
						break;
					case DBData:
						break;
					case DB_UpdateXML:
						break;
					case D_Wait:
						break;
					case Database:
						break;
					case Date:
						break;
					case December:
						break;
					case DeleteFiles:
						break;
					case DemoPage:
						break;
					case Dental_ClaimStatus:
						break;
					case DeriveBusinessDate:
						break;
					case DownloadDocument:
						break;
					case DownloadIEDocument:
						break;
					case Dropdown:
						break;
					case Edit_Enrollment:
						break;
					case EnterInput:
						break;
					case EonPagination:
						break;
					case ExecuteExe:
						break;
					case ExecuteJar:
						break;
					case Extract_Value:
						break;
					case F:
						break;
					case FIND:
						break;
					case FINDEDIT:
						break;
					case February:
						break;
					case FetchDBData:
						break;
					case FileCompare:
						break;
					case FileUpload:
						break;
					case FileUpload_CSV:
						break;
					case FileUpload_DM:
						break;
					case FlatFile:
						break;
					case FlatFileResponse:
						break;
					case FormVerificationDetails:
						break;
					case Freeze:
						break;
					case HandleWSFailure:
						break;
					case Highlight:
						break;
					case HolidayBatch:
						break;
					case I:
						break;
					case IFrame:
						break;
					case IVV:
						break;
					case I_$Value:
						break;
					case IgnoreString:
						break;
					case IgnoreTags:
						break;
					case Input:
						break;
					case InputBusinessDate:
						break;
					case InputDate:
						break;
					case InputUsingRefElements:
						break;
					case InputifExist:
						break;
					case JSClick:
						break;
					case JSEdit:
						break;
					case JSScript:
						break;
					case January:
						break;
					case July:
						break;
					case June:
						break;
					case LevelSelect:
						break;
					case ListBox:
						break;
					case LoaderWait:
						break;
					case LookupVerificationDetails:
						break;
					case March:
						break;
					case MaskedInputDate:
						break;
					case MaskedInputDatePopup:
						break;
					case MaximizeBrowser:
						break;
					case May:
						break;
					case Menu:
						break;
					case MouseClick:
						break;
					case MoveAndSaveDocument:
						break;
					case MoveDocument:
						break;
					case MoveFile:
						break;
					case MoveXMLDocument:
						break;
					case MoveXMLDocumentExpected:
						break;
					case MultiTable:
						break;
					case Multiselect:
						break;
					case NC:
						break;
					case NCIF:
						break;
					case NCIF_Button:
						break;
					case NCTable:
						break;
					case NewBrowser:
						break;
					case NoException:
						break;
					case November:
						break;
					case O:
						break;
					case October:
						break;
					case OpenAPI:
						break;
					case Opt_Back_Date:
						break;
					case Opt_CurrentDate:
						break;
					case Opt_WebElment:
						break;
					case Opt_scroll:
						break;
					case OutPutForm:
						break;
					case Override:
						break;
					case PDFComparisonMode:
						break;
					case PDFDocumentCompare:
						break;
					case PGN:
						break;
					case PT:
						break;
					case PageRefresh:
						break;
					case Pagination:
						break;
					case PolicyLink:
						break;
					case PopUpCount:
						break;
					case ProcessAndWrite:
						break;
					case QSP:
						break;
					case RD:
						break;
					case Radio:
						break;
					case Read:
						break;
					case Read2:
						break;
					case ReadD:
						break;
					case ReadTxtFile:
						break;
					case ReadValue:
						break;
					case Read_if:
						break;
					case Refresh:
						break;
					case RenameDocument:
						break;
					case RenameFile:
						break;
					case RenameXML:
						break;
					case ReplaceDefault:
						break;
					case ReplaceDynamicText:
						break;
					case RestServiceJSON_GET:
						break;
					case RestServiceJSON_POST:
						break;
					case Restful:
						break;
					case Robot:
						break;
					case RoleAssign:
						break;
					case RowNumbersToExecute:
						break;
					case RunExe:
						break;
					case RunJBeamBatch:
						break;
					case RunMacro:
						break;
					case SPNR_Wait:
						break;
					case SaveAsDocument:
						break;
					case SaveDocument:
						break;
					case Screenshot:
						break;
					case ScrollTo:
						break;
					case ScrollToElement:
						break;
					case ScrollToElement2:
						break;
					case Select:
						break;
					case SelectRadioButton:
						break;
					case SendFileUpload:
						break;
					case September:
						break;
					case SetCheckbox:
						break;
					case SikuliButton:
						break;
					case SikuliScreen:
						break;
					case SikuliType:
						break;
					case Slider:
						break;
					case SwitchtoNewBrowser:
						break;
					case T:
						break;
					case TAB:
						break;
					case TPA_Archive:
						break;
					case Tab:
						break;
					case TableInput:
						break;
					case TaskStatus:
						break;
					case TaskStatus_GUW:
						break;
					case TaskStatus_LTD:
						break;
					case TaskStatus_WOP:
						break;
					case Task_Load:
						break;
					case TestCaseUP:
						break;
					case TextboxToTextbox:
						break;
					case URL:
						break;
					case UTOV:
						break;
					case UnzipFolderAndRename:
						break;
					case UpdateCheckDataXML:
						break;
					case UpdateDBScript:
						break;
					case UpdateDataFromDB:
						break;
					case UpdateXML:
						break;
					case V:
						break;
					case VA:
						break;
					case VD:
						break;
					case VF:
						break;
					case VFCount:
						break;
					case VFPresence:
						break;
					case VFormCount:
						break;
					case VFormPresence:
						break;
					case VLookupValuesCount:
						break;
					case VLookupValuesPresence:
						break;
					case VV:
						break;
					case V_ClaimNo:
						break;
					case V_Disable:
						break;
					case V_EnableDisable:
						break;
					case V_Image:
						break;
					case V_List:
						break;
					case V_NotBlank:
						break;
					case V_ReadOnly:
						break;
					case V_Select:
						break;
					case V_TextImage:
						break;
					case V_attributeValue:
						break;
					case V_availabilityStatus:
						break;
					case V_backgroundStatus:
						break;
					case V_buttonOptionStatus:
						break;
					case V_checkboxStatus:
						break;
					case V_date:
						break;
					case V_disableStatus:
						break;
					case V_edit:
						break;
					case V_format:
						break;
					case V_opt:
						break;
					case V_stop:
						break;
					case V_text:
						break;
					case V_toggleCheckboxStatus:
						break;
					case V_toggleDisableStatus:
						break;
					case V_toggleOptionDisableStatus:
						break;
					case V_toggleOptionStatus:
						break;
					case ValidateTxtFile:
						break;
					case VerifyPopUpElement:
						break;
					case WMSG:
						break;
					case Wait:
						break;
					case WaitFor:
						break;
					case WaitForEC:
						break;
					case WaitForElementToBeFound:
						break;
					case WaitForElementToVisible:
						break;
					case WaitForJS:
						break;
					case WaitForObjectPresent:
						break;
					case WaitForPageToLoad:
						break;
					case WaitForPopUpbox:
						break;
					case WaitTillFileDownload:
						break;
					case WaitToLoad:
						break;
					case WaitTodisplayElementValue:
						break;
					case WaitUntilElementInvisible:
						break;
					case Wait_DM:
						break;
					case Wait_IfValue:
						break;
					case WebButton:
						break;
					case WebEdit:
						break;
					case WebEdit2:
						break;
					case WebEdit3:
						break;
					case WebEditEmail:
						break;
					case WebEdit_C:
						break;
					case WebEdit_Cal:
						break;
					case WebEdit_NoScroll:
						break;
					case WebElement:
						break;
					case WebImage:
						break;
					case WebLink:
						break;
					case WebList:
						break;
					case WebService:
						break;
					case WebService1:
						break;
					case WebService2:
						break;
					case WebService3:
						break;
					case WebServiceC:
						break;
					case WebServiceCSI:
						break;
					case WebServiceRP:
						break;
					case WebServiceV:
						break;
					case WebServiceV1:
						break;
					case WebServiceV2:
						break;
					case WebServiceV3:
						break;
					case WebServiceVAG:
						break;
					case WebServiceVI:
						break;
					case WebService_CheckUpdate:
						break;
					case WebService_Claims:
						break;
					case WebService_LNA:
						break;
					case WebService_VoidRef:
						break;
					case WebTable:
						break;
					case Window:
						break;
					case WindowAlertOk:
						break;
					case Write:
						break;
					case WriteAttribute:
						break;
					case WriteEntityReference:
						break;
					case XMLCompare:
						break;
					case XMLIgnoreTags_AttributeLevel:
						break;
					case XMLIgnoreTags_NodeLevel:
						break;
					case ifExist:
						break;
					case logoutLogic:
						break;
					case pickXLS:
						break;
					case screenshot:
						break;
					case spWait:
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
					// Below given missing conditions added to avoid warning message - Sel4
					case NC:
						Automation.refreshPage(controlName);
						break;
					case AP_Outbound:
						break;
					case ActionClick:
						break;
					case ActionClickandEsc:
						break;
					case ActionDoubleClick:
						break;
					case ActionMouseOver:
						break;
					case AjaxPath:
						break;
					case AjaxWebButton:
						break;
					case AjaxWebList:
						break;
					case Alert:
						break;
					case Apex_Archive:
						break;
					case April:
						break;
					case ArchiveFiles:
						break;
					case AttributeisPresent:
						break;
					case August:
						break;
					case Back:
						break;
					case Back_Date:
						break;
					case BatchStatus:
						break;
					case Billing:
						break;
					case Browser:
						break;
					case BrowserSwitchUsingURL:
						break;
					case C:
						break;
					case CA:
						break;
					case CRMLoaderWait:
						break;
					case CV:
						break;
					case Calendar:
						break;
					case CalendarEBP:
						break;
					case CalendarIPF:
						break;
					case CalendarNew:
						break;
					case Calendar_DM:
						break;
					case Capture:
						break;
					case CaptureDataFromDB:
						break;
					case CheckBatchStatus:
						break;
					case CheckBox:
						break;
					case CheckBoxOff:
						break;
					case Clear:
						break;
					case ClearUniqueNumberFileContent:
						break;
					case CloseAndLaunchNewBrowser:
						break;
					case CloseBrowser:
						break;
					case CloseWindow:
						break;
					case ComboList:
						break;
					case CompleteFormVariable:
						break;
					case ConvertToXml:
						break;
					case CopyFiles:
						break;
					case CopyFilesToLocal:
						break;
					case CopyFilesToServer:
						break;
					case CopyFilesToWindowsServer:
						break;
					case CopyFlatFile:
						break;
					case CopyFlatFileLinux:
						break;
					case CopyFlatFile_Text:
						break;
					case CopyFromServerLocation:
						break;
					case CopyInterfaceFile:
						break;
					case CreateDynamicData:
						break;
					case CreateFile:
						break;
					case CreateJSON_Claims:
						break;
					case CreateJsonFile:
						break;
					case CurrentDate:
						break;
					case DB:
						break;
					case DBData:
						break;
					case DB_UpdateXML:
						break;
					case D_Wait:
						break;
					case Database:
						break;
					case Date:
						break;
					case December:
						break;
					case DeleteFiles:
						break;
					case DemoPage:
						break;
					case Dental_ClaimStatus:
						break;
					case DeriveBusinessDate:
						break;
					case DownloadDocument:
						break;
					case DownloadIEDocument:
						break;
					case Dropdown:
						break;
					case Edit_Enrollment:
						break;
					case EnterInput:
						break;
					case EonPagination:
						break;
					case ExecuteExe:
						break;
					case ExecuteJar:
						break;
					case Extract_Value:
						break;
					case F:
						break;
					case FIND:
						break;
					case FINDEDIT:
						break;
					case February:
						break;
					case FetchDBData:
						break;
					case FileCompare:
						break;
					case FileUpload:
						break;
					case FileUpload_CSV:
						break;
					case FileUpload_DM:
						break;
					case FlatFile:
						break;
					case FlatFileResponse:
						break;
					case FormVerificationDetails:
						break;
					case Freeze:
						break;
					case HandleWSFailure:
						break;
					case Highlight:
						break;
					case HolidayBatch:
						break;
					case IFrame:
						break;
					case IVV:
						break;
					case I_$Value:
						break;
					case IgnoreString:
						break;
					case IgnoreTags:
						break;
					case Input:
						break;
					case InputBusinessDate:
						break;
					case InputDate:
						break;
					case InputUsingRefElements:
						break;
					case InputifExist:
						break;
					case JSClick:
						break;
					case JSEdit:
						break;
					case JSScript:
						break;
					case January:
						break;
					case July:
						break;
					case June:
						break;
					case LevelSelect:
						break;
					case ListBox:
						break;
					case LoaderWait:
						break;
					case LookupVerificationDetails:
						break;
					case March:
						break;
					case MaskedInputDate:
						break;
					case MaskedInputDatePopup:
						break;
					case MaximizeBrowser:
						break;
					case May:
						break;
					case Menu:
						break;
					case MouseClick:
						break;
					case MoveAndSaveDocument:
						break;
					case MoveDocument:
						break;
					case MoveFile:
						break;
					case MoveXMLDocument:
						break;
					case MoveXMLDocumentExpected:
						break;
					case MultiTable:
						break;
					case Multiselect:
						break;
					case NCIF:
						break;
					case NCIF_Button:
						break;
					case NCTable:
						break;
					case NewBrowser:
						break;
					case NoException:
						break;
					case November:
						break;
					case O:
						break;
					case October:
						break;
					case OpenAPI:
						break;
					case Opt_Back_Date:
						break;
					case Opt_CurrentDate:
						break;
					case Opt_WebElment:
						break;
					case Opt_scroll:
						break;
					case OutPutForm:
						break;
					case Override:
						break;
					case PDFComparisonMode:
						break;
					case PDFDocumentCompare:
						break;
					case PGN:
						break;
					case PT:
						break;
					case PageRefresh:
						break;
					case Pagination:
						break;
					case PolicyLink:
						break;
					case PopUpCount:
						break;
					case ProcessAndWrite:
						break;
					case QSP:
						break;
					case RD:
						break;
					case Radio:
						break;
					case Read:
						break;
					case Read2:
						break;
					case ReadD:
						break;
					case ReadTxtFile:
						break;
					case ReadValue:
						break;
					case Read_if:
						break;
					case Refresh:
						break;
					case RenameDocument:
						break;
					case RenameFile:
						break;
					case RenameXML:
						break;
					case ReplaceDefault:
						break;
					case ReplaceDynamicText:
						break;
					case RestServiceJSON_GET:
						break;
					case RestServiceJSON_POST:
						break;
					case Restful:
						break;
					case Robot:
						break;
					case RoleAssign:
						break;
					case RowNumbersToExecute:
						break;
					case RunExe:
						break;
					case RunJBeamBatch:
						break;
					case RunMacro:
						break;
					case SPNR_Wait:
						break;
					case SaveAsDocument:
						break;
					case SaveDocument:
						break;
					case Screenshot:
						break;
					case ScrollTo:
						break;
					case ScrollToElement:
						break;
					case ScrollToElement2:
						break;
					case Select:
						break;
					case SelectRadioButton:
						break;
					case SendFileUpload:
						break;
					case September:
						break;
					case SetCheckbox:
						break;
					case SikuliButton:
						break;
					case SikuliScreen:
						break;
					case SikuliType:
						break;
					case Slider:
						break;
					case SwitchtoNewBrowser:
						break;
					case T:
						break;
					case TAB:
						break;
					case TPA_Archive:
						break;
					case Tab:
						break;
					case TableInput:
						break;
					case TaskStatus:
						break;
					case TaskStatus_GUW:
						break;
					case TaskStatus_LTD:
						break;
					case TaskStatus_WOP:
						break;
					case Task_Load:
						break;
					case TestCaseUP:
						break;
					case TextboxToTextbox:
						break;
					case URL:
						break;
					case UTOV:
						break;
					case UnzipFolderAndRename:
						break;
					case UpdateCheckDataXML:
						break;
					case UpdateDB:
						break;
					case UpdateDBScript:
						break;
					case UpdateDataFromDB:
						break;
					case UpdateXML:
						break;
					case V:
						break;
					case VA:
						break;
					case VD:
						break;
					case VF:
						break;
					case VFCount:
						break;
					case VFPresence:
						break;
					case VFormCount:
						break;
					case VFormPresence:
						break;
					case VLookupValuesCount:
						break;
					case VLookupValuesPresence:
						break;
					case VV:
						break;
					case V_ClaimNo:
						break;
					case V_Disable:
						break;
					case V_EnableDisable:
						break;
					case V_Image:
						break;
					case V_List:
						break;
					case V_NotBlank:
						break;
					case V_ReadOnly:
						break;
					case V_Select:
						break;
					case V_TextImage:
						break;
					case V_attributeValue:
						break;
					case V_availabilityStatus:
						break;
					case V_backgroundStatus:
						break;
					case V_buttonOptionStatus:
						break;
					case V_checkboxStatus:
						break;
					case V_date:
						break;
					case V_disableStatus:
						break;
					case V_edit:
						break;
					case V_format:
						break;
					case V_opt:
						break;
					case V_stop:
						break;
					case V_text:
						break;
					case V_toggleCheckboxStatus:
						break;
					case V_toggleDisableStatus:
						break;
					case V_toggleOptionDisableStatus:
						break;
					case V_toggleOptionStatus:
						break;
					case ValidateTxtFile:
						break;
					case VerifyDataFromDB:
						break;
					case VerifyPopUpElement:
						break;
					case WMSG:
						break;
					case Wait:
						break;
					case WaitFor:
						break;
					case WaitForEC:
						break;
					case WaitForElementToBeFound:
						break;
					case WaitForElementToVisible:
						break;
					case WaitForJS:
						break;
					case WaitForObjectPresent:
						break;
					case WaitForPageToLoad:
						break;
					case WaitForPopUpbox:
						break;
					case WaitTillFileDownload:
						break;
					case WaitToLoad:
						break;
					case WaitTodisplayElementValue:
						break;
					case WaitUntilElementInvisible:
						break;
					case Wait_DM:
						break;
					case Wait_IfValue:
						break;
					case WebButton:
						break;
					case WebEdit:
						break;
					case WebEdit2:
						break;
					case WebEdit3:
						break;
					case WebEditEmail:
						break;
					case WebEdit_C:
						break;
					case WebEdit_Cal:
						break;
					case WebEdit_NoScroll:
						break;
					case WebElement:
						break;
					case WebImage:
						break;
					case WebLink:
						break;
					case WebList:
						break;
					case WebService:
						break;
					case WebService1:
						break;
					case WebService2:
						break;
					case WebService3:
						break;
					case WebServiceC:
						break;
					case WebServiceCSI:
						break;
					case WebServiceRP:
						break;
					case WebServiceV:
						break;
					case WebServiceV1:
						break;
					case WebServiceV2:
						break;
					case WebServiceV3:
						break;
					case WebServiceVAG:
						break;
					case WebServiceVI:
						break;
					case WebService_CheckUpdate:
						break;
					case WebService_Claims:
						break;
					case WebService_LNA:
						break;
					case WebService_VoidRef:
						break;
					case WebTable:
						break;
					case Window:
						break;
					case WindowAlertOk:
						break;
					case Write:
						break;
					case WriteAttribute:
						break;
					case WriteEntityReference:
						break;
					case XMLCompare:
						break;
					case XMLIgnoreTags_AttributeLevel:
						break;
					case XMLIgnoreTags_NodeLevel:
						break;
					case XMLUpload:
						break;
					case connect:
						break;
					case filedownload:
						break;
					case ifExist:
						break;
					case logoutLogic:
						break;
					case pickXLS:
						break;
					case screenshot:
						break;
					case spWait:
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

				case NC: // ScrollUp_Mayur
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
							System.out.println("Invalid contol value");
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
					 * System.out.println("Element does not exist");
					 * 
					 * }//Asif:Added 31-May:End
					 */
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

							System.out.println("Element Not Found on Page.");
						}
						break;

					case I_$Value: // Pradeep- Added case to handle Scroll to
									// Element as per dynamic Xpath contains
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
					case AP_Outbound:
						break;
					case ActionClick:
						break;
					case ActionClickandEsc:
						break;
					case ActionDoubleClick:
						break;
					case ActionMouseOver:
						break;
					case AjaxPath:
						break;
					case AjaxWebButton:
						break;
					case AjaxWebList:
						break;
					case Alert:
						break;
					case Apex_Archive:
						break;
					case April:
						break;
					case ArchiveFiles:
						break;
					case AttributeisPresent:
						break;
					case August:
						break;
					case Back:
						break;
					case Back_Date:
						break;
					case BatchStatus:
						break;
					case Billing:
						break;
					case Browser:
						break;
					case BrowserSwitchUsingURL:
						break;
					case C:
						break;
					case CA:
						break;
					case CRMLoaderWait:
						break;
					case CV:
						break;
					case Calendar:
						break;
					case CalendarEBP:
						break;
					case CalendarIPF:
						break;
					case CalendarNew:
						break;
					case Calendar_DM:
						break;
					case Capture:
						break;
					case CaptureDataFromDB:
						break;
					case CheckBatchStatus:
						break;
					case CheckBox:
						break;
					case CheckBoxOff:
						break;
					case Clear:
						break;
					case ClearUniqueNumberFileContent:
						break;
					case CloseAndLaunchNewBrowser:
						break;
					case CloseBrowser:
						break;
					case CloseWindow:
						break;
					case ComboList:
						break;
					case CompleteFormVariable:
						break;
					case ConvertToXml:
						break;
					case CopyFiles:
						break;
					case CopyFilesToLocal:
						break;
					case CopyFilesToServer:
						break;
					case CopyFilesToWindowsServer:
						break;
					case CopyFlatFile:
						break;
					case CopyFlatFileLinux:
						break;
					case CopyFlatFile_Text:
						break;
					case CopyFromServerLocation:
						break;
					case CopyInterfaceFile:
						break;
					case CreateDynamicData:
						break;
					case CreateFile:
						break;
					case CreateJSON_Claims:
						break;
					case CreateJsonFile:
						break;
					case CurrentDate:
						break;
					case DB:
						break;
					case DBData:
						break;
					case DB_UpdateXML:
						break;
					case D_Wait:
						break;
					case Database:
						break;
					case Date:
						break;
					case December:
						break;
					case DeleteFiles:
						break;
					case DemoPage:
						break;
					case Dental_ClaimStatus:
						break;
					case DeriveBusinessDate:
						break;
					case DownloadDocument:
						break;
					case DownloadIEDocument:
						break;
					case Dropdown:
						break;
					case Edit_Enrollment:
						break;
					case EnterInput:
						break;
					case EonPagination:
						break;
					case ExecuteExe:
						break;
					case ExecuteJar:
						break;
					case Extract_Value:
						break;
					case F:
						break;
					case FIND:
						break;
					case FINDEDIT:
						break;
					case February:
						break;
					case FetchDBData:
						break;
					case FileCompare:
						break;
					case FileUpload:
						break;
					case FileUpload_CSV:
						break;
					case FileUpload_DM:
						break;
					case FlatFile:
						break;
					case FlatFileResponse:
						break;
					case FormVerificationDetails:
						break;
					case Freeze:
						break;
					case HandleWSFailure:
						break;
					case Highlight:
						break;
					case HolidayBatch:
						break;
					case IFrame:
						break;
					case IVV:
						break;
					case IgnoreString:
						break;
					case IgnoreTags:
						break;
					case Input:
						break;
					case InputBusinessDate:
						break;
					case InputDate:
						break;
					case InputUsingRefElements:
						break;
					case InputifExist:
						break;
					case JSClick:
						break;
					case JSEdit:
						break;
					case JSScript:
						break;
					case January:
						break;
					case July:
						break;
					case June:
						break;
					case LevelSelect:
						break;
					case ListBox:
						break;
					case LoaderWait:
						break;
					case LookupVerificationDetails:
						break;
					case March:
						break;
					case MaskedInputDate:
						break;
					case MaskedInputDatePopup:
						break;
					case MaximizeBrowser:
						break;
					case May:
						break;
					case Menu:
						break;
					case MouseClick:
						break;
					case MoveAndSaveDocument:
						break;
					case MoveDocument:
						break;
					case MoveFile:
						break;
					case MoveXMLDocument:
						break;
					case MoveXMLDocumentExpected:
						break;
					case MultiTable:
						break;
					case Multiselect:
						break;
					case NCIF_Button:
						break;
					case NCTable:
						break;
					case NewBrowser:
						break;
					case NoException:
						break;
					case November:
						break;
					case O:
						break;
					case October:
						break;
					case OpenAPI:
						break;
					case Opt_Back_Date:
						break;
					case Opt_CurrentDate:
						break;
					case Opt_WebElment:
						break;
					case Opt_scroll:
						break;
					case OutPutForm:
						break;
					case Override:
						break;
					case PDFComparisonMode:
						break;
					case PDFDocumentCompare:
						break;
					case PGN:
						break;
					case PT:
						break;
					case PageRefresh:
						break;
					case Pagination:
						break;
					case PolicyLink:
						break;
					case PopUpCount:
						break;
					case ProcessAndWrite:
						break;
					case QSP:
						break;
					case RD:
						break;
					case Radio:
						break;
					case Read:
						break;
					case Read2:
						break;
					case ReadD:
						break;
					case ReadTxtFile:
						break;
					case ReadValue:
						break;
					case Read_if:
						break;
					case Refresh:
						break;
					case RenameDocument:
						break;
					case RenameFile:
						break;
					case RenameXML:
						break;
					case ReplaceDefault:
						break;
					case ReplaceDynamicText:
						break;
					case RestServiceJSON_GET:
						break;
					case RestServiceJSON_POST:
						break;
					case Restful:
						break;
					case Robot:
						break;
					case RoleAssign:
						break;
					case RowNumbersToExecute:
						break;
					case RunExe:
						break;
					case RunJBeamBatch:
						break;
					case RunMacro:
						break;
					case SPNR_Wait:
						break;
					case SaveAsDocument:
						break;
					case SaveDocument:
						break;
					case Screenshot:
						break;
					case ScrollTo:
						break;
					case ScrollToElement:
						break;
					case ScrollToElement2:
						break;
					case Select:
						break;
					case SelectRadioButton:
						break;
					case SendFileUpload:
						break;
					case September:
						break;
					case SetCheckbox:
						break;
					case SikuliButton:
						break;
					case SikuliScreen:
						break;
					case SikuliType:
						break;
					case Slider:
						break;
					case SwitchtoNewBrowser:
						break;
					case T:
						break;
					case TAB:
						break;
					case TPA_Archive:
						break;
					case Tab:
						break;
					case TableInput:
						break;
					case TaskStatus:
						break;
					case TaskStatus_GUW:
						break;
					case TaskStatus_LTD:
						break;
					case TaskStatus_WOP:
						break;
					case Task_Load:
						break;
					case TestCaseUP:
						break;
					case TextboxToTextbox:
						break;
					case URL:
						break;
					case UTOV:
						break;
					case UnzipFolderAndRename:
						break;
					case UpdateCheckDataXML:
						break;
					case UpdateDB:
						break;
					case UpdateDBScript:
						break;
					case UpdateDataFromDB:
						break;
					case UpdateXML:
						break;
					case V:
						break;
					case VA:
						break;
					case VD:
						break;
					case VF:
						break;
					case VFCount:
						break;
					case VFPresence:
						break;
					case VFormCount:
						break;
					case VFormPresence:
						break;
					case VLookupValuesCount:
						break;
					case VLookupValuesPresence:
						break;
					case VV:
						break;
					case V_ClaimNo:
						break;
					case V_Disable:
						break;
					case V_EnableDisable:
						break;
					case V_Image:
						break;
					case V_List:
						break;
					case V_NotBlank:
						break;
					case V_ReadOnly:
						break;
					case V_Select:
						break;
					case V_TextImage:
						break;
					case V_attributeValue:
						break;
					case V_availabilityStatus:
						break;
					case V_backgroundStatus:
						break;
					case V_buttonOptionStatus:
						break;
					case V_checkboxStatus:
						break;
					case V_date:
						break;
					case V_disableStatus:
						break;
					case V_edit:
						break;
					case V_format:
						break;
					case V_opt:
						break;
					case V_stop:
						break;
					case V_text:
						break;
					case V_toggleCheckboxStatus:
						break;
					case V_toggleDisableStatus:
						break;
					case V_toggleOptionDisableStatus:
						break;
					case V_toggleOptionStatus:
						break;
					case ValidateTxtFile:
						break;
					case VerifyDataFromDB:
						break;
					case VerifyPopUpElement:
						break;
					case WMSG:
						break;
					case Wait:
						break;
					case WaitFor:
						break;
					case WaitForEC:
						break;
					case WaitForElementToBeFound:
						break;
					case WaitForElementToVisible:
						break;
					case WaitForJS:
						break;
					case WaitForObjectPresent:
						break;
					case WaitForPageToLoad:
						break;
					case WaitForPopUpbox:
						break;
					case WaitTillFileDownload:
						break;
					case WaitToLoad:
						break;
					case WaitTodisplayElementValue:
						break;
					case WaitUntilElementInvisible:
						break;
					case Wait_DM:
						break;
					case Wait_IfValue:
						break;
					case WebButton:
						break;
					case WebEdit:
						break;
					case WebEdit2:
						break;
					case WebEdit3:
						break;
					case WebEditEmail:
						break;
					case WebEdit_C:
						break;
					case WebEdit_Cal:
						break;
					case WebEdit_NoScroll:
						break;
					case WebElement:
						break;
					case WebImage:
						break;
					case WebLink:
						break;
					case WebList:
						break;
					case WebService:
						break;
					case WebService1:
						break;
					case WebService2:
						break;
					case WebService3:
						break;
					case WebServiceC:
						break;
					case WebServiceCSI:
						break;
					case WebServiceRP:
						break;
					case WebServiceV:
						break;
					case WebServiceV1:
						break;
					case WebServiceV2:
						break;
					case WebServiceV3:
						break;
					case WebServiceVAG:
						break;
					case WebServiceVI:
						break;
					case WebService_CheckUpdate:
						break;
					case WebService_Claims:
						break;
					case WebService_LNA:
						break;
					case WebService_VoidRef:
						break;
					case WebTable:
						break;
					case Window:
						break;
					case WindowAlertOk:
						break;
					case Write:
						break;
					case WriteAttribute:
						break;
					case WriteEntityReference:
						break;
					case XMLCompare:
						break;
					case XMLIgnoreTags_AttributeLevel:
						break;
					case XMLIgnoreTags_NodeLevel:
						break;
					case XMLUpload:
						break;
					case connect:
						break;
					case filedownload:
						break;
					case ifExist:
						break;
					case logoutLogic:
						break;
					case pickXLS:
						break;
					case screenshot:
						break;
					case spWait:
						break;
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
				case WebService: // devishree
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
				case OpenAPI:// Mandar
					return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
							controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
							webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo, operationType,
							cycleDate, TransactionType);

				default:
					log.info("U r in Default");
					break;
				}
				// To handle processing icon in Claims
				try {
					Thread.sleep(200);
					for (int i = 1; i <= 340; i++) {
						List<WebElement> loader = Automation.driver
								.findElements(By.xpath("//body[contains(@class,'waiting')]"));
						int loaderSize = loader.size();
						// System.out.println("loaderSize is-->" + loaderSize);
						if (loaderSize >= 1)
							Thread.sleep(400);
						else
							break;
					}
				} catch (Exception e) {
				}

			} catch (WebDriverException we) {
				try {
					if (StringUtils.contains(we.getMessage().toLowerCase(), ("not clickable at point"))) {
						log.info("Click event failed. Tried JS scroll and click.");
						// ((JavascriptExecutor)
						// WebHelper.currentdriver).executeScript("arguments[0].click();",
						// webElement); //Avoid as it may not invoke rules
						((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].scrollIntoView();",
								webElement);
						webElement.click();
					} else
						throw new Exception("WebDriverException Occurred from Do Action : " + we.getMessage());
				} catch (Exception e) {
					// log.error(we.getMessage(), we);
					throw new Exception("WebDriverException Occurred from Do Action : " + we.getMessage());
				}
			} catch (IOException ioe) {
				// log.error(ioe.getMessage(), ioe);
				throw new Exception("IOException Occurred from Do Action : " + ioe.getMessage());
			} catch (Exception e) {
				// log.error(e.getMessage(), e);
				throw new Exception("Error Occurred from Do Action : " + e.getMessage());
			}
		}
		if ((action.toString().equalsIgnoreCase("V") || action.toString().equalsIgnoreCase("V_attributeValue")
				|| action.toString().equalsIgnoreCase("F") || action.toString().equalsIgnoreCase("VA")
				|| action.toString().equalsIgnoreCase("V_Disable") || action.toString().equalsIgnoreCase("V_List")
				|| action.toString().equalsIgnoreCase("V_EnableDisable")) && !ctrlValue.equalsIgnoreCase("")) {
			if (Results == true) {
				webDriver.setReport(WebHelperUtil.WriteToDetailResults(ctrlValue, currentValue, logicalName));
			}
		}
		return currentValue;
	}

	private static void webButton_Case_C(WebElement webElement) throws InterruptedException {
		WebDriverWait WaitForPageLoad1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(90));
		claimLoaderWait();
		WaitForPageLoad1.until(ExpectedConditions.visibilityOf(webElement));
		WaitForPageLoad1.until(ExpectedConditions.elementToBeClickable(webElement));
		webElement.click();
	}

	private static void webButton_Case_NC(String ctrlValue, WebElement webElement) {
		if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {

			((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].click();", webElement);

		}
	}

	private static void webButton_Case_I(String controlName, String ctrlValue, String logicalName,
			WebElement webElement, String TransactionType) throws InterruptedException, IOException {
		String uniqueNumber;
		if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes") || ctrlValue.equalsIgnoreCase("YY")
				|| !(ctrlValue.trim().equalsIgnoreCase(""))) {

			claimLoaderWait();
			Thread.sleep(500);
			scroll(controlName, webElement);
			// Thread.sleep(1000);

			if (Automation.browserType.toString().toUpperCase().contains("INTERNETEXPLORER")
					|| Automation.browserType.toString().toUpperCase().contains("MsEdge")) {
				System.out.println("IE");
				WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(90));
				WaitForPageLoad.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
				WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
				WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
				webElement.click();

				// ((JavascriptExecutor)currentdriver).executeScript("arguments[0].click();",
				// webElement);
			} else {
				if (!TransactionType.contains("HO") && (TransactionType.contains("Payment")
						&& (logicalName.equalsIgnoreCase("Preview") || logicalName.equalsIgnoreCase("BulkPaymentSubmit")
								|| logicalName.equalsIgnoreCase("ReleaseBulk")
								|| logicalName.equalsIgnoreCase("CloseBatch"))))// Mayur-
																				// Claims:
																				// Handles
																				// Duplicate
																				// //
																				// Payment
																				// Pop-up
				{
					WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(90));
					// ClaimLoader.claimLoaderWait();
					WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
					WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
					webElement.click();
					try {
						Thread.sleep(1000);
						claimLoaderWait();
						WebDriverWait WaitForPageLoadDP = new WebDriverWait(Automation.driver, Duration.ofSeconds(7));
						WebElement webElement1 = WaitForPageLoadDP.until(ExpectedConditions.elementToBeClickable(By
								.xpath("//div[@data-labelkey='mm.icd.ClaimModification.Common.DuplicatePaymentFound']//div[@data-cid='btnContinue']")));
						claimLoaderWait();
						webElement1.click();

						Thread.sleep(1000);
						log.info("Pop-up found on the web page" + "ctrlValue");
					} catch (Exception e) {
						log.error(e.getMessage(), e);
						System.out.println("No Duplicate-Payment pop up found");
					}
				} else if ((logicalName.equalsIgnoreCase("Login") || logicalName.equalsIgnoreCase("Home"))) // Mayur-
																											// Handling
																											// Chat-bot
				{
					WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(90));
					// ClaimLoader.claimLoaderWait();
					WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
					WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
					webElement.click();

					try {
						Thread.sleep(1000);
						claimLoaderWait();
						WebDriverWait WaitForPageLoadWhtsNew = new WebDriverWait(Automation.driver,
								Duration.ofSeconds(15));

						WebElement WhatsNewClose = WaitForPageLoadWhtsNew.until(ExpectedConditions.elementToBeClickable(
								By.xpath("//div[@class='modal-dialog']//button[@class='close']")));
						WhatsNewClose.click();

					} catch (Exception ex1) {
						log.info("No What's New Pop up found");
					}
					try { // Handling Chat-Bot
						Thread.sleep(2000);
						claimLoaderWait();
						WebDriverWait WaitForPageLoadChatBot = new WebDriverWait(Automation.driver,
								Duration.ofSeconds(10));
						WebElement webElementChatBot = WaitForPageLoadChatBot
								.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@id='floatBtn']//img")));
						webElementChatBot.click();

						Thread.sleep(1000);
						WebElement webElementChatBotIFrame = WaitForPageLoadChatBot
								.until(ExpectedConditions.presenceOfElementLocated(By.xpath(
										"//div[@data-cid='mlChatWrapper']//iframe[contains(@name,'ml_frame')]")));
						Automation.driver.switchTo().frame(webElementChatBotIFrame);

						WebElement webElementChatBotDock = WaitForPageLoadChatBot.until(ExpectedConditions
								.elementToBeClickable(By.xpath("//div[@class='dockCheckBox']//label")));
						// webElementChatBotDock.click();
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
								webElementChatBotDock);

						Automation.driver.switchTo().defaultContent();
					} catch (Exception ex) {
						log.info("No Chat-Bot found");
					}
				}

				else if (logicalName.equalsIgnoreCase("LoginWhatsNewPopup")) // Pradeep-
																				// Handle
																				// WhatsNew
																				// popup
																				// which
																				// comes
																				// randomly
																				// afer
																				// login
				{
					WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(90));
					// Thread.sleep(1000);
					WaitForPageLoad.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
					WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
					WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
					System.out.println("Login/Home Click");
					webElement.click();

					try {

						WaitForPageLoad.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
						WebDriverWait WaitForPageLoadChatBot = new WebDriverWait(Automation.driver,
								Duration.ofSeconds(30));
						// WebElement webElementChatBot =
						// WaitForPageLoadChatBot.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@id='floatBtn']//img")));
						WebElement webElementChatBot = WaitForPageLoadChatBot.until(ExpectedConditions
								.elementToBeClickable(By.xpath("//iframe[contains(@src,'loadwhatsnewpopup')]")));
						Thread.sleep(100);
						System.out.println("WhatsNewPopup Found");
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
						&& ctrlValue.equalsIgnoreCase("Y")) // Mayur)
				{
					uniqueNumber = ReadFromExcel(ctrlValue);
					WebElement webElement2 = WebHelper.wait.until(ExpectedConditions
							.elementToBeClickable(By.xpath(controlName + "[contains(text(),'" + uniqueNumber + "')]")));
					System.out.println(webElement2);
					WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(120));
					// ClaimLoader.claimLoaderWait();
					webElement2.click();

				} else if (ctrlValue.equalsIgnoreCase("YY")) // Mayur-double
																// Click
				{
					webElement.click();
					Thread.sleep(400);
					webElement.click();
					// builderdoubleClick.doubleClick(webElement).perform();
				} else {

					WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(90));
					// ClaimLoader.claimLoaderWait();
					WaitForPageLoad.until(ExpectedConditions.visibilityOf(webElement));
					WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
					Thread.sleep(500);
					webElement.click();

				}
			}
		}
	}

	private static String webEdit_V_attributeValue(String imageType, String ctrlValue, WebElement webElement,
			String currentValue) {
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
				if (StringUtils.equalsIgnoreCase(currentValue, null) || currentValue.equalsIgnoreCase(""))
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
		return currentValue;
	}

	private static String webEdit_V_TextImage(String controlName, String ctrlValue, String currentValue) {
		if (!(ctrlValue.equalsIgnoreCase(""))) {
			WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(60));
			WebElement Verifyelement = null;
			Verifyelement = WebHelper.currentdriver.findElement(By.xpath(controlName));
			currentValue = Verifyelement.getAttribute("src");
			Verifyelement = null;
			if (currentValue.equals(ctrlValue)) {
				System.out.println("The text matched with the ctrl value provided");
			}
		}

		else {
			System.out.println("Element not Present");
		}
		return currentValue;
	}

	private static String webEdit_V_DynamiClaimNo(String controlName, String ctrlValue) throws InterruptedException {
		String currentValue;
		claimLoaderWait();
		Reporter report = new Reporter();
		report.setReport(report);

		WebElement Verifyelement1 = null;
		// Verifyelement1 =
		// WebHelper.wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
		Verifyelement1 = WebHelper.currentdriver.findElement(By.xpath(controlName));

		// currentValue = Verifyelement1.getText();
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
		return currentValue;
	}

	private static void webEdit_V_ClaimNos(String controlName, String ctrlValue) throws IOException {
		String uniqueNumber;
		WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(60));

		uniqueNumber = ReadFromExcel(ctrlValue);
		String ExpectedValue = ctrlValue.replace("$ClaimNo$", uniqueNumber);
		WebElement WebElement = WebHelper.wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));

		System.out.println(WebElement);
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
	}

	private static String webEdit_I(String controlType, String controlId, String controlName, String ctrlValue,
			String logicalName, WebElement webElement, String rowNo, String colNo, String TransactionType)
			throws InterruptedException, IOException, Exception {
		String uniqueNumber;
		if (!ctrlValue.equalsIgnoreCase("null") || !(ctrlValue.trim().equalsIgnoreCase(""))) {
			scroll(controlName, webElement);

			WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(60));
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
						|| ctrlValue.equalsIgnoreCase("Future") || ctrlValue.equalsIgnoreCase("Past"))) // Mayur
				{
					if (ctrlValue.equalsIgnoreCase("Today")) {
						DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
						Date date = new Date();
						String todayDate = dateFormat.format(date);
						System.out.println(todayDate);
						webElement.clear();
						// Thread.sleep(500);
						if (logicalName.contains("VendorManagemtBusiness")) {
							String vendorName = "Automation" + todayDate;
							webElement.sendKeys(vendorName);
						} else {
							webElement.sendKeys(todayDate);
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
						System.out.println(todayDate);
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
						System.out.println(todayDate);
						webElement.clear();
						// Thread.sleep(500);
						webElement.sendKeys(todayDate);
					}
				} else if ((logicalName.contains("ClaimInput") || logicalName.contains("PolicyInput")
						|| logicalName.contains("$Input")) && ctrlValue.equalsIgnoreCase("Y")) // Mayur
				{
					uniqueNumber = ReadFromExcel(ctrlValue);
					webElement.clear();
					webElement.sendKeys(uniqueNumber);
				} else if ((logicalName.contains("CheckNo") || logicalName.equalsIgnoreCase("Invoice"))
						&& ctrlValue.equalsIgnoreCase("Y")) // Mayur
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
					System.out.println(saltStr);
					webElement.clear();
					webElement.sendKeys(saltStr);
					if (TransactionType.contains("HO")
							&& ((logicalName.contains("CheckNo") && TransactionType.contains("LossPayment"))
									|| logicalName.equalsIgnoreCase("Invoice"))) // Expense
																					// Payment
																					// invoice
																					// fetch
																					// data
						// webElement.sendKeys(Keys.TAB);
						WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo,
								colNo);
				} else if (logicalName.contains("Email") && ctrlValue.contains("$randomvalue")) // Mayur_
																								// For
																								// generating
																								// new
																								// email
																								// for
																								// very
																								// transaction.
				{
					String REmail = AddRandomCodeGenerator(ctrlValue);
					// Thread.sleep(500);
					webElement.clear();
					webElement.sendKeys(REmail);
					WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo,
							colNo);
				} else if (ctrlValue.contains("Today+") || ctrlValue.contains("TODAY+")) // Pradeep-
																							// Added
																							// condition
																							// to
																							// handle
																							// dynamic
																							// future
																							// dates.12/26/2019
				{
					int ctrlValueLength = ctrlValue.length();
					String PlusDays = ctrlValue.substring(6, ctrlValueLength);
					int PlusDaysValue = Integer.parseInt(PlusDays);

					LocalDate localDate = LocalDate.now();
					LocalDate Futuredate = localDate.plusDays(PlusDaysValue);
					System.out.println(Futuredate);
					DateTimeFormatter FOMATTER = DateTimeFormatter.ofPattern("MM/dd/yyyy");
					// String FutureDateFormatted =
					// Futuredate.format(DateTimeFormatter.ofPattern("MM/DD/YYYY"));
					String FutureDateFormatted = FOMATTER.format(Futuredate);
					System.out.println(FutureDateFormatted);
					webElement.clear();
					Thread.sleep(500);
					webElement.sendKeys(FutureDateFormatted);

				} else if (ctrlValue.equalsIgnoreCase("Clear")) // Pradeep-
																// Added
																// condition to
																// only clear
																// the data from
																// text field
				{
					Thread.sleep(100);
					webElement.clear();
					Thread.sleep(100);
				}

				else if (logicalName.contains("PayeeCode")) {

					String SALTCHARS = "1234567890";
					StringBuilder salt = new StringBuilder();
					Random rnd = new Random();
					while (salt.length() < 9) { // length of
												// the
												// random
												// string.
						int index = (int) (rnd.nextFloat() * SALTCHARS.length());
						salt.append(SALTCHARS.charAt(index));
					}
					String saltStr = salt.toString();
					System.out.println(saltStr);
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
		return ctrlValue;
	}

	private static void webEdit_Write(String controlType, String controlId, String controlName, String ctrlValue,
			WebElement webElement, String rowNo, String colNo) throws InterruptedException, Exception {
		if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
				|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
			claimLoaderWait();
			WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo, colNo);
		} else {
			System.out.println("Not needed");
		}
	}

	private static void webEdit_Read(String ctrlValue, WebElement webElement) throws IOException, InterruptedException {
		String uniqueNumber;
		uniqueNumber = ReadFromExcel(ctrlValue);
		claimLoaderWait();
		// log.info("!!!!!!!!!!!!!!!!");
		// log.info("uniqueNumber:"+uniqueNumber);
		webElement.clear();
		webElement.sendKeys(uniqueNumber);
	}

	/*
	 * private static void webEditNoScroll(String ctrlValue, WebElement webElement,
	 * Constants.ControlTypeEnum actionName) throws InterruptedException { switch
	 * (actionName) { case I: if (ctrlValue != null ||
	 * !ctrlValue.trim().equalsIgnoreCase("")) { log.info("ctrlValue is :" +
	 * ctrlValue); // Thread.sleep(500); webElement.clear();
	 * webElement.sendKeys(ctrlValue); Thread.sleep(100); } break;
	 */
	private static void webEditNoScroll(String ctrlValue, WebElement webElement, Constants.ControlTypeEnum actionName)
			throws InterruptedException {
		switch (actionName) {
		case I:
			if (ctrlValue != null && !ctrlValue.trim().isEmpty()) {
				log.info("ctrlValue is :" + ctrlValue);
				// Thread.sleep(500);
				webElement.clear();
				webElement.sendKeys(ctrlValue);
				Thread.sleep(100);
			}
			break;
		// Handle other cases if necessary
		default:
			// Handle default case if needed
			break;

		// Below given missing conditions added to avoid warning message - Sel4
		case AP_Outbound:
			break;
		case ActionClick:
			break;
		case ActionClickandEsc:
			break;
		case ActionDoubleClick:
			break;
		case ActionMouseOver:
			break;
		case AjaxPath:
			break;
		case AjaxWebButton:
			break;
		case AjaxWebList:
			break;
		case Alert:
			break;
		case Apex_Archive:
			break;
		case April:
			break;
		case ArchiveFiles:
			break;
		case AttributeisPresent:
			break;
		case August:
			break;
		case Back:
			break;
		case Back_Date:
			break;
		case BatchStatus:
			break;
		case Billing:
			break;
		case Browser:
			break;
		case BrowserSwitchUsingURL:
			break;
		case C:
			break;
		case CA:
			break;
		case CRMLoaderWait:
			break;
		case CV:
			break;
		case Calendar:
			break;
		case CalendarEBP:
			break;
		case CalendarIPF:
			break;
		case CalendarNew:
			break;
		case Calendar_DM:
			break;
		case Capture:
			break;
		case CaptureDataFromDB:
			break;
		case CheckBatchStatus:
			break;
		case CheckBox:
			break;
		case CheckBoxOff:
			break;
		case Clear:
			break;
		case ClearUniqueNumberFileContent:
			break;
		case CloseAndLaunchNewBrowser:
			break;
		case CloseBrowser:
			break;
		case CloseWindow:
			break;
		case ComboList:
			break;
		case CompleteFormVariable:
			break;
		case ConvertToXml:
			break;
		case CopyFiles:
			break;
		case CopyFilesToLocal:
			break;
		case CopyFilesToServer:
			break;
		case CopyFilesToWindowsServer:
			break;
		case CopyFlatFile:
			break;
		case CopyFlatFileLinux:
			break;
		case CopyFlatFile_Text:
			break;
		case CopyFromServerLocation:
			break;
		case CopyInterfaceFile:
			break;
		case CreateDynamicData:
			break;
		case CreateFile:
			break;
		case CreateJSON_Claims:
			break;
		case CreateJsonFile:
			break;
		case CurrentDate:
			break;
		case DB:
			break;
		case DBData:
			break;
		case DB_UpdateXML:
			break;
		case D_Wait:
			break;
		case Database:
			break;
		case Date:
			break;
		case December:
			break;
		case DeleteFiles:
			break;
		case DemoPage:
			break;
		case Dental_ClaimStatus:
			break;
		case DeriveBusinessDate:
			break;
		case DownloadDocument:
			break;
		case DownloadIEDocument:
			break;
		case Dropdown:
			break;
		case Edit_Enrollment:
			break;
		case EnterInput:
			break;
		case EonPagination:
			break;
		case ExecuteExe:
			break;
		case ExecuteJar:
			break;
		case Extract_Value:
			break;
		case F:
			break;
		case FIND:
			break;
		case FINDEDIT:
			break;
		case February:
			break;
		case FetchDBData:
			break;
		case FileCompare:
			break;
		case FileUpload:
			break;
		case FileUpload_CSV:
			break;
		case FileUpload_DM:
			break;
		case FlatFile:
			break;
		case FlatFileResponse:
			break;
		case FormVerificationDetails:
			break;
		case Freeze:
			break;
		case HandleWSFailure:
			break;
		case Highlight:
			break;
		case HolidayBatch:
			break;
		case IFrame:
			break;
		case IVV:
			break;
		case I_$Value:
			break;
		case IgnoreString:
			break;
		case IgnoreTags:
			break;
		case Input:
			break;
		case InputBusinessDate:
			break;
		case InputDate:
			break;
		case InputUsingRefElements:
			break;
		case InputifExist:
			break;
		case JSClick:
			break;
		case JSEdit:
			break;
		case JSScript:
			break;
		case January:
			break;
		case July:
			break;
		case June:
			break;
		case LevelSelect:
			break;
		case ListBox:
			break;
		case LoaderWait:
			break;
		case LookupVerificationDetails:
			break;
		case March:
			break;
		case MaskedInputDate:
			break;
		case MaskedInputDatePopup:
			break;
		case MaximizeBrowser:
			break;
		case May:
			break;
		case Menu:
			break;
		case MouseClick:
			break;
		case MoveAndSaveDocument:
			break;
		case MoveDocument:
			break;
		case MoveFile:
			break;
		case MoveXMLDocument:
			break;
		case MoveXMLDocumentExpected:
			break;
		case MultiTable:
			break;
		case Multiselect:
			break;
		case NC:
			break;
		case NCIF:
			break;
		case NCIF_Button:
			break;
		case NCTable:
			break;
		case NewBrowser:
			break;
		case NoException:
			break;
		case November:
			break;
		case O:
			break;
		case October:
			break;
		case OpenAPI:
			break;
		case Opt_Back_Date:
			break;
		case Opt_CurrentDate:
			break;
		case Opt_WebElment:
			break;
		case Opt_scroll:
			break;
		case OutPutForm:
			break;
		case Override:
			break;
		case PDFComparisonMode:
			break;
		case PDFDocumentCompare:
			break;
		case PGN:
			break;
		case PT:
			break;
		case PageRefresh:
			break;
		case Pagination:
			break;
		case PolicyLink:
			break;
		case PopUpCount:
			break;
		case ProcessAndWrite:
			break;
		case QSP:
			break;
		case RD:
			break;
		case Radio:
			break;
		case Read:
			break;
		case Read2:
			break;
		case ReadD:
			break;
		case ReadTxtFile:
			break;
		case ReadValue:
			break;
		case Read_if:
			break;
		case Refresh:
			break;
		case RenameDocument:
			break;
		case RenameFile:
			break;
		case RenameXML:
			break;
		case ReplaceDefault:
			break;
		case ReplaceDynamicText:
			break;
		case RestServiceJSON_GET:
			break;
		case RestServiceJSON_POST:
			break;
		case Restful:
			break;
		case Robot:
			break;
		case RoleAssign:
			break;
		case RowNumbersToExecute:
			break;
		case RunExe:
			break;
		case RunJBeamBatch:
			break;
		case RunMacro:
			break;
		case SPNR_Wait:
			break;
		case SaveAsDocument:
			break;
		case SaveDocument:
			break;
		case Screenshot:
			break;
		case ScrollTo:
			break;
		case ScrollToElement:
			break;
		case ScrollToElement2:
			break;
		case Select:
			break;
		case SelectRadioButton:
			break;
		case SendFileUpload:
			break;
		case September:
			break;
		case SetCheckbox:
			break;
		case SikuliButton:
			break;
		case SikuliScreen:
			break;
		case SikuliType:
			break;
		case Slider:
			break;
		case SwitchtoNewBrowser:
			break;
		case T:
			break;
		case TAB:
			break;
		case TPA_Archive:
			break;
		case Tab:
			break;
		case TableInput:
			break;
		case TaskStatus:
			break;
		case TaskStatus_GUW:
			break;
		case TaskStatus_LTD:
			break;
		case TaskStatus_WOP:
			break;
		case Task_Load:
			break;
		case TestCaseUP:
			break;
		case TextboxToTextbox:
			break;
		case URL:
			break;
		case UTOV:
			break;
		case UnzipFolderAndRename:
			break;
		case UpdateCheckDataXML:
			break;
		case UpdateDB:
			break;
		case UpdateDBScript:
			break;
		case UpdateDataFromDB:
			break;
		case UpdateXML:
			break;
		case V:
			break;
		case VA:
			break;
		case VD:
			break;
		case VF:
			break;
		case VFCount:
			break;
		case VFPresence:
			break;
		case VFormCount:
			break;
		case VFormPresence:
			break;
		case VLookupValuesCount:
			break;
		case VLookupValuesPresence:
			break;
		case VV:
			break;
		case V_ClaimNo:
			break;
		case V_Disable:
			break;
		case V_EnableDisable:
			break;
		case V_Image:
			break;
		case V_List:
			break;
		case V_NotBlank:
			break;
		case V_ReadOnly:
			break;
		case V_Select:
			break;
		case V_TextImage:
			break;
		case V_attributeValue:
			break;
		case V_availabilityStatus:
			break;
		case V_backgroundStatus:
			break;
		case V_buttonOptionStatus:
			break;
		case V_checkboxStatus:
			break;
		case V_date:
			break;
		case V_disableStatus:
			break;
		case V_edit:
			break;
		case V_format:
			break;
		case V_opt:
			break;
		case V_stop:
			break;
		case V_text:
			break;
		case V_toggleCheckboxStatus:
			break;
		case V_toggleDisableStatus:
			break;
		case V_toggleOptionDisableStatus:
			break;
		case V_toggleOptionStatus:
			break;
		case ValidateTxtFile:
			break;
		case VerifyDataFromDB:
			break;
		case VerifyPopUpElement:
			break;
		case WMSG:
			break;
		case Wait:
			break;
		case WaitFor:
			break;
		case WaitForEC:
			break;
		case WaitForElementToBeFound:
			break;
		case WaitForElementToVisible:
			break;
		case WaitForJS:
			break;
		case WaitForObjectPresent:
			break;
		case WaitForPageToLoad:
			break;
		case WaitForPopUpbox:
			break;
		case WaitTillFileDownload:
			break;
		case WaitToLoad:
			break;
		case WaitTodisplayElementValue:
			break;
		case WaitUntilElementInvisible:
			break;
		case Wait_DM:
			break;
		case Wait_IfValue:
			break;
		case WebButton:
			break;
		case WebEdit:
			break;
		case WebEdit2:
			break;
		case WebEdit3:
			break;
		case WebEditEmail:
			break;
		case WebEdit_C:
			break;
		case WebEdit_Cal:
			break;
		case WebEdit_NoScroll:
			break;
		case WebElement:
			break;
		case WebImage:
			break;
		case WebLink:
			break;
		case WebList:
			break;
		case WebService:
			break;
		case WebService1:
			break;
		case WebService2:
			break;
		case WebService3:
			break;
		case WebServiceC:
			break;
		case WebServiceCSI:
			break;
		case WebServiceRP:
			break;
		case WebServiceV:
			break;
		case WebServiceV1:
			break;
		case WebServiceV2:
			break;
		case WebServiceV3:
			break;
		case WebServiceVAG:
			break;
		case WebServiceVI:
			break;
		case WebService_CheckUpdate:
			break;
		case WebService_Claims:
			break;
		case WebService_LNA:
			break;
		case WebService_VoidRef:
			break;
		case WebTable:
			break;
		case Window:
			break;
		case WindowAlertOk:
			break;
		case Write:
			break;
		case WriteAttribute:
			break;
		case WriteEntityReference:
			break;
		case XMLCompare:
			break;
		case XMLIgnoreTags_AttributeLevel:
			break;
		case XMLIgnoreTags_NodeLevel:
			break;
		case XMLUpload:
			break;
		case connect:
			break;
		case filedownload:
			break;
		case ifExist:
			break;
		case logoutLogic:
			break;
		case pickXLS:
			break;
		case screenshot:
			break;
		case spWait:
			break;

		}
	}

	public static String getMonth() {
		return WebHelper.month;
	}

	public static void scroll(String XPath, WebElement we) {
		try {
			((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].scrollIntoViewIfNeeded(true);",
					we);
		} catch (Exception e) {
		}
	}

	public static String getDataFetchingSQLQuery(String sqlQuery, String ctrlValue) throws IOException {
		if (sqlQuery.contains("$value2")) {
			String[] arrTempVal = ctrlValue.split(",");
			sqlQuery = sqlQuery.replace("$value1", arrTempVal[0].toString());
			sqlQuery = sqlQuery.replace("$value2", arrTempVal[1].toString());
		} else {
			sqlQuery = sqlQuery.replace("$value", ctrlValue);
		}
		return sqlQuery;
	}

	public static String AddTodayDate(String sqlQuery) { // Generate and adds
															// Todays date to
															// SQL Query
		DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
		Date date = new Date();
		String todayDate = dateFormat.format(date);
		System.out.println(todayDate);
		sqlQuery = sqlQuery.replace("$datevalue", todayDate);
		return sqlQuery;
	}

	public static String AddRandomCodeGenerator(String sqlQuery) { // Generate
																	// and adds
																	// Random 7
																	// digit
																	// String to
																	// SQL Query
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
		WebHelper.wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));
	}

	/**
	 * Will read the Structure Sheet.
	 * 
	 * @param FilePath
	 * @param rowValues
	 * @param valuesRowIndex
	 * @param valuesRowCount
	 */
	public static void GetCellInfo(String FilePath, Row rowValues, int valuesRowIndex, int valuesRowCount) {
		String logicalName = "";
		String sqlQuery, readFromColumns, columnsToBeIncludedInQuery, writeToCol;
		boolean reportWrittenAfterException = false;
		try {
			WebHelper.TIFilePath = FilePath;
			WebHelper.currentdriver = webDriver.getReport().getDriver();
			WebHelperDS.implementWait();
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
			WebHelper.Config_endComparison = Config.endComparison;
			WebHelper.structureHeader = WebHelperUtil.getValueFromHashMap(sheetStructure);
			WebHelper.columnName = null;
			int dynamicIndexNumber;// Added for Action Loop
			String imageType, indexVal, controlName, executeFlag, action, controltype, controlID, dynamicIndex,
					newDynamicIndex, rowNo, colNo;// newly
			webDriver.getReport().setMessage("");
			webDriver.getReport().setStatus("PASS");
			VerifyDB verifyDB = VerifyDB.getNewInstance();
			// Loop over rows in the input sheet.
			// [DivyaK]: Reading over structure sheet For Loop.
			for (int rowIndex1 = 1; rowIndex1 < rowCount1 && !controller.pauseExecution; rowIndex1++) {
				controlName = WebHelperUtil.getCellData("ControlName", sheetStructure, rowIndex1,
						WebHelper.structureHeader);// structureRow.getCell(Duration.ofSeconds(3));
				executeFlag = WebHelperUtil.getCellData("ExecuteFlag", sheetStructure, rowIndex1,
						WebHelper.structureHeader);// structureRow.getCell(0);

				if (executeFlag.toString().equals("Y")) {
					imageType = WebHelperUtil.getCellData("ImageType", sheetStructure, rowIndex1,
							WebHelper.structureHeader);
					action = WebHelperUtil.getCellData("Action", sheetStructure, rowIndex1, WebHelper.structureHeader);// structureRow.getCell(1);
					logicalName = WebHelperUtil.getCellData("LogicalName", sheetStructure, rowIndex1,
							WebHelper.structureHeader);// structureRow.getCell(2);
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

					/*
					 * [DivyaK] if (logicalName.equalsIgnoreCase("PolicyNo")) {
					 * System.out.println("stop"); } if
					 * (logicalName.equalsIgnoreCase("FinalSubmitButton")) {
					 * System.out.println("stop"); } [DivyaK]
					 */

					if (action.equalsIgnoreCase("LOOP")) {
						WebHelper.loopRow = rowIndex1 + 1;
					}
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

					boolean executeBooleanFlag = Boolean.FALSE;
					if (StringUtils.isNotBlank(executeFlag) && "Y".equals(executeFlag)) {
						executeBooleanFlag = Boolean.TRUE;
					}
					calldoAction2(headerValues, logicalName, rowValues, controller.controllerTransactionType.toString(),
							valuesRowIndex, action, controlName, sheetStructure, rowIndex1,
							controller.controllerTestCaseID.toString(), controltype, controlID, indexVal, imageType,
							FilePath, rowCount1, rowNo, colNo, TransactionMapping.operationType, verifyDB,
							executeBooleanFlag);

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
			log.error(e.getMessage(), e);
			controller.pauseFun("FieldName Or LogicalName: " + logicalName + ".\n" + e.getMessage());
			reportWrittenAfterException = true;
		} finally {
			WebHelper.structureHeader.clear();
			WebHelper.valuesHeader.clear();
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

	/* [DivyaK] */
	private static class VerifyDB {
		private static final String READ_FROM_COL_NAME = "ReadFromColName";
		private static final String SQL_QUERY = "SQLQuery";
		private static final String TEST_CASE_ID = "Test Case Id";
		String testCaseId, transactionType, executeFlag, cycleDate, scenarioId, sqlQuery, readFromColName,
				columnToBeFetchedFromDb, writeToColName;

		private static VerifyDB getNewInstance() {
			return new VerifyDB();
		}

		private VerifyDB() {

		}

		public void initialise() {
			this.testCaseId = "";
			this.transactionType = "";
			this.executeFlag = "";
			this.cycleDate = "";
			this.scenarioId = "";
			this.sqlQuery = "";
			this.readFromColName = "";
		}

		public void populateAttribute(final String propertyName, final String propertyValue) {
			switch (propertyName) {
			case SQL_QUERY:
				this.sqlQuery = propertyValue;
				break;
			case READ_FROM_COL_NAME:
				this.readFromColName = propertyName;
				break;
			case TEST_CASE_ID:
				this.testCaseId = propertyValue;
				break;
			}

		}

		/* [DivyaK] */
		public String getTestCaseId() {
			return testCaseId;
		}

		public void setTestCaseId(String testCaseId) {
			this.testCaseId = testCaseId;
		}

		public String getTransactionType() {
			return transactionType;
		}

		public void setTransactionType(String transactionType) {
			this.transactionType = transactionType;
		}

		public String getExecuteFlag() {
			return executeFlag;
		}

		public void setExecuteFlag(String executeFlag) {
			this.executeFlag = executeFlag;
		}

		public String getCycleDate() {
			return cycleDate;
		}

		public void setCycleDate(String cycleDate) {
			this.cycleDate = cycleDate;
		}

		public String getScenarioId() {
			return scenarioId;
		}

		public void setScenarioId(String scenarioId) {
			this.scenarioId = scenarioId;
		}

		public String getSqlQuery() {
			return sqlQuery;
		}

		public void setSqlQuery(String sqlQuery) {
			this.sqlQuery = sqlQuery;
		}

		public String getReadFromColName() {
			return readFromColName;
		}

		public void setReadFromColName(String readFromColName) {
			this.readFromColName = readFromColName;
		}

		public String getColumnToBeFetchedFromDb() {
			return columnToBeFetchedFromDb;
		}

		public void setColumnToBeFetchedFromDb(String columnToBeFetchedFromDb) {
			this.columnToBeFetchedFromDb = columnToBeFetchedFromDb;
		}

		public String getWriteToColName() {
			return writeToColName;
		}

		public void setWriteToColName(String writeToColName) {
			this.writeToColName = writeToColName;
		}

	}

	private static class QueryData {
		private String colName, colVal;

		public String getColName() {
			return colName;
		}

		public void setColName(String colName) {
			this.colName = colName;
		}

		public String getColVal() {
			return colVal;
		}

		public void setColVal(String colVal) {
			this.colVal = colVal;
		}

		private QueryData() {

		}

		private static QueryData newInstance(String colName, String colVal) {
			QueryData obj = new QueryData();
			obj.colName = colName;
			obj.colVal = colVal;
			return obj;
		}

	}
}
