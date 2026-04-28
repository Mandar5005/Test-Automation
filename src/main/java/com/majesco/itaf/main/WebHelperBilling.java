package com.majesco.itaf.main;

import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.imageio.ImageIO;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementNotInteractableException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.Point;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.interactions.Locatable;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.sikuli.script.Screen;
import com.majesco.itaf.recovery.StartRecovery;
import com.majesco.itaf.util.BillingProduct;
import com.majesco.itaf.util.Constants;
import com.majesco.itaf.util.DB_Validation;
import com.majesco.itaf.util.DB_ValidationBilling;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.util.JDBCConnection;
import com.majesco.itaf.verification.WebVerification;
import com.majesco.itaf.vo.Reporter;


public class WebHelperBilling {

	private final static Logger log = LogManager.getLogger(WebHelperBilling.class.getName());
	public static Reporter report = new Reporter();
	public static File FlatFile;
	public static Cell eftTransaction = null;
	public static Cell WebservicecycleDate = null;
	public static String eftFlag = null;
	public static String wscycledate, NodeName, ValueToBeCompared;
	public static Boolean blank = false;
	public static boolean recovery_done = false;
	public static Sheet MainControllerSheet = null;
	public static Boolean colnotfound = false;
	public static HashMap<String, Object> vColumnheaderIndex = new HashMap<String, Object>();
	public static HashMap<String, Object> vColumnheaderValues = new HashMap<String, Object>();
	public static Date toDate = null;
	public static String testCase = null;
	public static Wait<WebDriver> waitForElementPresence;
	public static Screen sikuliScreen = null;
	public static List<String> searchValue1 = null;
	public static Boolean pageLoaded = false;
	public static String stransactionType = "";
	static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	static MainController controller = ObjectFactory.getMainController();
	public static String Account_No, columnOrderBy, MathcingColumnsList;
	public static int FailedVerification = 0;

	static void implementWait() {

		WebHelper.wait = new FluentWait<>(WebHelper.currentdriver)
				.withTimeout(Duration.ofSeconds(Integer.parseInt(Config.timeOut))).pollingEvery(Duration.ofSeconds(5))
				.ignoring(NoSuchElementException.class);
		waitForElementPresence = new FluentWait<>(Automation.driver).withTimeout(Duration.ofSeconds(30))
				.pollingEvery(Duration.ofSeconds(5)).ignoring(NoSuchElementException.class);

	}

	public static void calldoAction(Sheet headerValues, String logicalName, Row rowValues, String TransactionType,
			int valuesRowIndex, String action, String controlName, Sheet sheetStructure, int rowIndex,
			String TestCaseID, String controltype, String controlID, String indexVal, String imageType, String FilePath,
			int rowCount, String rowNo, String colNo, String operationType) throws Exception {
		String ctrlValue1 = null;
		String ctrlValue2 = null;
		String ctrlValue = null;
		String cycleDate = null;
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
		WebHelper.cycleDate_Values2 = rowValues
				.getCell(Integer.parseInt(WebHelper.valuesHeader.get("CycleDate").toString()));
		if (WebHelper.testcaseID == null) {
			testCase = "";
		} else {
			testCase = WebHelper.testcaseID.toString();
		}
		WebHelper.transactionType = rowValues
				.getCell(Integer.parseInt(WebHelper.valuesHeader.get("TransactionType").toString()));
		stransactionType = TransactionType.toString();

		if (stransactionType.toString().startsWith("FlatFile") || stransactionType.toString().contains("Outbound")) {
			log.info("transaction is FlatFile");
			WebHelper.ctrlValue1Cell = rowValues
					.getCell(Integer.parseInt(WebHelper.valuesHeader.get("ValidateTag").toString()));
			WebHelper.ctrlValue2Cell = rowValues
					.getCell(Integer.parseInt(WebHelper.valuesHeader.get("ValidationMsg").toString()));
		}

		if (stransactionType.toString().startsWith("WebService")) {
			WebHelper.ctrlValue1Cell = rowValues
					.getCell(Integer.parseInt(WebHelper.valuesHeader.get("PathToNode").toString()));
			WebHelper.ctrlValue2Cell = rowValues
					.getCell(Integer.parseInt(WebHelper.valuesHeader.get("ColumnName").toString()));
			eftTransaction = rowValues
					.getCell(Integer.parseInt(WebHelper.valuesHeader.get("WebServiceName").toString()));
			eftFlag = eftTransaction.toString();
			WebservicecycleDate = rowValues
					.getCell(Integer.parseInt(WebHelper.valuesHeader.get("CycleDate").toString()));
			if (WebHelper.ctrlValue1Cell != null && WebHelper.ctrlValue2Cell != null) {
				ctrlValue1 = WebHelper.ctrlValue1Cell.toString();
				ctrlValue2 = WebHelper.ctrlValue2Cell.toString();
			} else {
				ctrlValue1 = "";
				ctrlValue2 = "";
			}
			wscycledate = WebservicecycleDate.toString();
			SimpleDateFormat cdFormat = new SimpleDateFormat("dd-MMM-yyyy");
			DateFormat cycleDateFormat = new SimpleDateFormat("MM/dd/yyyy");
			if (wscycledate.contains("-")) {
				Date CycleDate_Values = cdFormat.parse(wscycledate);
				wscycledate = cycleDateFormat.format(CycleDate_Values);
			}
		}
		if (stransactionType.startsWith("Rest")) {
			//log.info("This is a REST service.");//Sel4 - 25/06/24
		}

		if (stransactionType.toString().equalsIgnoreCase("ChangeBusinessDate")) {
			ctrlValue = controller.businessDateValue;
		} else {
			if (colnotfound == false) {
				ctrlValue = WebHelperUtil.getCellData(logicalName, headerValues, valuesRowIndex,
						WebHelper.valuesHeader);
				if (stransactionType.toString().equalsIgnoreCase("CopyFlatFile")
						&& logicalName.toString().equalsIgnoreCase("FileName")
						&& (ctrlValue == null || ctrlValue.equals(""))) {
					ctrlValue = "NOFILE";
				}

			} else {
				ctrlValue = "";
			}
			cycleDate = WebHelperUtil.getCellData("CycleDate", headerValues, valuesRowIndex, WebHelper.valuesHeader);
		}

		// To handle $RDATETIME replacement during runtime for SIT
		if (ctrlValue.contains("${RDATETIME}")) {
			for (int i = 1; i <= 10; i++) {
				if (ctrlValue.startsWith("POL") && ctrlValue.contains("${RDATETIME}" + i)) {
					ctrlValue = WebHelperUtil.ReadFromExcel("", "PAS_PolicyNum" + i);
				} else if (ctrlValue.startsWith("AGT") && ctrlValue.contains("${RDATETIME}" + i)) {
					ctrlValue = WebHelperUtil.ReadFromExcel("", "ProducerCode" + i);
				} else if (ctrlValue.startsWith("ACC") && ctrlValue.contains("${RDATETIME}" + i)) {
					ctrlValue = WebHelperUtil.ReadFromExcel("", "AccountNum" + i);
				} else {
					String uniqueNum = WebHelperUtil.ReadFromExcel("", "UniqueNum");
					ctrlValue = ctrlValue.replace("${RDATETIME}", uniqueNum);
				}
			}
		}
		Pattern trimregex = Pattern.compile("^\\s+|\\s+$");
		Matcher match = trimregex.matcher(ctrlValue);
		StringBuffer ctrlValue_output = new StringBuffer();
		while (match.find())
			match.appendReplacement(ctrlValue_output, "");
		match.appendTail(ctrlValue_output);
		if (action.equalsIgnoreCase("Capture")) {
			@SuppressWarnings("unused")
			Reporter report = new Reporter();
			log.info("Inside Capture Case");
			controlName = WebHelperUtil.getCellData("ControlName", sheetStructure, rowIndex, WebHelper.structureHeader);
			logicalName = WebHelperUtil.getCellData("LogicalName", sheetStructure, rowIndex, WebHelper.structureHeader);
			if (ctrlValue.equalsIgnoreCase("Y")) {
				log.info("CYCLEDATE is :" + cycleDate);
				TransactionMapping.TransactionCaptureData(cycleDate, TestCaseID, controlName,
						Config.transactionInputFilePath);
			}
		}
		WebHelper.inputValue = ctrlValue;
		if (((action.equals("I") && !StringUtils.isEmpty(ctrlValue))
				|| (action.equals("V") && !StringUtils.isEmpty(ctrlValue)) || !action.equals("I"))
				&& !action.equalsIgnoreCase("FIND") && !action.equalsIgnoreCase("FINDEDIT")
				&& !action.equalsIgnoreCase("TABLEINPUT") || action.equals("PGN") && !StringUtils.isEmpty(ctrlValue)
				|| (action.equals("I") && logicalName.equalsIgnoreCase("ValidateResponsecheck"))) {

			if (logicalName.equalsIgnoreCase("WAIT")) {
				log.info(action + " " + logicalName + " For " + controlName);
			} else if ((logicalName.equalsIgnoreCase("ReadFromColumn"))
					|| (logicalName.equalsIgnoreCase("WriteToColumn"))) {

			} else if (logicalName.equalsIgnoreCase("ValidateResponsecheck") && (ctrlValue == "")) {
				log.info("ValidateResponsecheck value is empty in the Webservice.xlsx sheet,making it 'N'");
				ctrlValue = "N";
			} else {
				log.info(action + " On " + controltype + " " + logicalName);
			}

			if (logicalName.equalsIgnoreCase("CreateBatch")) {
				log.info("wait");
			}

			if (!controltype.startsWith("Sikuli")) {
				if (!action.equalsIgnoreCase("LOOP") && !controltype.equalsIgnoreCase("Wait")
						&& !controltype.equalsIgnoreCase("CloseAndLaunchNewBrowser")
						&& !action.equalsIgnoreCase("END_LOOP") && !controltype.equalsIgnoreCase("Browser")
						&& !controltype.equalsIgnoreCase("NewBrowser") && !controltype.equalsIgnoreCase("Window")
						&& !controltype.equalsIgnoreCase("Alert") && !controltype.equalsIgnoreCase("URL")
						&& !controltype.equalsIgnoreCase("WaitForJS") && !controltype.contains("Robot")
						&& !controltype.equalsIgnoreCase("Calendar") && !controltype.equalsIgnoreCase("CalendarNew")
						&& !controltype.equalsIgnoreCase("CalendarIPF") && !controltype.equalsIgnoreCase("CalendarEBP")
						&& (!action.equalsIgnoreCase("Read")
								|| ((action.equalsIgnoreCase("Read") && !StringUtils.isEmpty(controlName))))
						&& !controltype.equalsIgnoreCase("JSScript") && !controltype.equalsIgnoreCase("DB")
						&& !controlID.equalsIgnoreCase("XML") && !controltype.startsWith("Process")
						&& !controltype.startsWith("Destroy") && !controltype.startsWith("ReadSikuli")
						&& !controltype.equalsIgnoreCase("WebService") && !action.equalsIgnoreCase("VA")
						&& !action.equalsIgnoreCase("FileCompare") && !controltype.equalsIgnoreCase("Screenshot")
						&& !controltype.equalsIgnoreCase("WebService1") && !controltype.equalsIgnoreCase("WebService2")
						&& !controltype.equalsIgnoreCase("WebService3") && !controltype.equalsIgnoreCase("WebServiceV")
						&& !controltype.equalsIgnoreCase("WebServiceC") && !controltype.equalsIgnoreCase("WebServiceRP")
						&& !controltype.equalsIgnoreCase("WebServiceV1")
						&& !controltype.equalsIgnoreCase("WebServiceV2")
						&& !controltype.equalsIgnoreCase("WebServiceVAG")
						&& !controltype.equalsIgnoreCase("WebServiceV3") && !controltype.equalsIgnoreCase("OutPutForm")
						&& !controltype.equalsIgnoreCase("CopyFlatFile")
						&& !controltype.equalsIgnoreCase("CopyFlatFileLinux")
						&& !controltype.equalsIgnoreCase("CopyInterfaceFile")
						&& !controltype.equalsIgnoreCase("CopyFlatFile_Text")
						&& !controltype.equalsIgnoreCase("WebServiceCSI")
						&& !controltype.equalsIgnoreCase("WebService_CheckUpdate")
						&& !controltype.equalsIgnoreCase("WebService_VoidRef")
						&& !controltype.equalsIgnoreCase("CovertToXml") && !controltype.equalsIgnoreCase("Pagination"))

				{
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
							if ((action.equalsIgnoreCase("VV") || action.equalsIgnoreCase("I")
									|| action.equalsIgnoreCase("PGN") || action.equalsIgnoreCase("IVV")
									|| action.equalsIgnoreCase("WMSG")) && ctrlValue.equals("")) {
								WebHelper.webElementForROBOT = null;
								// }

							}

							else {
								webElement = getElementByType(controlID, controlName, controltype, imageType,
										ctrlValue);
								WebHelper.webElementForROBOT = webElement;
							}
						} catch (NoSuchElementException nse) {
							log.error("Failed to find Elements using FindBy for Control ID " + controlID
									+ " <-|-> controlName :" + controlName + "<-|-> LocalizeMessage "
									+ nse.getLocalizedMessage() + " <-|-> Message " + nse.getMessage() + " <-|-> Cause "
									+ nse.getCause(), nse);
							StartRecovery.initiateRecovery();
							throw new NoSuchElementException("Failed to find Elements using FindBy for Control ID "
									+ controlID + " <-|-> controlName :" + controlName + " <-|-> LocalizeMessage "
									+ nse.getLocalizedMessage() + " <-|-> Message" + nse.getMessage() + " <-|-> Cause "
									+ nse.getCause());
						} catch (StaleElementReferenceException sere) {
							log.error("Element is no longer appearing on the DOM page for Control ID" + controlID
									+ " <-|-> controlName :" + controlName + "<-|-> LocalizeMessage "
									+ sere.getLocalizedMessage() + " <-|-> Message " + sere.getMessage()
									+ " <-|-> Cause " + sere.getCause(), sere);
							StartRecovery.initiateRecovery();
							throw new StaleElementReferenceException(
									"Element is no longer appearing on the DOM page for Control ID " + controlID
											+ " <-|-> controlName :" + controlName + " <-|-> LocalizeMessage "
											+ sere.getLocalizedMessage() + " <-|-> Message" + sere.getMessage()
											+ " <-|-> Cause " + sere.getCause());
						} catch (ElementNotInteractableException env) {
							log.error(
									"Element is not visible Control ID" + controlID + " <-|-> controlName :"
											+ controlName + "<-|-> LocalizeMessage " + env.getLocalizedMessage()
											+ " <-|-> Message " + env.getMessage() + " <-|-> Cause " + env.getCause(),
									env);
							StartRecovery.initiateRecovery();

							throw new ElementNotInteractableException(
									"Element is not visible Control ID " + controlID + " <-|-> controlName :"
											+ controlName + " <-|-> LocalizeMessage " + env.getLocalizedMessage()
											+ " <-|-> Message" + env.getMessage() + " <-|-> Cause " + env.getCause());
						}

					}

					// Pagination//FINDEDIT added for basebilling team
					else if (action.equalsIgnoreCase("NoException") || action.equalsIgnoreCase("FIND")
							|| action.equalsIgnoreCase("FINDEDIT") || action.equalsIgnoreCase("TABLEINPUT")
							|| action.equalsIgnoreCase("I") || action.equalsIgnoreCase("PT")
							|| action.equalsIgnoreCase("IVV") || action.equalsIgnoreCase("WMSG")
							|| action.equalsIgnoreCase("PGN")) {
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
							
						case XPath_if://Sel4 19/07/24

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
								webElement = getElementByType(controlID, controlName, controltype, imageType,
										ctrlValue);
							} catch (NoSuchElementException nse) {
								log.error("Failed to find Elements using FindBy for Control ID " + scontrolID
										+ " <-|-> controlName :" + controlName + "<-|-> LocalizeMessage "
										+ nse.getLocalizedMessage() + " <-|-> Message " + nse.getMessage()
										+ " <-|-> Cause " + nse.getCause(), nse);
								StartRecovery.initiateRecovery();

							} catch (StaleElementReferenceException sere) {
								log.error("Element is no longer appearing on the DOM page for Control ID" + scontrolID
										+ " <-|-> controlName :" + controlName + "<-|-> LocalizeMessage "
										+ sere.getLocalizedMessage() + " <-|-> Message " + sere.getMessage()
										+ " <-|-> Cause " + sere.getCause(), sere);
								StartRecovery.initiateRecovery();
							} catch (ElementNotInteractableException env) {
								log.error("Element is not visible Control ID" + scontrolID + " <-|-> controlName :"
										+ controlName + "<-|-> LocalizeMessage " + env.getLocalizedMessage()
										+ " <-|-> Message " + env.getMessage() + " <-|-> Cause " + env.getCause(), env);
								StartRecovery.initiateRecovery();

							} catch (Exception env) {
								log.error("Exception is thrown for Control ID" + scontrolID + " <-|-> controlName :"
										+ controlName + "<-|-> LocalizeMessage " + env.getLocalizedMessage()
										+ " <-|-> Message " + env.getMessage() + " <-|-> Cause " + env.getCause(), env);
								StartRecovery.initiateRecovery();

							}

						} else {
							return;
						}
					} else {
						controlList = WebHelperUtil.getElementsByType(controlID, controlName, WebHelper.control,
								imageType, ctrlValue);

						if (controlList != null && controlList.size() > 1) {
							try {
								webElement = WebHelperUtil.GetControlByIndex(indexVal, controlList, controlID,
										controlName, WebHelper.control, ctrlValue);
							} catch (NoSuchElementException nse) {
								log.error("Failed to find Elements using FindBy with index for Control ID " + controlID
										+ " <-|-> controlName :" + controlName + "<-|-> LocalizeMessage "
										+ nse.getLocalizedMessage() + " <-|-> Message " + nse.getMessage()
										+ " <-|-> Cause " + nse.getCause(), nse);
								StartRecovery.initiateRecovery();
								throw new NoSuchElementException(
										"Failed to find Elements using FindBy with index for Control ID " + controlID
												+ " <-|-> controlName :" + controlName + " <-|-> LocalizeMessage "
												+ nse.getLocalizedMessage() + " <-|-> Message" + nse.getMessage()
												+ " <-|-> Cause " + nse.getCause());
							} catch (StaleElementReferenceException sere) {
								log.error("Element is no longer appearing on the DOM page with index for Control ID"
										+ controlID + " <-|-> controlName :" + controlName + "<-|-> LocalizeMessage "
										+ sere.getLocalizedMessage() + " <-|-> Message " + sere.getMessage()
										+ " <-|-> Cause " + sere.getCause(), sere);
								StartRecovery.initiateRecovery();
								throw new StaleElementReferenceException(
										"Element is no longer appearing on the DOM page with index for Control ID "
												+ controlID + " <-|-> controlName :" + controlName
												+ " <-|-> LocalizeMessage " + sere.getLocalizedMessage()
												+ " <-|-> Message" + sere.getMessage() + " <-|-> Cause "
												+ sere.getCause());
							} catch (ElementNotInteractableException env) {
								log.error("Element is not visible with index for Control ID" + controlID
										+ " <-|-> controlName :" + controlName + "<-|-> LocalizeMessage "
										+ env.getLocalizedMessage() + " <-|-> Message " + env.getMessage()
										+ " <-|-> Cause " + env.getCause(), env);
								StartRecovery.initiateRecovery();
								throw new ElementNotInteractableException(
										"Element is not visible with index for Control ID " + controlID
												+ " <-|-> controlName :" + controlName + " <-|-> LocalizeMessage "
												+ env.getLocalizedMessage() + " <-|-> Message" + env.getMessage()
												+ " <-|-> Cause " + env.getCause());
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
					headerValues, rowIndex, rowCount, rowNo, colNo, operationType, cycleDate, TransactionType);
		}
	}

	/** Locating Web Element **/
	public static WebElement getElementByType(String controlId, String controlName, String controlType,
			String imageType, String controlValue) throws Exception {

		WebElement controlList = null;
		try {

			if (controlId.equalsIgnoreCase("Id") || controlId.equalsIgnoreCase("HTMLID")) {
				controlList = WebHelper.wait.until(ExpectedConditions.elementToBeClickable(By.id(controlName)));

			} else if (controlId.equalsIgnoreCase("XPath")) {
				controlList = WebHelper.wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

			} else if (controlId.equalsIgnoreCase("Name")) {
				controlList = WebHelper.wait.until(ExpectedConditions.presenceOfElementLocated(By.name(controlName)));

			} else if (controlId.equalsIgnoreCase("XPath_Menu")) {
				controlList = waitForElementPresence
						.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

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
			} else if (controlId.equalsIgnoreCase("Id_p") || controlId.equalsIgnoreCase("HTMLID_p")) {
				controlList = WebHelper.wait.until(ExpectedConditions.presenceOfElementLocated(By.id(controlName)));
			} else if (controlId.equalsIgnoreCase("XPath_p")) {
				controlList = WebHelper.wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
			} else if (controlId.equalsIgnoreCase("XPathValue")) {
				String tempCtrlName = controlName;
				String tempReplaceString = tempCtrlName.replace("$value", controlValue);
				controlList = WebHelper.wait
						.until(ExpectedConditions.elementToBeClickable(By.xpath(tempReplaceString)));
			}
			// Removing enum constants
			return controlList;
		}

		catch (Exception ex) {
			String errorMessage = ex.getMessage();
			//String firstLine = errorMessage.split("\\r?\\n")[0];
			String firstLine = errorMessage.split("\\R")[0];
			log.error("Error :" + firstLine);
			//log.error(ex.getMessage(), ex);
			log.info("NAME : Element is not clickable : " + controlName);
			webDriver.getReport().setMessage(ex.getLocalizedMessage());
			webDriver.getReport().setStatus("FAIL");
			StartRecovery.initiateRecovery();
			/*
			 * throw new Exception("Failed while access controlName: " + controlName +
			 * " <-|-> LocalizeMessage " + ex.getLocalizedMessage() + " <-|-> Message " +
			 * ex.getMessage() + " <-|-> Cause " + ex.getCause());
			 */
			throw new Exception("Error : " + ex.getMessage());
		}
	}

	@SuppressWarnings({ "incomplete-switch" })
	public static String doAction(String FilePath, Row rowValues, String testCase, String imageType, String controlType,
			String controlId, String controlName, String ctrlValue, String ctrlValue1, String ctrlValue2,
			String wscycledate, String logicalName, String action, WebElement webElement, Boolean Results,
			Sheet strucSheet, Sheet valSheet, int rowIndex, int rowcount, String rowNo, String colNo,
			String operationType, String cycleDate, String TransactionType)
			throws WebDriverException, IOException, Exception {

		@SuppressWarnings("unused")
		String cdate, clocation;
		String currentValue = null;
		String uniqueNumber = "";
		WebVerification.isFromVerification = false;
		Constants.ControlTypeEnum controlTypeEnum = Constants.ControlTypeEnum.valueOf(controlType);
		Constants.ControlTypeEnum actionName = Constants.ControlTypeEnum.valueOf(action);

		WebHelper.sikscreen = Config.SikuliScr;

		if (controlType.contains("Robot") && !WebHelper.isIntialized) {
			log.info("In method doaction debug1");
			WebHelper.robot = new Robot();
			WebHelper.isIntialized = true;
		}

		if (!WebHelperUtil.stringIn(action,
				new String[] { "I", "PGN", "V", "F", "VA", "VV", "CV", "IVV", "PT", "WMSG" })
				|| !ctrlValue.equalsIgnoreCase("")) {
			// log.info("In method doaction debug2");
			try {
				// log.info("In method doaction debug3");
				switch (controlTypeEnum)

				{

				case WebEdit:
					switch (actionName) {
					case Read:

						if (ctrlValue.equalsIgnoreCase("IGNORE") || !webElement.isEnabled() || ctrlValue.equals(null)) {
							break;
						} else {
							uniqueNumber = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
							webElement.clear();
							log.info("uniqueNumber:" + uniqueNumber);
							((JavascriptExecutor) WebHelper.currentdriver).executeScript(
									"arguments[0].setAttribute('value', '" + uniqueNumber + "')", webElement);
							Thread.sleep(1000);
							webElement.clear();
							Thread.sleep(1000);
							webElement.sendKeys(uniqueNumber);
						}
						break;

					case ReadValue:

						if (ctrlValue.equalsIgnoreCase("IGNORE")) {
							break;
						} else {
							uniqueNumber = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
							webElement.clear();
							Thread.sleep(1000);
							webElement.sendKeys(uniqueNumber);
						}
						break;
					case Write:
						WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo,
								colNo);
						break;
					case I:
						if (ctrlValue.equalsIgnoreCase("BLANK")) {
							webElement.click();
							Thread.sleep(1000);
							webElement.clear();
						} else if (!ctrlValue.equalsIgnoreCase("null")) {
							log.info("ctrlValue is :" + ctrlValue);
							Thread.sleep(500);
							try {
								if (Automation.browserType.toString().toUpperCase().contains("CHROME")
										|| Automation.browserType.toString().toUpperCase().contains("MSEDGE")) {
									((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
											webElement);
									Thread.sleep(500);
								} else {
									webElement.click();
								}
							} catch (Exception ex) {
								log.error("Error before sleeping for 30 seconds");
								Thread.sleep(30000);
								webElement.click();
							}
							webElement.clear();
							Thread.sleep(100);
							((JavascriptExecutor) WebHelper.currentdriver).executeScript(
									"arguments[0].setAttribute('value', '" + ctrlValue + "')", webElement);
							webElement.clear();
							Thread.sleep(100);
							webElement.sendKeys(ctrlValue);
							Thread.sleep(100);

						} else {
							webElement.clear();
						}
						break;

					case PGN:
						if (!ctrlValue.equalsIgnoreCase("null")) {
							log.info("ctrlValue is :" + ctrlValue);
							((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].click();",
									webElement);
							webElement.clear();
							Thread.sleep(1000);
							((JavascriptExecutor) WebHelper.currentdriver).executeScript(
									"arguments[0].setAttribute('value', '" + ctrlValue + "')", webElement);

							webElement.clear();
							Thread.sleep(1000);
							webElement.sendKeys(ctrlValue);
							Thread.sleep(1000);
							// webElement.click();

							((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].click();",
									webElement);// Testing

							Thread.sleep(1000);
							Actions enterAction = new Actions(Automation.driver);
							enterAction.sendKeys(Keys.ENTER).build().perform();
						} else {
							webElement.clear();
						}
						break;
					case V:
						currentValue = webElement.getText();
						break;
					}
					break;

				// ***Case added for basebilling team to support email list on a new window
				// screen - 14/08/2020
				case WebEditEmail:
					switch (actionName) {
					case I:
						if (!ctrlValue.equalsIgnoreCase("null")) {
							log.info("ctrlValue is :" + ctrlValue);
							// Thread.sleep(1000);
							try {
								if (Automation.browserType.toString().toUpperCase().contains("CHROME")
										|| Automation.browserType.toString().toUpperCase().contains("MSEDGE")) {
									((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
											webElement);
									Thread.sleep(3000);
								} else {
									webElement.click();
								}
							} catch (Exception ex) {
								log.error("Error before sleeping for 30 seconds");
								Thread.sleep(30000);
								webElement.click();
							}
							webElement.clear();
							Thread.sleep(1000);
							((JavascriptExecutor) WebHelper.currentdriver).executeScript(
									"arguments[0].setAttribute('value', '" + ctrlValue + "')", webElement);
							// webElement.sendKeys(ctrlValue);
							webElement.sendKeys(Keys.ENTER);
							Thread.sleep(2000);
						} else {
							webElement.clear();
						}
						break;
					}
					break;
				// ***Case end

				case WebButton:
					switch (actionName) {
					
					case I:

						if (logicalName.equalsIgnoreCase("CloseBatch")) {
							log.info("stop");
						}

						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {

							((JavascriptExecutor) WebHelper.currentdriver)
									.executeScript("arguments[0].scrollIntoView();", webElement);

							// Wait for the element to be clickable before performing the click
							WebDriverWait wait = new WebDriverWait(WebHelper.currentdriver, Duration.ofSeconds(10));
							wait.until(ExpectedConditions.elementToBeClickable(webElement));

							List<String> browserList = Arrays.asList(
									new String[] { "InternetExplorer", "Chrome", "MsEdge", "Firefox", "Safari" });
							if (browserList.contains(Automation.browserType.toString())) {
								((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].click();",
										webElement);

								Thread.sleep(5000);
							} else {
								webElement.click();
							}
						}
						if (logicalName.equalsIgnoreCase("OkButton")) {
							log.info("stop");
						}
						break;
					 
					/*
					 * case I: if (ctrlValue == null || ctrlValue.trim().isEmpty() || webElement ==
					 * null) { break; } if (!StringUtils.equalsIgnoreCase(ctrlValue, "No")) {
					 * 
					 * int attempts = 6; while (attempts > 0) { try {
					 * System.out.println("WebButton Click Attempt: " + attempts + " with XPATH: " +
					 * controlName);
					 * 
					 * WebDriverWait wait = new WebDriverWait(Automation.driver,
					 * Duration.ofSeconds(attempts)); webElement =
					 * wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
					 * 
					 * if (webElement.isEnabled() || webElement.isSelected()) {
					 * ((JavascriptExecutor)
					 * Automation.driver).executeScript("arguments[0].scrollIntoView(true);",
					 * webElement); webElement.click(); //break; } else {
					 * System.out.println("Element not enabled."); attempts--; }
					 * 
					 * Thread.sleep(500);
					 * 
					 * boolean isSelected = false; List<WebElement> options; int options_size=-1;
					 * try { isSelected = webElement.isSelected(); options =
					 * wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath(
					 * controlName))); options_size = options.size(); } catch (Exception e) {}
					 * 
					 * if (isSelected || options_size > 0) {
					 * System.out.println("WebButton is selected through webElement.click."); break;
					 * } else { System.out.println("WebButton is not selected.");
					 * Thread.sleep(1000); attempts--; } } catch (Exception e) {
					 * System.out.println("Inside WebButton main catch. Try alternative option now."
					 * ); //log.error(e.getClass().getCanonicalName(), e);
					 * 
					 * try { WebDriverWait wait = new WebDriverWait(Automation.driver,
					 * Duration.ofSeconds(attempts)); webElement =
					 * wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
					 * Thread.sleep(500);
					 * 
					 * boolean isSelected = false;
					 * 
					 * if (webElement.isEnabled() || webElement.isSelected()) { if (attempts > 1) {
					 * ((JavascriptExecutor)
					 * Automation.driver).executeScript("arguments[0].scrollIntoView(true);",
					 * webElement); ((JavascriptExecutor)
					 * Automation.driver).executeScript("arguments[0].click();", webElement);
					 * Thread.sleep(500);
					 * 
					 * List<WebElement> options; int options_size=-1; try { isSelected = true;
					 * options =
					 * wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath(
					 * controlName))); options_size = options.size(); } catch (Exception ee) {}
					 * 
					 * if (isSelected || options_size > 0) {
					 * System.out.println("WebButton is selected through JavascriptExecutor.");
					 * break; } else { System.out.println("WebButton is not selected."); attempts--;
					 * } } else { System.out.println("Last attempt using Sendkeys.");
					 * webElement.sendKeys(Keys.ENTER); attempts--; break; } } Thread.sleep(500); }
					 * catch (Exception e1) { System.out.
					 * println("Last catch attempt to click WebButton using MouseclickAction."); try
					 * { Actions actions = new Actions(Automation.driver); Action mouseClickAction =
					 * actions.moveToElement(webElement).clickAndHold().release().build();
					 * mouseClickAction.perform(); attempts--; } catch (Exception lastE) {
					 * System.out.println("Element does not exist."); log.error(lastE.getMessage(),
					 * lastE); //Thread.sleep(1000); attempts--; break; } } } } } else {
					 * System.out.println("Don't perform any action."); } break;
					 */
					case DC:


						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {

							((JavascriptExecutor) WebHelper.currentdriver)
									.executeScript("arguments[0].scrollIntoView();", webElement);

							// Wait for the element to be clickable before performing the click
							WebDriverWait wait = new WebDriverWait(WebHelper.currentdriver, Duration.ofSeconds(10));
							wait.until(ExpectedConditions.elementToBeClickable(webElement));

							List<String> browserList = Arrays.asList(
									new String[] { "InternetExplorer", "Chrome", "MsEdge", "Firefox", "Safari" });
							 
						    // Use Actions class to double-click
						    Actions actions = new Actions(WebHelper.currentdriver);
						    
						    if (browserList.contains(Automation.browserType.toString())) {
						        // Double-click using Actions for supported browsers
						        actions.moveToElement(webElement).doubleClick().perform();
						        actions.moveToElement(webElement).doubleClick().perform();
						        Thread.sleep(5000);  // Optional: Pause to observe behavior or wait for popup/dialog to appear
						    } else {
						        // Fallback to normal single-click for unsupported browsers
						        webElement.click();
						    }
						}
						break;
					 
					case NC:

						List<String> browserList = Arrays
								.asList(new String[] { "InternetExplorer", "Chrome", "MsEdge", "Firefox", "Safari" });
						if (browserList.contains(Automation.browserType.toString())) {
							((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].click();",
									webElement);
							
						} else {
							webElement.click();
						}

						break;

					case V:
						if (webElement.isDisplayed()) {
							if (webElement.isEnabled() == true)
								currentValue = "True";
							else
								currentValue = "False";
						}
						break;

					case FileUpload:
						String autoitFileDir = Config.inputDataFilePath
								+ TransactionMapping.directoryPathFileUpload.toString();
						webElement.click();
						clocation = Config.resultFilePath + "\\ScreenShots\\" + "Before_FileUpload" + "_"
								+ WebHelper.screenshotnum + ".png";
						WebHelper.image = new Robot()
								.createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
						ImageIO.write(WebHelper.image, "png", new File(clocation));
						// ---

						Thread.sleep(5000);
						Runtime.getRuntime()
								.exec(autoitFileDir + "\\FileUpload.exe " + autoitFileDir + "\\" + ctrlValue);
						Thread.sleep(5000);
						clocation = Config.resultFilePath + "\\ScreenShots\\" + "After_FileUpload" + "_"
								+ WebHelper.screenshotnum + ".png";
						WebHelper.image = new Robot()
								.createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
						ImageIO.write(WebHelper.image, "png", new File(clocation));
						// ---***---//
						break;

					case XMLUpload: // Added for xml upload
						String autoitFileDir1 = Config.expXMLFilePath.toString();
						webElement.click();

						clocation = Config.resultFilePath + "\\ScreenShots\\" + "Before_FileUpload" + "_"
								+ WebHelper.screenshotnum + ".png";
						WebHelper.image = new Robot()
								.createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
						ImageIO.write(WebHelper.image, "png", new File(clocation));

						Thread.sleep(5000);
						Runtime.getRuntime()
								.exec(autoitFileDir1 + "\\FileUpload.exe " + autoitFileDir1 + "\\" + ctrlValue); // Meghna--03/08/2018--to

						Thread.sleep(5000);

						clocation = Config.resultFilePath + "\\ScreenShots\\" + "After_FileUpload" + "_"
								+ WebHelper.screenshotnum + ".png";
						WebHelper.image = new Robot()
								.createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
						ImageIO.write(WebHelper.image, "png", new File(clocation));

						break;

					case FileUpload_CSV:

						String autoitFileDir2;
						if (!(ctrlValue.trim().equalsIgnoreCase(""))) {
							autoitFileDir2 = Config.inputDataFilePath + TransactionMapping.directoryPathFileUpload;
							webElement.click();
							Thread.sleep(5000);
							// Added AutoIT folder in Structure sheet folder
							Runtime.getRuntime().exec(Config.structureSheetFilePath
									+ "\\AutoIT\\FileUpload\\FileUpload.exe " + autoitFileDir2 + "\\" + ctrlValue);
							Thread.sleep(5000);
						}
						break;
						
					case NCIFB://Sel4 19/07/24
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
						
					case NCIF_Button:

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
									} else {
										((JavascriptExecutor) WebHelper.currentdriver).executeScript(
												"arguments[0].click();",
												Automation.driver.findElement(By.xpath(controlName)));
									}
								}
							} else {
								break;
							}
						} catch (Exception e) {
							// throw(e);//log.error(e.getMessage(), e);
							log.info("Element Not Found on Page in NCIF case");
						}
						break;

					case ActionClick: //// Sonali : 08/20/2020

						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
							Actions builderClick = new Actions(Automation.driver);
							Action clickAction = builderClick.moveToElement(webElement).clickAndHold().release()
									.build();
							clickAction.perform();
						}
						break;

					case NCIF:

						try {
							if (Automation.driver.findElement(By.xpath(controlName)).isDisplayed()) {
								List<WebElement> ls = WebHelper.currentdriver
										.findElements(By.className("warningModalTitle"));
								System.out.println(ls.size());
								if (ls.size() > 0) {
									Thread.sleep(50);
									Automation.driver.findElement(By.xpath(controlName)).click();
								}
								// for error
								List<WebElement> ls1 = WebHelper.currentdriver
										.findElements(By.className("errorModalTitle"));
								System.out.println(ls1.size());
								if (ls1.size() > 0) {
									Thread.sleep(50);
									Automation.driver.findElement(By.xpath(controlName)).click();
								}
							} else {
								break;
							}
						} catch (Exception e) {
							// throw(e);//log.error(e.getMessage(), e);
							log.info("Element Not Found on Page in NCIF case");
						}
						break;
					}
					break;

				case DBData:

					switch (actionName) {
					case NC:
						DB_Validation.DBValidation(ctrlValue);
						Thread.sleep(500);
						break;

					case I:
						DB_Validation.DBValidation(ctrlValue);
						Thread.sleep(500);
						break;
					}
					break;

				case WebElement:
					WebVerification.isFromVerification = true;
					switch (actionName) {
					case NC:
						webElement.click();
						Thread.sleep(500);
						break;

					case Read:
						uniqueNumber = ReadFromExcel(ctrlValue);
						webElement.clear();

						((JavascriptExecutor) WebHelper.currentdriver).executeScript(
								"arguments[0].setAttribute('value', '" + uniqueNumber + "')", webElement);
						Thread.sleep(1000);
						webElement.sendKeys(uniqueNumber);
						break;
					case Write:
						WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo,
								colNo);
						break;
					case V:
						if (WebVerification.isFromVerification == true) {
							currentValue = webElement.getText();
							if (currentValue.equalsIgnoreCase(null) || currentValue.equalsIgnoreCase(""))
								currentValue = webElement.getAttribute("value");
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
						
					case I:
					    if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
					            || !ctrlValue.trim().isEmpty()) {
					    	
					    	if(logicalName.toString().equalsIgnoreCase("MajescoLogoClick")) {
					    		log.info("Clicking on the Majesco Logo");
					    		
					    	}
					            //!(ctrlValue.trim().equalsIgnoreCase(""))) 
					        WebDriverWait WaitForPageLoad = new WebDriverWait(WebHelper.currentdriver,
					                Duration.ofSeconds(10));
					        //claimLoaderWait();
					        //scroll(controlName, webElement); // Scroll element into view
					        
					        WaitForPageLoad.until(ExpectedConditions.elementToBeClickable(webElement));
					       
					        
					        
					     // Get the current position of the element
					        Object isElementInView = ((JavascriptExecutor) WebHelper.currentdriver)
					            .executeScript("var elem = arguments[0],         " +
					                           "box = elem.getBoundingClientRect(), " +
					                           "elemTop = box.top, " +
					                           "elemBottom = box.bottom, " +
					                           "isVisible = (elemTop >= 0) && (elemBottom <= window.innerHeight); " +
					                           "return isVisible;", webElement);

					        // Check if the element is in view
					        if (!(Boolean) isElementInView) {
					            // Scroll the element into view if it's not visible
					        	log.info(" Scrolling the element:" +webElement+"- into view");
					        	
					            ((JavascriptExecutor) WebHelper.currentdriver)
					                .executeScript("arguments[0].scrollIntoView();", webElement);
					        }else {
					        	
					        	log.info(" Not Scrolling the element:" +webElement+"- as its in the view already");
					        
								/*
								 * ((JavascriptExecutor) WebHelper.currentdriver)
								 * .executeScript("arguments[0].scrollIntoView();", webElement);
								 */
					        }
							/*
							 * ((JavascriptExecutor) WebHelper.currentdriver).executeScript(
							 * "arguments[0].scrollIntoViewIfNeeded(true);", webElement);
							 */
					        
					        //***Below given change done for BSP-61997 -Billing (GAIC)- One Time payment keyboard (flex) not working(23/05/2025)
					        if (webElement.getAttribute("class").contains("hg-button ng-star-inserted")) {
					            String dynamicXPath = "//div[@class='hg-button ng-star-inserted' and span='" + ctrlValue + "']";

					            webElement = WebHelper.currentdriver.findElement(By.xpath(dynamicXPath));
					        }
					        //***
					        
					        // Attempt clicking the element up to 3 times
					        int clickAttempts = 0;
					        boolean clicked = false;
					        while (clickAttempts < 3 && !clicked) {
					            try {
					            	
					                // Create an instance of Actions class
					                Actions actions = new Actions(WebHelper.currentdriver);//Disabled - 21/10/24
					                
									/*
									 * if (webElement.isEnabled()) { log.info("The element is enabled."); } else {
									 * log.info("The element is not enabled."); }
									 */
					                if (isElementPresent(webElement)) {
					                	Thread.sleep(1000);
					                // Perform the click action using Actions class
					                actions.click(webElement).perform();////Disabled - 21/10/24
					                Thread.sleep(1000);
					                //webElement.click();
					                //((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].click();", webElement);
					                clicked = true; // Set clicked to true if click succeeds
					               log.info("Click attempt " + (clickAttempts + 1) + " performed.");
					               
					            } else {
				                    log.info("Element is stale. Attempting to refresh.");
				                    webElement = WebHelper.currentdriver.findElement(By.xpath(controlName)); // Refresh element
				                }
					            }catch (StaleElementReferenceException e) {
					                // Handle StaleElementReferenceException by refreshing the element reference
					            	 log.info("Stale element reference detected. Refreshing element reference.");
					                webElement = WebHelper.currentdriver.findElement(By.xpath(controlName));
					            } catch (Exception e) {
					                // Log other exceptions
					            	 log.info("Click attempt " + (clickAttempts + 1) + " failed due to: " + e.getMessage());
					            }
					            clickAttempts++; // Increment click attempts
					        }

					        // If clicking failed after 3 attempts, throw an exception
					        if (!clicked) {
					            throw new RuntimeException("Failed to click element after 3 attempts.");
					        }
					    } else {
					        log.info("Element is not clicked");
					    }
					    
					    break;


					case IVV:
						if (!ctrlValue.equalsIgnoreCase("")) {
							Reporter report = new Reporter();
							report.setReport(report);
							Date vvDate = new Date();
							// String svvDate =
							// Automation.dtFormat.format(vvDate);
							WebHelper.ActualValue = webElement.getText();
							WebHelper.ExpectedValue = ctrlValue;
							log.info("ActualValue is : " + WebHelper.ActualValue);
							log.info("ExpectedValue is : " + WebHelper.ExpectedValue);
							// if(ActualValue.equalsIgnoreCase(ExpectedValue))
							WebHelper.ActualValue = WebHelper.ActualValue.toUpperCase();
							WebHelper.ExpectedValue = WebHelper.ExpectedValue.toUpperCase();
							if (WebHelper.ActualValue.contains(WebHelper.ExpectedValue)) {
								report.setStatus("PASS");
								report.setStatus(report.getStatus());
								report.setMessage("Values Matched");
								report.setToDate(Config.dtFormat.format(vvDate));

								if (ITAFWebDriver.isSuiteApplication()) {
									WebHelperUtil.saveScreenShot("Screenshot");
									WebHelper.ActualValue = WebHelper.ActualValue + " | " + WebHelper.ExpectedValue;
									WebHelper.ExpectedValue = WebHelper.ActualValue;
								} else
									WebHelperUtil.saveScreenShot();

								if (webElement.isDisplayed()) {
									try {
										WebElement OKwebButton = Automation.driver
												.findElement(By.xpath(".//*[@name='Ok']"));
										((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
												OKwebButton);
										log.info("OK Button clicked, Actual value is  : " + WebHelper.ActualValue);
										webDriver.getReport()
												.setMessage("Message Displayed : " + WebHelper.ActualValue);

										// OKwebButton.sendKeys(controlType);
										Thread.sleep(1000);
									} catch (Exception e)// Mandar
									{
										log.error(e.getMessage(), e);
										log.info("Ok Button on Error Msg Not Found" + e.getLocalizedMessage());
										@SuppressWarnings("unused")
										WebElement CloseDeposit = Automation.driver.findElement(By.xpath(
												".//*[@data-module='CloseDeposit']/div/div/*[@data-name='moduleTitle']"));
										// ((JavascriptExecutor)Automation.driver).executeScript("arguments[0].exist();",
										// CloseDeposit);
										log.info("Close Deposit Label present on Screen, Actual value is  : "
												+ WebHelper.ActualValue);
										webDriver.getReport()
												.setMessage("Close Deposit Sucessfully done - Label on Screen : "
														+ WebHelper.ActualValue);
									}

								}
							} else {
								report.setStatus("FAIL");
								report.setStatus(report.getStatus());
								report.setMessage("Expected error msg not matching Actual: Expected Msg is : "
										+ WebHelper.ExpectedValue + "Actual Msg is : " + WebHelper.ActualValue);
								String message = webDriver.getReport().getMessage();
								report.setToDate(Config.dtFormat.format(vvDate));
								log.error("Expected error msg not matching Actual: Expected Msg is : "
										+ WebHelper.ExpectedValue + "Actual Msg is : " + WebHelper.ActualValue);
								StartRecovery.initiateRecovery();
								controller.pauseFun(message);

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

					case WMSG:
						if (!ctrlValue.equalsIgnoreCase("")) {
							Reporter report = new Reporter();
							report.setReport(report);
							Date vvDate = new Date();
							// String svvDate =
							// Automation.dtFormat.format(vvDate);
							WebHelper.ActualValue = webElement.getText();
							WebHelper.ExpectedValue = ctrlValue;
							log.info("ActualValue is : " + WebHelper.ActualValue);
							log.info("ExpectedValue is : " + WebHelper.ExpectedValue);
							// if(ActualValue.equalsIgnoreCase(ExpectedValue))
							WebHelper.ActualValue = WebHelper.ActualValue.toUpperCase();
							WebHelper.ExpectedValue = WebHelper.ExpectedValue.toUpperCase();

							if (WebHelper.ActualValue.contains(WebHelper.ExpectedValue)) {
								report.setStatus("PASS");
								report.setStatus(report.getStatus());
								report.setMessage("Values Matched");
								report.setToDate(Config.dtFormat.format(vvDate));

								if (ITAFWebDriver.isSuiteApplication()) {
									WebHelperUtil.saveScreenShot("Screenshot");
									WebHelper.ActualValue = WebHelper.ActualValue + " | " + WebHelper.ExpectedValue;
									WebHelper.ExpectedValue = WebHelper.ActualValue;
								} else
									WebHelperUtil.saveScreenShot();

								if (webElement.isDisplayed()) {
									try {
										WebElement YeswebButton = Automation.driver
												.findElement(By.xpath(".//*[@name='Yes' or @name='Ok']"));
										((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
												YeswebButton);
										log.info("Yes Button on warning msg clicked, Actual value is  : "
												+ WebHelper.ActualValue);
										webDriver.getReport()
												.setMessage("Message Displayed : " + WebHelper.ActualValue);
										// OKwebButton.sendKeys(controlType);
										Thread.sleep(1000);
									} catch (Exception e) {
										log.error(e.getMessage(), e);
										log.info("Yes/Ok Button on Warning Msg Not Found" + e.getLocalizedMessage());
									}

								}
							}

							else {
								report.setStatus("FAIL");
								report.setStatus(report.getStatus());
								webDriver.getReport()
										.setMessage("Expected message not displayed : " + WebHelper.ActualValue);// Mandar
								String message = webDriver.getReport().getMessage();
								report.setToDate(Config.dtFormat.format(vvDate));
								controller.pauseFun(message);
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

					case VV:
						Reporter report = new Reporter();
						report.setReport(report);
						Date vvDate = new Date();
						WebHelper.ActualValue = webElement.getText();
						if (WebHelper.ActualValue.startsWith("Select")) {
							Select webElement1 = new Select(Automation.driver.findElement(By.xpath(controlName)));
							WebHelper.ActualValue = webElement1.getFirstSelectedOption().getText();
						}
						WebHelper.ExpectedValue = ctrlValue;
						log.info("ActualValue is : " + WebHelper.ActualValue);
						log.info("ExpectedValue is : " + WebHelper.ExpectedValue);
						// if(ActualValue.equalsIgnoreCase(ExpectedValue))
						WebHelper.ActualValue = WebHelper.ActualValue.toUpperCase();
						WebHelper.ExpectedValue = WebHelper.ExpectedValue.toUpperCase();
						Boolean StringPass = true;
						if (StringPass = WebHelper.ActualValue.contains(WebHelper.ExpectedValue)) {
							if (StringPass == true) {
								report.setStatus("PASS");
								report.setStatus(report.getStatus());
								report.setMessage("Values Matched");
								report.setToDate(Config.dtFormat.format(vvDate));

								if (ITAFWebDriver.isSuiteApplication()) {
									WebHelperUtil.saveScreenShot("Screenshot");
									WebHelper.ActualValue = WebHelper.ActualValue + " | " + WebHelper.ExpectedValue;
									WebHelper.ExpectedValue = WebHelper.ActualValue;
								} else {
									WebHelperUtil.saveScreenShot();
								}
							}
						} else if (StringPass == false
								&& (ITAFWebDriver.isBillingApplication() || (ITAFWebDriver.isSuiteApplication()
										&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Billing")))) {
							FailedVerification++;
							report.setStatus("FAIL");
							report.setMessage("Expected error msg not matching Actual: Expected Msg is : "
									+ WebHelper.ExpectedValue + "Actual Msg is : " + WebHelper.ActualValue);
							String message = report.getMessage();
							report.setToDate(Config.dtFormat.format(vvDate));
							controller.pauseFun(message);
							StartRecovery.initiateRecovery();//3/5/2024
							return message;//3/5/2024
						} else {
							report.setStatus("FAIL");
							report.setMessage("Expected error msg not matching Actual: Expected Msg is : "
									+ WebHelper.ExpectedValue + " Actual Msg is : " + WebHelper.ActualValue);
							String message = report.getMessage();
							report.setToDate(Config.dtFormat.format(vvDate));
							controller.pauseFun(message);
							StartRecovery.initiateRecovery();//3/5/2024
							return message;//3/5/2024
						}

						WebHelper.columns.add("");
						WebHelper.columnsData.add(WebHelper.columns);
						int temprowcount = 0;
						int tempcolcount = 0;
						ExcelUtility.WriteToCompareDetailResults(testCase, TransactionType, WebHelper.columns,
								WebHelper.columnsData, temprowcount, tempcolcount, report, WebHelper.ExpectedValue,
								WebHelper.ActualValue, logicalName, operationType, cycleDate);

						break;

					case CV:
						Thread.sleep(5000);
						Reporter reportt = new Reporter();
						reportt.setReport(reportt);
						Date cvDate = new Date();
						// String svvDate =
						// Automation.dtFormat.format(vvDate);
						WebHelper.ActualValue = webElement.getText();
						WebHelper.ExpectedValue = ctrlValue;
						log.info("ActualValue is : " + WebHelper.ActualValue);
						log.info("ExpectedValue is : " + WebHelper.ExpectedValue);
						// if(ActualValue.equalsIgnoreCase(ExpectedValue))
						WebHelper.ActualValue = WebHelper.ActualValue.toUpperCase();
						WebHelper.ExpectedValue = WebHelper.ExpectedValue.toUpperCase();
						if (WebHelper.ActualValue.equals(WebHelper.ExpectedValue)) {
							reportt.setStatus("PASS");
							reportt.setStatus(reportt.getStatus());
							reportt.setMessage("Values Matched");
							reportt.setToDate(Config.dtFormat.format(cvDate));

							if (ITAFWebDriver.isSuiteApplication()) {
								WebHelperUtil.saveScreenShot("Screenshot");
								WebHelper.ActualValue = WebHelper.ActualValue + " | " + WebHelper.ExpectedValue;
								WebHelper.ExpectedValue = WebHelper.ActualValue;
							} else
								WebHelperUtil.saveScreenShot();
						} else {
							reportt.setStatus("FAIL");
							reportt.setStatus(reportt.getStatus());
							reportt.setMessage("Values Not Matched");
							reportt.setToDate(Config.dtFormat.format(cvDate));
							controller.pauseFun("Values Not Matched");
						}
						WebHelper.columns.add("");
						WebHelper.columnsData.add(WebHelper.columns);
						int temprowcountt = 0;
						int tempcolcountt = 0;
						ExcelUtility.WriteToCompareDetailResults(testCase, TransactionType, WebHelper.columns,
								WebHelper.columnsData, temprowcountt, tempcolcountt, reportt, WebHelper.ExpectedValue,
								WebHelper.ActualValue, logicalName, operationType, cycleDate);
						break;

					}
					break;

				case JSScript:
					((JavascriptExecutor) WebHelper.currentdriver).executeScript(controlName, ctrlValue);
					break;

				case Wait:

					if ((Config.applyStaticWait).equalsIgnoreCase("True")) {
						Thread.sleep(Integer.parseInt(controlName) * 1000);
					} else {
						log.info("Wait not applied");
					}

					break;

				case WaitFor:
					return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
							controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
							webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo, operationType,
							cycleDate, TransactionType);

				case WaitToLoad:
					return WebHelper.doAction(null, null, null, null, controlType, controlId, controlName, ctrlValue,
							null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet, rowIndex,
							rowcount, rowNo, colNo, null, null, null);

				case CheckBox:
					switch (actionName) {
					case I:
						Thread.sleep(3000);
						log.info("In checkbox");
						if (ctrlValue.equalsIgnoreCase("") || StringUtils.isEmpty(ctrlValue)) {
							break;
						} else if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")) {
							if (!webElement.isSelected()) {

								((JavascriptExecutor) WebHelper.currentdriver)
										.executeScript("arguments[0].scrollIntoView();", webElement);
								((JavascriptExecutor) WebHelper.currentdriver).executeScript("arguments[0].click();",
										webElement);
							}

						}//Manual Refund change - 05/09/2024 - Added scroll in to view

						/*
						 * else if (ctrlValue.equalsIgnoreCase("N") || ctrlValue.equalsIgnoreCase("No"))
						 * { if (webElement.isEnabled()) { if (webElement.isSelected()) {
						 * webElement.click(); if (webElement.isSelected()) { webElement.click(); } } }
						 * 
						 * }
						 */
						
						//26/09/24 - Below given change done for billing to support the N condition. 
						//They have added N for check box in various testpacks where they donot want to
						//do anything ion the screen. i.e, if keep the check box as it is.
						
						else if (ctrlValue.equalsIgnoreCase("N") || ctrlValue.equalsIgnoreCase("No")) {
							if (webElement.isEnabled()) {
								// Only perform action if the checkbox is currently selected
								if (webElement.isSelected()) {

									log.info("Webelement is in selected state, un selecting it...");
									((JavascriptExecutor) WebHelper.currentdriver)
											.executeScript("arguments[0].scrollIntoView();", webElement);
									((JavascriptExecutor) WebHelper.currentdriver)
											.executeScript("arguments[0].click();", webElement);
									// webElement.click(); // Deselect the checkbox by clicking it
								}else {

								log.info("Not clicking the checkbox...");// Do nothing if the checkbox is already
								}											// unchecked (not selected)
							}
						}

						break;

					case NC:
						if (!webElement.isSelected()) {
							webElement.click();
						}
						break;
					}
					break;

				case Radio:
				case WebLink:
				case CloseWindow:
				case WaitForJS:
				case ListBox:
				case WebList:
				case AjaxWebList:
				case Refresh:
				case Browser:
				case URL:
				case Menu:
				case Alert:
				case WebImage:
				case ActionClick:
				case ActionDoubleClick:
				case ActionClickandEsc:
				case TPA_Archive:
					return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
							controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
							webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo, operationType,
							cycleDate, TransactionType);

				case IFrame:
					switch (actionName) { // Sonali 25-09-19
					case I:
						if (ctrlValue == null || ctrlValue.trim().equals("")) {
							break;
						}
						System.out.println("In method doaction debug4");
						Thread.sleep(5000);

						if (controlName.startsWith("//iframe")) {
							WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(700));// Sel 4
							wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
							Automation.driver.switchTo().frame(Automation.driver.findElement(By.xpath(controlName)));
						} else {
							Automation.driver.switchTo().frame(controlName);
						}
						System.out.println("In method doaction debug5");
						break;

					case NC:
						System.out.println("In method doaction debug4");
						Thread.sleep(5000);
						if (controlName.startsWith("//iframe")) {
							WebDriverWait wait1 = new WebDriverWait(Automation.driver, Duration.ofSeconds(700));// Sel 4
							wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
							Automation.driver.switchTo().frame(Automation.driver.findElement(By.xpath(controlName)));
						} else {
							Automation.driver.switchTo().frame(controlName);
						}

						System.out.println("In method doaction debug5");
						break;
					}
					break;
				case ActionMouseOver:
					Actions builderMouserOver = new Actions(WebHelper.currentdriver);
					builderMouserOver.moveToElement(webElement).perform();
					break;

				case NewBrowser:
					switch (actionName) {
					case I:
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
					break;

				case CloseBrowser:
					switch (actionName) {
					case I:
						if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
								|| !(ctrlValue.trim().equalsIgnoreCase(""))) {
							Automation.driver.close();
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

				/** Code for window popups **/
				case Window:
					switch (actionName) {
					case O:
						return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
								controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
								webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo,
								operationType, cycleDate, TransactionType);
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

						String tableV = WebHelperUtil.checkTable(logicalName, rowValues, ExcelUtility.TIvaluesheetrows);

						if (!(tableV.equals(""))) {
							@SuppressWarnings("unused")
							WebElement tableFound1 = WebHelper.wait
									.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));

							WebHelper.findtablefound = WebHelper.currentdriver.findElements(By.xpath(controlName))
									.size() > 0;

							if (WebHelper.findtablefound == true) {

								WebElement tableFound = WebHelper.wait
										.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));

								BillingProduct.TableInputAction(tableFound, controlName, logicalName, rowValues,
										WebHelper.valuesHeader, ExcelUtility.TIvaluesheetrows);
								Thread.sleep(1000);
							} else {
								log.info("Table not found. TABLE INPUT Functionality failed");
								break;
							}
						}

						break;

					case FIND:

						String findV = WebHelperUtil.CheckFind(logicalName, rowValues);
						if (!(findV.equals(""))) {
							Thread.sleep(5000);
							WebHelper.findtablefound = WebHelper.currentdriver.findElements(By.xpath(controlName))
									.size() > 0;
							if (WebHelper.findtablefound == true) {
								WebElement tableFound = WebHelper.wait
										.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
								BillingProduct.findAction(tableFound, controlName, logicalName, rowValues,
										WebHelper.valuesHeader);
								Thread.sleep(1000);
							} else {
								log.info("Table not found. FIND Functionality failed");
								break;
							}
						}
						break;

					case FINDEDIT:

						String findVE = WebHelperUtil.CheckFind(logicalName, rowValues);

						if (!(findVE.equals(""))) {
							Thread.sleep(5000);
							WebHelper.findtablefound = WebHelper.currentdriver.findElements(By.xpath(controlName))
									.size() > 0;

							if (WebHelper.findtablefound == true) {
								WebElement tableFound = WebHelper.wait
										.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
								BillingProduct.findActionEditable(tableFound, controlName, logicalName, rowValues,
										WebHelper.valuesHeader);
								Thread.sleep(1000);
							} else {
								log.info("Table not found. FIND Functionality failed");
								break;
							}
						}
						break;

					case I:
						Thread.sleep(10000);
						WebHelper.findtablefound = WebHelper.currentdriver.findElements(By.xpath(controlName))
								.size() > 0;
						if (WebHelper.findtablefound == true) {
							WebElement tableFound = WebHelper.wait
									.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
							List<WebElement> table_Rows = tableFound.findElements(By.tagName("tr"));
							List<WebElement> table_Columns = table_Rows.get(1).findElements(By.tagName("td"));

							int ApplicationtableRowsize = table_Rows.size();
							int Applicationtablecolumnsize = table_Columns.size();
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
												((JavascriptExecutor) WebHelper.currentdriver)
														.executeScript("arguments[0].scrollIntoView();", newelement);// Meghana
												// --
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

												((JavascriptExecutor) WebHelper.currentdriver)
														.executeScript("arguments[0].scrollIntoView();", newelement);
												if (Automation.browserType.toString().toUpperCase().contains("CHROME")
														|| Automation.browserType.toString().toUpperCase()
																.contains("MSEDGE")) {
													Actions actions = new Actions(Automation.driver);
													actions.moveToElement(newelement).click().perform();

													// End
												} else {
													((JavascriptExecutor) WebHelper.currentdriver).executeScript(
															"arguments[0].scrollIntoView();", newelement);
													((JavascriptExecutor) WebHelper.currentdriver)
															.executeScript("arguments[0].click();", newelement);
												}
												Thread.sleep(500);
												Thread.sleep(500);
												WebHelper.objfound = false;
											}
										} else if (ColumnType.equalsIgnoreCase("WebCheckBox")) {
										}
									}
								}
							}

						}
						break;
					}
					break;

				case Screenshot:
					switch (actionName) {
					case NC:
					case Screenshot:
						return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
								controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
								webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo,
								operationType, cycleDate, TransactionType);
					}
					break;

				case Robot:
					if (controlName.equalsIgnoreCase("SetFilePath")) {
						StringSelection stringSelection = new StringSelection(ctrlValue);
						Toolkit.getDefaultToolkit().getSystemClipboard().setContents(stringSelection, null);
						WebHelper.robot.delay(1000);
						WebHelper.robot.keyPress(KeyEvent.VK_CONTROL);
						WebHelper.robot.keyPress(KeyEvent.VK_V);
						WebHelper.robot.keyRelease(KeyEvent.VK_V);
						WebHelper.robot.keyRelease(KeyEvent.VK_CONTROL);

					} else if (controlName.equalsIgnoreCase("TAB")) {
						try {
							// *** Below given change made for TAB as ROBOT.sendkeys is not consistent ***
							WebHelper.robot.delay(1000);
							Actions enterAction1 = new Actions(Automation.driver);
							enterAction1.sendKeys(Keys.TAB).build().perform();

						} catch (Exception ex) {
							System.out.println("Object was not available");
						}
					} else if (controlName.equalsIgnoreCase("SPACE")) {
						WebHelper.robot.keyPress(KeyEvent.VK_SPACE);
						WebHelper.robot.keyRelease(KeyEvent.VK_SPACE);
					} else if (controlName.equalsIgnoreCase("ENTER")) {
						WebHelper.robot.keyPress(KeyEvent.VK_ENTER);
						WebHelper.robot.keyRelease(KeyEvent.VK_ENTER);
						Thread.sleep(3000);
					}
					break;

				case DB:
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

				case MaskedInputDatePopup:
					switch (actionName) {
					case I: {
						if (!ctrlValue.equalsIgnoreCase("null")) {
							webElement.sendKeys(Keys.chord(Keys.CONTROL, "a"), ctrlValue);
							webElement.sendKeys(ctrlValue);
							webElement.sendKeys(Keys.TAB);
						} else {
							webElement.clear();
						}
					}
					}
					break;

				case MaskedInputDate:
					if (!ctrlValue.equalsIgnoreCase("null")) {
						webElement.clear();
						webElement.click();
						((JavascriptExecutor) WebHelper.currentdriver)
								.executeScript("arguments[0].setAttribute('value', '" + ctrlValue + "')", webElement);
						webElement.clear();
						Thread.sleep(1000);
						webElement.sendKeys(ctrlValue);
						webElement.sendKeys(Keys.TAB);
					} else {
						webElement.clear();
					}
					break;

				case FileUpload:

					((JavascriptExecutor) WebHelper.currentdriver)
							.executeScript("arguments[0].setAttribute('value', '" + ctrlValue + "')", webElement);
					webElement.clear();
					Thread.sleep(1000);
					webElement.sendKeys(ctrlValue);
					break;

				case ScrollTo:
					Locatable element = (Locatable) webElement;
					Point p = element.getCoordinates().onScreen();
					JavascriptExecutor js = (JavascriptExecutor) WebHelper.currentdriver;
					js.executeScript("window.scrollTo(" + p.getX() + "," + (p.getY() + 150) + ");");
					break;
				case Freeze:
				case AP_Outbound:
				case FlatFileResponse:
				case FlatFile:
				case CopyFlatFile:
				case CopyFlatFileLinux:
				case CopyInterfaceFile:
				case ReplaceDynamicText:
				case CopyFlatFile_Text:
				case Task_Load:
				case Apex_Archive:
					return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
							controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
							webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo, operationType,
							cycleDate, TransactionType);
				case WebServiceCSI:
				case WebService_CheckUpdate:
				case WebService_VoidRef:
				case WebService:
				case WebService1:
				case WebService2:
				case WebService3:
				case WebServiceV:
				case WebServiceC:
				case WebServiceRP:
				case WebServiceVI:
				case WebServiceV1:
				case WebServiceV2:
				case WebServiceVAG:
				case WebServiceV3:
				case OpenAPI:
				case Restful:
					return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
							controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
							webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo, operationType,
							cycleDate, TransactionType);

				case OutPutForm:
					return WebHelper.doAction(FilePath, rowValues, testCase, imageType, controlType, controlId,
							controlName, ctrlValue, ctrlValue1, ctrlValue2, wscycledate, logicalName, action,
							webElement, Results, strucSheet, valSheet, rowIndex, rowcount, rowNo, colNo, operationType,
							cycleDate, TransactionType);

				case MoveAndSaveDocument:
					String[] temp = ctrlValue.split(";");

					String existingSameFile = Config.inputDataFilePath + "\\" + temp[1];
					File file1 = new File(existingSameFile);

					if (file1.exists()) {
						file1.delete();
						System.out.println("Old file " + temp[0] + " deleted successfully");
					}

					// move downloaded file to final PDF subfolder
					String downloadedFileName1 = Config.runtimeFileDownloadFolder + "\\" + temp[0];
					File moveFile = new File(downloadedFileName1);

					String distinationFileName = Config.inputDataFilePath + "\\" + temp[1];

					if (moveFile.renameTo(file1)) {
						// if file copied successfully then delete the original file
						if (moveFile.exists()) {
							moveFile.delete();
						}
						System.out.println(
								"File " + downloadedFileName1 + " moved successfully to " + distinationFileName);
					} else {
						System.out.println("Failed to move the file " + downloadedFileName1);
					}

					break;

				case DownloadDocument:
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case IgnoreString:
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case MoveDocument:
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case RenameDocument:
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case PDFDocumentCompare:
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case WaitTillFileDownload:
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case MoveXMLDocument: // Added for XML comparison
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case XMLIgnoreTags_AttributeLevel: // Added for XML comparison
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case XMLIgnoreTags_NodeLevel: // Added for XML comparison
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case XMLCompare: // Added for XML comparison
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case UnzipFolderAndRename: // Added for XML comparison
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case RenameXML: // Added for XML comparison for renaming the downloaded xml
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case XMLUpload: // Added for XML comparison
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case MoveXMLDocumentExpected: // Added for XML comparison to move document to download folder
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case DownloadIEDocument: // Added for XML comparison
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

				case WaitForElementToVisible: // Added for XML comparison
					return WebHelper.doAction(null, null, null, imageType, controlType, controlId, controlName,
							ctrlValue, null, null, null, logicalName, action, webElement, Results, strucSheet, valSheet,
							rowIndex, rowcount, rowNo, colNo, null, null, null);

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
					WebHelper.currentdriver = Automation.driver;
					WebHelperBilling.implementWait();
					break;

				case FetchDBData:

					switch (actionName) {
					case I:
						if (logicalName.equalsIgnoreCase("InAccount_No"))
							Account_No = ctrlValue;
						else if (logicalName.equalsIgnoreCase("ColumnOrderBy"))
							columnOrderBy = ctrlValue;

						else if (logicalName.equalsIgnoreCase("InputQuery")) {
							String InputQuery = ctrlValue;
							Connection billingDBCon;
							try {
								billingDBCon = JDBCConnection.establishHTML5BillingDBConn();
								String newQuery = InputQuery + " '" + Account_No + "' " + columnOrderBy;
								// open connection
								ResultSet rs = null;
								Statement st = billingDBCon.createStatement();
								rs = st.executeQuery(newQuery);
								HashMap<List<String>, List<List<String>>> Databasemap = new HashMap<List<String>, List<List<String>>>();
								List<String> KeySet = new ArrayList<String>();
								List<List<String>> valueSet = new ArrayList<List<String>>();
								List<String> valueSubSet = new ArrayList<String>();

								ResultSetMetaData rsmd = rs.getMetaData();
								int ccount = rsmd.getColumnCount();
								for (int i = 1; i <= rsmd.getColumnCount(); i++) {
									String KeyColomn = rsmd.getColumnName(i);
									KeySet.add(KeyColomn);
									System.out.println(KeyColomn);
								}
								int rcnt = 0;
								while (rs.next()) {
									for (rcnt = 1; rcnt <= ccount; rcnt++) {
										String TValue = rs.getString(rcnt);
										System.out.println(TValue);
										valueSubSet.add(TValue);
									}
									valueSet.add(valueSubSet);
								}
								Databasemap.put(KeySet, valueSet);
								DB_ValidationBilling.DBValidationWrtitetoExcel(Databasemap);
								InputQuery = null;
								billingDBCon.close();
								rs.close();
							} catch (Exception e) {
								throw new Exception("Error in RunDBQueries : " + e.getMessage());

							}
							break;
						}
					}
					break;

				default:
					log.info("U r in Default");
					break;
				}

				// To handle processing icon in Billing
				try {
					// Thread.sleep(200);
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

			} catch (Exception e) {
				if (!WebHelperUtil.isLastRetry()) {
					throw e;
				}
				log.error(e.getMessage(), e);
				webDriver.getReport().setMessage(e.getMessage());
				webDriver.getReport().setMessage(e.getLocalizedMessage());
				webDriver.getReport().setStatus("FAIL");
				controller.pauseFun(e.getMessage());
				log.error("Exception in doAction <-|-> LocalizeMessage " + e.getLocalizedMessage() + " <-|-> Message "
						+ e.getMessage() + " <-|-> Cause " + e.getCause());
				throw new WebDriverException("Exception in doAction <-|-> LocalizeMessage " + e.getLocalizedMessage()
						+ " <-|-> Message" + e.getMessage() + " <-|-> Cause " + e.getCause(), e);
			}
		}
		// TM-02/02/2015: Radio button found ("F") & AJAX control ("VA")
		if ((action.equalsIgnoreCase("V") || action.equalsIgnoreCase("F") || action.equalsIgnoreCase("VA"))
				&& !ctrlValue.equalsIgnoreCase("")) {
			if (Results == true) {
				webDriver.setReport(WebHelperUtil.WriteToDetailResults(ctrlValue, currentValue, logicalName));
			}
		}

		return currentValue;

	}

	static String ReadFromExcel(String controlValue) throws IOException {
		return WebHelperUtil.ReadFromExcel(controlValue, controlValue);
	}
	
	// Method to check if an element is present
	private static boolean isElementPresent(WebElement element) {
	    try {
	        element.getTagName(); // Check if the element is still present
	        return true;
	    } catch (StaleElementReferenceException | NoSuchElementException e) {
	        return false;
	    } catch (Exception e) {
	        log.error("Error checking element presence: " + e.getMessage());
	        return false;
	    }
	}


}