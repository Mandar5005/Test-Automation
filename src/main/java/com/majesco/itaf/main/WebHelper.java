package com.majesco.itaf.main;

import io.restassured.response.Response;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.PrintStream;
import java.io.PrintWriter;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Comparator;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.Objects;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.json.XML;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.Cookie;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.interactions.Locatable;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.sikuli.script.App;
import org.springframework.http.HttpMethod;
import com.majesco.compare.pdf.PDFResultBean;
import com.majesco.itaf.rest.service.RestService;
import com.majesco.itaf.rest.utils.CommonUtils;
import com.majesco.itaf.rest.utils.ValidateJson;
import com.majesco.itaf.rest.utils.XML2JsonCompare;
import com.majesco.itaf.rest.utils.JsonUtility;
import com.majesco.itaf.util.ArchiveFiles;
import com.majesco.itaf.util.BillingProduct;
import com.majesco.itaf.util.CalendarSnippet;
import com.majesco.itaf.util.CommonExpectedConditions;
import com.majesco.itaf.util.Constants;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.util.FTPFileTransferLinux;
import com.majesco.itaf.util.FlatFileComparison;
import com.majesco.itaf.util.GenerateOutputXML;
import com.majesco.itaf.util.InboundFileTransfer;
import com.majesco.itaf.util.JDBCConnection;
import com.majesco.itaf.util.Jacob;
import com.majesco.itaf.util.OutboundTransferFiles;
import com.majesco.itaf.util.PDFComparisonUtil;
import com.majesco.itaf.util.RemoveExtraSpaces;
import com.majesco.itaf.util.WaitTool;
import com.majesco.itaf.util.XmlComparisonUtil;
import com.majesco.itaf.verification.WebVerification;
import com.majesco.itaf.webservice.WebService;

public class WebHelper {

	private final static Logger log = LogManager.getLogger(WebHelper.class.getName());

	public static String description = null;
	public static Boolean faultstring = false;
	public static Boolean success = false;
	public static Date frmDate = null;
	public static File file;
	public static File dashboardfile;
	public static Wait<WebDriver> wait;
	public static String TIFilePath = "";
	public static WebDriver currentdriver;
	public static Boolean isDynamicNumFound = false;
	public static DataFormat format = null;
	public static String Config_endComparison = null;
	public static HashMap<String, Integer> structureHeader = new HashMap<>();
	public static HashMap<String, Integer> valuesHeader = new HashMap<>();
	public static String columnName;
	public static int loopRow = 1;
	public static int loopvalueendIndex = 0;
	public static Row loopexpectedRow = null;
	public static int loginCnt = 0;
	public static String inputValue = null;
	public static Cell transactionType = null;
	public static Boolean webserviceFailed = false;
	public static String ErrDescription = null;
	public static int fieldVerFailCount = 0;
	public static PrintStream print = null;
	public static int screenshotnum = 1;
	public static String cfileName = "";
	public static String control;
	public static Cell testcaseID = null;
	public static String sikscreen = null;
	public static Boolean isIntialized = false;
	public static Robot robot;
	public static BufferedImage image;
	public static String ActualValue = null;
	public static String ExpectedValue = null;
	public static String month;
	public static String responseXml = null;
	public static String request_xml;
	public static Boolean findtablefound = false;
	public static String OutPutFormCode = null;
	public static String BatchNo = null;
	public static String AccountNo = null;
	public static String BrokerNo = null;
	public static String PolicyNo = null;
	public static String OutPutForm_XML = null;
	public static String WebServiceResponse = null;
	public static boolean objfound = false;
	public static String FailedResponseTagValue = null;
	public static Boolean nullvalue = false;
	public static Boolean failed = false;
	public static String column_Name;
	public static String pathtoNode;
	public static String wsdl_url;
	public static String request_url;
	static int i = 0;
	public static WebElement webElementForROBOT;
	public static String group_no;
	public static String job_status;
	public static String account_no;
	public static String expected_status;
	public static Cell cycleDate_Values2 = null;
	public static String xfl_filename;
	public static String file_cycledate;
	public static String local_path;
	public static String remote_path_in;
	public static String remote_path_out;
	public static String file_to_be_converted;
	public static String extension;
	public static String archive = "";
	public static String uniqueFFNo = "";
	public static String reportFilePath = "";
	public static String validateTag = null;
	public static String validationMsg = null;
	public static Cell ctrlValue1Cell = null;
	public static Cell ctrlValue2Cell = null;
	public static String file_names;
	public static String destination_folder;
	public static String interface_folder;
	public static String server_folder;
	public static String compare_file;
	public static String file_name;
	public static String restErrorResDesc = "";
	public static String oasAuthType = "", queryParam = "", pathParam = "", reqHeader = "", readFromResponse = "",
			updateReqBody = "", multiPart = "";
	public static String nodeDescription = null, pathToNode = null, validateResponseIgnoreKeys = "",
			validateResponseIgnoreKeysWebservice = "", validateResponseFlag = null;
	static List<String> columns = new ArrayList<String>();
	static List<List<String>> columnsData = new ArrayList<List<String>>();
	public static String User_ID, User_ID1, User_ID2, User_ID3, User_ID4, User_ID5, User_ID6, User_ID7, User_ID8,
			User_ID9, User_ID10;
	public static int Module_No, User_No, Allocated_Count = 0;
	public static List<String> UserList = new ArrayList<String>();
	public static String LocationNo;
	private static String restUrl, requestMethod, contentType, folderPath, requestJson, authType;
	private static boolean validateResponse = false;
	public static String restResponse = null;
	public static String updateExpResponseBody = "";
	private static boolean isDMApplication = ITAFWebDriver.isDMApplication();
	public static String LocNo, PolNo;
	public static int charCount;
	private static String xmlTagName, startIndexSearch, endIndexSearch;
	public static String ValidateResponsecheck;
	static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	static MainController controller = ObjectFactory.getMainController();
	static {
		if (Config.verificationResultPath != null) {
			file = new File(Config.verificationResultPath);
		}
		if (Config.dashboardResultPath != null) {
			dashboardfile = new File(Config.dashboardResultPath);
		}
	}

	public static String doAction(String FilePath, Row rowValues, String testCase, String imageType, String controlType,
			String controlId, String controlName, String ctrlValue, String ctrlValue1, String ctrlValue2,
			String wscycledate, String logicalName, String action, WebElement webElement, Boolean Results,
			Sheet strucSheet, Sheet valSheet, int rowIndex, int rowcount, String rowNo, String colNo,
			String operationType, String cycleDate, String TransactionType)
			throws WebDriverException, IOException, Exception {

		String cdate, clocation;
		List<WebElement> WebElementList = null;
		String currentValue = null;
		String uniqueNumber = "";
		WebVerification.isFromVerification = false;
		Constants.ControlTypeEnum controlTypeEnum = Constants.ControlTypeEnum.valueOf(controlType);
		Constants.ControlTypeEnum actionName = Constants.ControlTypeEnum.valueOf(action.toString());
		String DestinationFlatFile = null;
		String SourceFlatFile = null;

		@SuppressWarnings("unused")
		Actions actions = null;
		if (!StringUtils.equalsIgnoreCase(Config.executionScope, "only_api")) {
			actions = new Actions(Automation.driver); // sonali-SF
			currentdriver = Automation.driver;
		}
		sikscreen = Config.SikuliScr;
		if (controlType.contains("Robot") && !isIntialized) {
			log.info("In method doaction debug1");
			robot = new Robot();
			isIntialized = true;
		}
		switch (controlTypeEnum)

		{

		case JSScript:
			((JavascriptExecutor) currentdriver).executeScript(controlName, ctrlValue);
			break;

		case WaitFor:

			Thread.sleep(Integer.parseInt(controlName) * 1000);
			log.info("Static Wait applied");
			break;

		case Radio:
			switch (actionName) {
			case I:

				if (ITAFWebDriver.isClaimsApplication()) {

					if (ctrlValue.equalsIgnoreCase("Y") || !ctrlValue.equalsIgnoreCase("")) {
						uniqueNumber = WebHelperClaims.ReadFromExcel(ctrlValue);
						if (ctrlValue.equalsIgnoreCase("Y") && (logicalName.contains("ReleaseClaimInputCheckBox"))) {
							try {
								WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
										Duration.ofSeconds(7));
								WebElement webElement2 = WaitForPageLoad.until(ExpectedConditions
										.elementToBeClickable(By.xpath(controlName + "[contains(text(),'" + uniqueNumber
												+ "')]//ancestor::div[contains(@id,'OpenBatchData')]//following::td[@data-colid]//input[@type='checkbox']")));
								log.info(webElement2);
								Thread.sleep(1000);
								((JavascriptExecutor) WebHelper.currentdriver)
										.executeScript("arguments[0].scrollIntoView();", webElement2);
								Thread.sleep(1000);
								webElement2.click();
							} catch (Exception e) {
								log.error(e.getMessage(), e);
								WebElement webElement1 = WebHelper.wait.until(ExpectedConditions.elementToBeClickable(By
										.xpath("//div[@id='mainRegion']/div[@class='page-region']//div[@data-name='tabExpPayDetl_tab']//li[@class='last']/a")));
								((JavascriptExecutor) WebHelper.currentdriver)
										.executeScript("arguments[0].scrollIntoView();", webElement1);
								webElement1.click();
								WebElement webElement2 = WebHelper.wait.until(ExpectedConditions
										.elementToBeClickable(By.xpath(controlName + "[contains(text(),'" + uniqueNumber
												+ "')]//ancestor::div[contains(@id,'OpenBatchData')]//following::td[@data-colid]//input[@type='checkbox']")));
								log.info(webElement2);
								Thread.sleep(1000);
								((JavascriptExecutor) WebHelper.currentdriver)
										.executeScript("arguments[0].scrollIntoView();", webElement2);
								Thread.sleep(1000);
								webElement2.click();
							}
						} else if (ctrlValue.equalsIgnoreCase("Y")
								&& (logicalName.contains("DeleteClaimBulkPayment"))) {
							try {
								WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
										Duration.ofSeconds(7));
								WebElement webElement2 = WaitForPageLoad.until(ExpectedConditions
										.elementToBeClickable(By.xpath(controlName + "[contains(text(),'" + uniqueNumber
												+ "')]//ancestor::div[contains(@id,'OpenBatchData')]//following::td[@data-colid]//img[contains(@src,'delete')]")));
								log.info(webElement2);
								Thread.sleep(1000);
								((JavascriptExecutor) WebHelper.currentdriver)
										.executeScript("arguments[0].scrollIntoView();", webElement2);
								Thread.sleep(1000);
								webElement2.click();
							} catch (Exception e) {
								log.error(e.getMessage(), e);
								WebElement webElement1 = WebHelper.wait.until(ExpectedConditions.elementToBeClickable(By
										.xpath("//div[@id='mainRegion']/div[@class='page-region']//div[@data-name='tabExpPayDetl_tab']//li[@class='last']/a")));
								((JavascriptExecutor) WebHelper.currentdriver)
										.executeScript("arguments[0].scrollIntoView();", webElement1);
								webElement1.click();
								WebElement webElement2 = WebHelper.wait.until(ExpectedConditions
										.elementToBeClickable(By.xpath(controlName + "[contains(text(),'" + uniqueNumber
												+ "')]//ancestor::div[contains(@id,'OpenBatchData')]//following::td[@data-colid]//img[contains(@src,'delete')]")));
								log.info(webElement2);
								Thread.sleep(1000);
								((JavascriptExecutor) WebHelper.currentdriver)
										.executeScript("arguments[0].scrollIntoView();", webElement2);
								Thread.sleep(1000);
								webElement2.click();
							}
						}
					}

				} else {
					controlName = controlName.replace("]", "");
					controlName = controlName + " and @value='" + ctrlValue + "']";
					log.info("controlName is:" + controlName);
					webElement = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
					if (!webElement.isSelected()) {
						Thread.sleep(1000);
						((JavascriptExecutor) currentdriver).executeScript("arguments[0].click();", webElement);
						Thread.sleep(1000);
					}
				}

				break;

			case NC:
				if (!webElement.isSelected()) {
					webElement.click();
				}
				break;
			case V:
				if (webElement.isSelected()) {
					currentValue = webElement.getAttribute(controlName.toString());
				}
				break;
			case F:
				if (webElement != null) {
					currentValue = "Y";
				}
				break;
			default:
				break;
			}
			break;

		case WebLink:
		case CloseWindow:// added this Case to bypass page loading after
			// clicking the event
			switch (actionName) {
			case Read:
				log.info("for read transaction is: " + TransactionType);
				uniqueNumber = WebHelperBilling.ReadFromExcel(ctrlValue);
				WebElementList = WebHelperUtil.getElementsByType(controlId, controlName, controlType, imageType,
						uniqueNumber);
				webElement = WebHelperUtil.GetControlByIndex("", WebElementList, controlId, controlName, controlType,
						uniqueNumber);
				webElement.click();
				break;
			case Write:
				log.info("for write transaction is: " + TransactionType);
				WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo, colNo);
				break;
			case I:
				if (controlId.equalsIgnoreCase("LinkValue")) {
					webElement.click();
				} else {
					if (ctrlValue.equalsIgnoreCase("Y") || ctrlValue.equalsIgnoreCase("Yes")
							|| !(ctrlValue.trim().equalsIgnoreCase(""))) {

						if (Automation.browserType.toString().toUpperCase().contains("INTERNETEXPLORER"))

						{
							((JavascriptExecutor) currentdriver).executeScript("arguments[0].click();", webElement);
						}

						else if (Automation.browserType.toString().toUpperCase().contains("CHROME")
								|| Automation.browserType.toString().toUpperCase().contains("MSEDGE")) {

							Actions actions1 = new Actions(Automation.driver);
							actions1.moveToElement(webElement).click().perform();
							// end

						} else {
							webElement.click();

						}
					} else if (ctrlValue.equalsIgnoreCase("") || StringUtils.isEmpty(ctrlValue)) {
						break;
					}
				}

				break;
			case NC:
			case NoException:
				try {
					((JavascriptExecutor) currentdriver).executeScript("arguments[0].click();", webElement);
				} catch (Exception ex) {
					log.error("Error before sleeping for 30 seconds");
					Thread.sleep(30000);
					((JavascriptExecutor) currentdriver).executeScript("arguments[0].click();", webElement);
				}

				break;
			default:
				break;
			}
			break;

		case WaitForJS:
			WebHelperUtil.waitForCondition();
			break;

		case ListBox:
		case WebList:
			switch (actionName) {
			case Read:
				uniqueNumber = WebHelperBilling.ReadFromExcel(ctrlValue);
				new Select(webElement).selectByVisibleText(uniqueNumber);
				break;
			case Write:
				WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo, colNo);
				break;

			case ifExist: // added to handle values with start and end as hash

				if (ctrlValue.startsWith("#") || ctrlValue.endsWith("#")) {

					if (webElement != null) {

						Select dropdown = new Select(webElement);
						Thread.sleep(1000);
						// log.info("DropDown selected");
						log.info("DropDown selected : CtrlValue is :" + ctrlValue);
						if (logicalName.equalsIgnoreCase("PolicyStatus")) {
							dropdown.selectByValue(ctrlValue);
						} else {
							dropdown.selectByVisibleText(ctrlValue);
						}
						Thread.sleep(1000);
					}
				}

				break;

			case I:
				Thread.sleep(500);
				if (ctrlValue.startsWith("#") || ctrlValue.endsWith("#")) {
					ctrlValue = ctrlValue.replace("#", " ");
					log.info("ctrlValue is : " + ctrlValue);
				}
				ExpectedCondition<Boolean> isTextPresent = CommonExpectedConditions.textToBePresentInElement(webElement,
						ctrlValue);
				if (isTextPresent != null) {
					if (webElement != null) {
						if (ctrlValue.startsWith(" ") || ctrlValue.endsWith(" ")) {
							ctrlValue = ctrlValue.replace(" ", "");
							new Select(webElement).selectByVisibleText(ctrlValue);
						} else {
							Select dropdown = new Select(webElement);
							Thread.sleep(100);
							// log.info("DropDown selected");
							if (logicalName.equalsIgnoreCase("PolicyStatus")) // devishree
							{
								dropdown.selectByValue(ctrlValue);
							} else {
								dropdown.selectByVisibleText(ctrlValue);
								Thread.sleep(100);
								if ("GroupBilling".equalsIgnoreCase(Config.productTeam)) {
									dropdown.selectByVisibleText(ctrlValue);// To test GroupBilling changes 18/02/2020
								}
							}
							log.info("DropDown selected : CtrlValue is :" + ctrlValue);
							Thread.sleep(100);

						}
					}
				}

				break;

			case PT:
				if (ctrlValue.startsWith("#") || ctrlValue.endsWith("#")) {
					ctrlValue = ctrlValue.replace("#", " ");
					log.info("ctrlValue is : " + ctrlValue);
				}
				String partialText = ctrlValue;
				List<WebElement> list = currentdriver.findElements(By.tagName("option"));
				Iterator<WebElement> i = list.iterator();
				while (i.hasNext()) {
					WebElement wel = i.next();
					if (wel.getText().contains(partialText)) {
						wel.click();
					}
				}

				Thread.sleep(1000);

				break;
			case V:
				if (!ctrlValue.contains(",")) {
					currentValue = new Select(webElement).getFirstSelectedOption().getText();
					if (StringUtils.isEmpty(currentValue)) {
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
			default:
				break;
			}
			break;

		// New code for AJAX Dropdown with dojo
		case AjaxWebList:
			switch (actionName) {
			case I:
				webElement.click();
				break;
			case VA:
				Thread.sleep(20000);
				currentValue = new String();
				List<WebElement> currentValues = new ArrayList<WebElement>();
				currentValues = currentdriver.findElements(By.xpath(controlName));

				for (int j = 0; j < currentValues.size(); j++) {
					if (j + 1 == currentValues.size())
						currentValue = currentValue.concat(currentValues.get(j).getText());
					else {
						currentValue = currentValue.concat(currentValues.get(j).getText() + ",");
					}
				}
				break;
			default:
				break;

			}
			break;

		// Case Refresh to refresh login page in case of
		case Refresh:
			log.info("Refreshing login page");
			Automation.refreshPage(controlName);

		case Browser:
			Set<String> handlers = null;
			handlers = currentdriver.getWindowHandles();
			for (String handler : handlers) {
				currentdriver = currentdriver.switchTo().window(handler);

				// TM-19/01/2015: Changed following comparison from
				// equalsIgnoreCase to contains
				if (currentdriver.getTitle().contains(controlName)) {
					log.info("Focus on window with title: " + currentdriver.getTitle());
					break;
				}
			}
			break;

		case BrowserSwitchUsingURL:
			Set<String> handlersForSwitch = null;
			handlersForSwitch = currentdriver.getWindowHandles();
			for (String handler : handlersForSwitch) {
				currentdriver = currentdriver.switchTo().window(handler);

				if (currentdriver.getCurrentUrl().equalsIgnoreCase(controlName)) {
					log.info("Focus on window with URL: " + currentdriver.getCurrentUrl());
					break;
				}
			}
			break;

		case NewBrowser://
			switch (actionName) {
			case I:
				if (!ctrlValue.equalsIgnoreCase("null")) {
					String parentWindow = Automation.driver.getWindowHandle();
					log.info("Title of the parentWindow: " + Automation.driver.getTitle());
					log.info("Parent window handle: " + parentWindow);

					Set<String> handles = Automation.driver.getWindowHandles();
					for (String windowHandle : handles) {
						if (!windowHandle.equals(parentWindow)) {
							Automation.driver.switchTo().window(windowHandle);
							Thread.sleep(1000);
							log.info("Title of the new window: " + Automation.driver.getTitle());
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
			default:
				break;
			}
			break;//

		case URL:
			switch (actionName) {
			case I:
				currentdriver.navigate().to(ctrlValue);
				break;
			case NC:
				currentdriver.navigate().to(controlName);
				break;
			default:
				break;
			}
			break;

		case Menu:
			webElement.click();
			break;

		case Alert:
			switch (actionName) {
			case V:
				Alert alert = currentdriver.switchTo().alert();
				if (alert != null) {
					currentValue = alert.getText();
					log.info("Alert found on the web page");
					log.info(currentValue);
					alert.accept();
				}
				break;
			case NC:
				Alert alert1 = currentdriver.switchTo().alert();
				if (alert1 != null) {
					alert1.accept();
					Thread.sleep(2000);
				}
				break;
			default:
				break;
			}
			break;

		case WebImage:
			webElement.sendKeys(Keys.TAB);
			webElement.click();
			Thread.sleep(5000);
			for (int Seconds = 0; Seconds <= Integer.parseInt(Config.timeOut); Seconds++) {
				if (!((currentdriver.getWindowHandles().size()) > 1)) {
					webElement.click();
					Thread.sleep(5000);
				} else {
					break;
				}
			}
			break;

		case ActionClick:
			Actions builderClick = new Actions(currentdriver);
			Action clickAction = builderClick.moveToElement(webElement).clickAndHold().release().build();
			clickAction.perform();
			break;

		case ActionDoubleClick:
			Actions builderdoubleClick = new Actions(currentdriver);
			builderdoubleClick.doubleClick(webElement).build().perform();
			break;

		case ActionClickandEsc:
			Actions clickandEsc = new Actions(currentdriver);
			Action clickEscAction = clickandEsc.moveToElement(webElement).click().sendKeys(Keys.ENTER, Keys.ESCAPE)
					.build();
			clickEscAction.perform();
			break;

		case ActionMouseOver:
			Actions builderMouserOver = new Actions(currentdriver);
			builderMouserOver.moveToElement(webElement).perform();
			break;

		case Calendar:
			// Thread.sleep(5000);
			Boolean isCalendarDisplayed = currentdriver.switchTo().activeElement().isDisplayed();
			log.info(isCalendarDisplayed);
			if (isCalendarDisplayed == true) {
				String[] dtMthYr = ctrlValue.split("/");
				WebElement Year = WaitTool.waitForElement(currentdriver, By.name("year"),
						Integer.parseInt(Config.timeOut));// currentdriver.findElement(By.name("year"));
				while (!Year.getAttribute("value").equalsIgnoreCase(dtMthYr[2])) {
					if (Integer.parseInt(Year.getAttribute("value")) > Integer.parseInt(dtMthYr[2])) {
						WebElement yearButton = WaitTool.waitForElement(currentdriver, By.id("button1"),
								Integer.parseInt(Config.timeOut));// currentdriver.findElement(By.id("button1"));
						yearButton.click();
					} else if (Integer.parseInt(Year.getAttribute("value")) < Integer.parseInt(dtMthYr[2])) {
						WebElement yearButton = WaitTool.waitForElement(currentdriver, By.id("Button5"),
								Integer.parseInt(Config.timeOut));// currentdriver.findElement(By.id("Button5"));
						yearButton.click();
					}
				}
				Select date = new Select(
						WaitTool.waitForElement(currentdriver, By.name("month"), Integer.parseInt(Config.timeOut)));
				month = CalendarSnippet.getMonthForInt(Integer.parseInt(dtMthYr[1]));
				date.selectByVisibleText(month);
				WebElement Day = WaitTool.waitForElement(currentdriver, By.id("Button6"),
						Integer.parseInt(Config.timeOut));// currentdriver.findElement(By.id("Button6"));
				int day = 6;
				while (Day.getAttribute("value") != null) {
					Day = WaitTool.waitForElement(currentdriver, By.id("Button" + day),
							Integer.parseInt(Config.timeOut));// currentdriver.findElement(By.id("Button"+day));
					if (Day.getAttribute("value").toString().equalsIgnoreCase(dtMthYr[0])) {
						Day.click();
						break;
					}
					day++;
				}
			} else {
				log.info("Calendar not Diplayed");
			}
			break;

		case CalendarNew:
			isCalendarDisplayed = currentdriver.switchTo().activeElement().isDisplayed();
			log.info(isCalendarDisplayed);
			if (isCalendarDisplayed == true) {

				String[] dtMthYr = ctrlValue.split("/");
				Thread.sleep(2000);
				// String[] CurrentDate =
				// dtFormat.format(frmDate).split("/");
				WebElement Monthyear = currentdriver.findElement(By.xpath("//table/thead/tr/td[2]"));
				String Monthyear1 = Monthyear.getText();
				String[] Monthyear2 = Monthyear1.split(",");
				Monthyear2[1] = Monthyear2[1].trim();

				month = CalendarSnippet.getMonthForString(Monthyear2[0]);

				while (!Monthyear2[1].equalsIgnoreCase(dtMthYr[2])) {
					if (Integer.parseInt(Monthyear2[1]) > Integer.parseInt(dtMthYr[2])) {
						WebElement yearButton = currentdriver.findElement(By.cssSelector("td:contains('ï¿½')"));
						yearButton.click();
						Monthyear2[1] = Integer.toString(Integer.parseInt(Monthyear2[1]) - 1);
					} else if (Integer.parseInt(Monthyear2[1]) < Integer.parseInt(dtMthYr[2])) {
						WebElement yearButton = currentdriver.findElement(By.cssSelector("td:contains('ï¿½')"));
						yearButton.click();
						Monthyear2[1] = Integer.toString(Integer.parseInt(Monthyear2[1]) + 1);
					}
				}

				while (!month.equalsIgnoreCase(dtMthYr[1])) {
					if (Integer.parseInt(month) > Integer.parseInt(dtMthYr[1])) {
						WebElement monthButton = currentdriver.findElement(By.cssSelector("td:contains('ï¿½')"));
						monthButton.click();
						if (Integer.parseInt(month) < 11) {
							month = "0" + Integer.toString(Integer.parseInt(month) - 1);
						} else {
							month = Integer.toString(Integer.parseInt(month) - 1);
						}

					} else if (Integer.parseInt(month) < Integer.parseInt(dtMthYr[1])) {
						WebElement monthButton = currentdriver.findElement(By.cssSelector("td:contains('ï¿½')"));
						monthButton.click();
						if (Integer.parseInt(month) < 9) {
							month = "0" + Integer.toString(Integer.parseInt(month) + 1);
						} else {
							month = Integer.toString(Integer.parseInt(month) + 1);
						}
					}
				}

				WebElement dateButton = currentdriver
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
			String xpathYear = "//div[@class='datepicker datepicker-dropdown dropdown-menu datepicker-orient-left datepicker-orient-bottom']/div[@class='datepicker-years']";
			String xpathMonth = "//div[@class='datepicker datepicker-dropdown dropdown-menu datepicker-orient-left datepicker-orient-bottom']/div[@class='datepicker-months']";
			String xpathDay = "//div[@class='datepicker datepicker-dropdown dropdown-menu datepicker-orient-left datepicker-orient-bottom']/div[@class='datepicker-days']";

			// Selecting year in 3 steps
			currentdriver.findElement(By.xpath(xpathDay + "/table/thead/tr[1]/th[2]")).click();
			currentdriver.findElement(By.xpath(xpathMonth + "/table/thead/tr/th[2]")).click();
			currentdriver
					.findElement(By
							.xpath(xpathYear + "/table/tbody/tr/td/span[@class='year'][contains(text()," + year + ")]"))
					.click();

			// Selecting month in 1 step
			currentdriver.findElement(By.xpath(xpathMonth + "/table/tbody/tr/td/span[" + monthNum + "]")).click();

			// Selecting day in 1 step
			currentdriver
					.findElement(By.xpath(xpathDay + "/table/tbody/tr/td[@class='day'][contains(text()," + day + ")]"))
					.click();

		case CalendarEBP:
			String[] dtMthYrEBP = ctrlValue.split("/");
			Thread.sleep(2000);
			String yearEBP = dtMthYrEBP[2];
			String monthNumEBP = CalendarSnippet.getMonthForInt(Integer.parseInt(dtMthYrEBP[1])).substring(0, 3);
			String dayEBP = dtMthYrEBP[0];

			// common path used for most of the elements
			String pathToVisibleCalendar = "//div[@class='ajax__calendar'][contains(@style, 'visibility: visible;')]/div";

			// following is to click the title once to reach the year page
			wait.until(ExpectedConditions.elementToBeClickable(
					By.xpath(pathToVisibleCalendar + "/div[@class='ajax__calendar_header']/div[3]/div"))).click();
			// check if 'Dec' is visibly clickable after refreshing
			wait.until(ExpectedConditions.elementToBeClickable(
					By.xpath(pathToVisibleCalendar + "/div/div/table/tbody/tr/td/div[contains(text(), 'Dec')]")));
			// following is to click the title once again to reach the year
			// page
			currentdriver
					.findElement(By.xpath(pathToVisibleCalendar + "/div[@class='ajax__calendar_header']/div[3]/div"))
					.click();

			// common path used for most of the elements while selection of
			// year, month and date
			pathToVisibleCalendar = "//div[@class='ajax__calendar'][contains(@style, 'visibility: visible;')]/div/div/div/table/tbody/tr/td";

			// each of the following line selects the year, month and date
			wait.until(ExpectedConditions
					.elementToBeClickable(By.xpath(pathToVisibleCalendar + "/div[contains(text()," + yearEBP + ")]")))
					.click();
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath(pathToVisibleCalendar
					+ "/div[@class='ajax__calendar_month'][contains(text(),'" + monthNumEBP + "')]"))).click();
			wait.until(ExpectedConditions.elementToBeClickable(By.xpath(
					pathToVisibleCalendar + "/div[@class='ajax__calendar_day'][contains(text()," + dayEBP + ")]")))
					.click();

			break;

		/** Code for window popups **/
		case Window:
			switch (actionName) {
			case O:
				String parentHandle = currentdriver.getWindowHandle();
				for (String winHandle : currentdriver.getWindowHandles()) {
					currentdriver.switchTo().window(winHandle);
					if (currentdriver.getTitle().equalsIgnoreCase(controlName)) {
						currentdriver.close();
					}
				}
				currentdriver.switchTo().window(parentHandle);
				break;
			default:
				break;
			}
			break;

		case WebTable:
			switch (actionName) {
			case Read:
				WebHelperBilling.ReadFromExcel(ctrlValue);
				break;
			case Write:
				WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo, colNo);
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
								List<Object> elementProperties = WebHelperUtil.getPropertiesOfWebElement(
										tableColumns.get(Integer.parseInt(colNo)), imageType);
								controlName = elementProperties.get(0).toString();
								if (controlName.equals("")) {
									controlName = elementProperties.get(1).toString();
								}
								controlType = elementProperties.get(2).toString();
								webElement = (WebElement) elementProperties.get(3);
								doAction(FilePath, rowValues, testCase, imageType, controlType, controlId, controlName,
										ctrlValue, "", "", "", logicalName, action, webElement, Results, strucSheet,
										valSheet, tableRowIndex, rowcount, rowNo, colNo, operationType, cycleDate,
										TransactionType);
								break;
							}
						if (matchFound) {
							break;
						}
					}

				}
				break;
			case V:
				WebHelperUtil.WriteToDetailResults(ctrlValue, "", logicalName);
				try {
					Thread.sleep(2000);
				} catch (InterruptedException e) {
					log.error(e.getMessage(), e);
					e.printStackTrace();
				}
				break;

			case TableInput:

				String tableV = WebHelperUtil.checkTable(logicalName, rowValues, ExcelUtility.TIvaluesheetrows);

				if (!(tableV.equals(""))) {
					wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
					findtablefound = currentdriver.findElements(By.xpath(controlName)).size() > 0;
					if (findtablefound == true) {
						WebElement tableFound = wait
								.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName))); // Meghna--for
						BillingProduct.TableInputAction(tableFound, controlName, logicalName, rowValues, valuesHeader,
								ExcelUtility.TIvaluesheetrows);
						Thread.sleep(100);
					} else {
						log.info("Table not found. TABLE INPUT Functionality failed");
						break;
					}
				}
				break;

			// Modified the code to handle performance issues//
			case FIND:

				String findV = WebHelperUtil.CheckFind(logicalName, rowValues);

				if (!(findV.equals(""))) {
					Thread.sleep(5000);
					findtablefound = currentdriver.findElements(By.xpath(controlName)).size() > 0;

					if (findtablefound == true) {
						WebElement tableFound = wait
								.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));

						BillingProduct.findAction(tableFound, controlName, logicalName, rowValues, valuesHeader);
						Thread.sleep(1000);
					} else {
						log.info("Table not found. FIND Functionality failed");
						break;
					}
				}

				break;

			case I:
				Thread.sleep(10000);
				findtablefound = currentdriver.findElements(By.xpath(controlName)).size() > 0;
				if (findtablefound == true) {
					WebElement tableFound = wait
							.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
					List<WebElement> table_Rows = tableFound.findElements(By.tagName("tr"));
					List<WebElement> table_Columns = table_Rows.get(1).findElements(By.tagName("td"));
					// ApplicationtableRowsize = no of rows in the
					// WebTable
					int ApplicationtableRowsize = table_Rows.size();
					int Applicationtablecolumnsize = table_Columns.size();

					String ColumnName = ctrlValue.split(",")[0];
					String ColumnType = ctrlValue.split(",")[1];

					for (int i = 1; i <= Applicationtablecolumnsize; i++) {
						Thread.sleep(1000);
						String ApplicationColumnHeaderxapth = controlName + "/thead/tr/th[" + i + "]";
						log.info("ApplicationColumnHeader is:" + ApplicationColumnHeaderxapth);
						WebElement element = currentdriver.findElement(By.xpath(ApplicationColumnHeaderxapth));
						String ApplicationColumnHeader = element.getText();
						if ((ColumnName).equalsIgnoreCase(ApplicationColumnHeader)) {
							for (int r = 1; r <= ApplicationtableRowsize; r++) {
								if (ColumnType.equalsIgnoreCase("Webcheckbox")) {
									String XPath = controlName + "/tbody/tr[" + r + "]/td[" + i + "]/div/div/input";
									objfound = currentdriver.findElements(By.xpath(XPath)).size() > 0;
									if (objfound == true) {
										WebElement newelement = currentdriver.findElement(By.xpath(XPath));
										((JavascriptExecutor) currentdriver)
												.executeScript("arguments[0].scrollIntoView();", newelement);// Meghana
										// --
										newelement.click();
										Thread.sleep(500);
										((JavascriptExecutor) currentdriver)
												.executeScript("arguments[0].scrollIntoView();", newelement);
										Thread.sleep(500);
										objfound = false;
									}

								} else if (ColumnType.equalsIgnoreCase("WebLink")) {
									String XPath = controlName + "/tbody/tr[" + r + "]/td[" + i + "]/div/span";
									objfound = currentdriver.findElements(By.xpath(XPath)).size() > 0;
									if (objfound == true) {
										WebElement newelement = currentdriver.findElement(By.xpath(XPath));
										log.info("link xpath " + XPath);

										((JavascriptExecutor) currentdriver)
												.executeScript("arguments[0].scrollIntoView();", newelement);// Meghana
										// --
										((JavascriptExecutor) currentdriver).executeScript("arguments[0].click();",
												newelement);// Meghana
										Thread.sleep(500);
										Thread.sleep(500);
										objfound = false;
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

		// capture screenshot START (Billing)
		case Screenshot:
			switch (actionName) {
			case NC:
				cdate = null;
				if (StringUtils.isNotBlank(webDriver.getReport().getFromDate()))
					cdate = webDriver.getReport().getFromDate().replaceAll("[-/: ]", "");
				else
					webDriver.getReport().setFromDate(Config.dtFormat.format(new Date()));
				String transactionname = webDriver.getReport().getTrasactionType();

				if (ITAFWebDriver.isBillingApplication() || ITAFWebDriver.isSuiteApplication()
						&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Billing")) {
					String Screenshotcycledate = cycleDate.replace("/", "_");
					cfileName = webDriver.getReport().getTestcaseId() + "_" + transactionname + "_"
							+ Screenshotcycledate;
				} else {
					cfileName = webDriver.getReport().getTestcaseId() + "_" + transactionname;
				}

				clocation = Config.resultFilePath + "\\ScreenShots\\" + cfileName + "_" + screenshotnum + ".png";

				if (ITAFWebDriver.isSuiteApplication()) {
					String subfolder = Config.resultFilePath + "\\ScreenShots\\SelectiveScreenshots";
					File createDir = new File(subfolder);
					if (!createDir.exists()) {
						createDir.mkdirs();
					}
					clocation = subfolder + "\\" + cfileName + "_" + screenshotnum + ".png";
				}

				// TakesScreenshot scrShot = ((TakesScreenshot)webDriver);
				TakesScreenshot scrShot = ((TakesScreenshot) Automation.driver);
				File SrcFile = scrShot.getScreenshotAs(OutputType.FILE);

				File DestFile = new File(clocation);
				try {
					FileUtils.copyFile(SrcFile, DestFile);
				} catch (IOException e) {
					e.printStackTrace();
				}

				// image = new Robot().createScreenCapture(new
				// Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
				// ImageIO.write(image, "png", new File(clocation));
				screenshotnum = screenshotnum + 1;

				break;

			// Case Screenshot from State Farm Project
			case Screenshot:
				cdate = null;
				if (StringUtils.isNotBlank(webDriver.getReport().getFromDate()))
					cdate = webDriver.getReport().getFromDate().replaceAll("[-/: ]", "");
				else
					webDriver.getReport().setFromDate(Config.dtFormat.format(new Date()));

				String cycleDate_value3 = cycleDate_Values2.toString().replaceAll("/", "_");
				String strTestCaseID = controller.controllerTestCaseID.toString();
				String SC_NO[] = strTestCaseID.split("P");
				String strScNo = SC_NO[0].replace("_", "");
				cfileName = strScNo + "_" + strTestCaseID + "_" + cycleDate_value3 + "_" + cdate;
				log.info("Screenshot path : " + cfileName);
				clocation = System.getProperty("user.dir") + "\\" + Config.projectName + "\\Resources\\Screenshots\\"
						+ strScNo + "\\" + cfileName + ".jpg";

				if (ITAFWebDriver.isSuiteApplication()) {
					String subfolder = Config.resultFilePath + "\\ScreenShots\\SelectiveScreenshots";
					File createDir = new File(subfolder);
					if (!createDir.exists()) {
						createDir.mkdirs();
					}
					clocation = subfolder + "\\" + strScNo + "\\" + cfileName + ".jpg";
				}

				File screen_location = new File(clocation);

				while (screen_location.exists()) {
					i++;
					clocation = clocation.replace(".jpg", "");
					clocation = clocation + "_" + i;
					screen_location = new File(clocation);
					clocation = clocation + ".jpg";
				}
				screen_location = new File(clocation);

				File scrFile = ((TakesScreenshot) currentdriver).getScreenshotAs(OutputType.FILE);
				FileUtils.copyFile(scrFile, screen_location);
				break;

			default:
				break;
			}
			break;

		case Robot:
			if (controlName.equalsIgnoreCase("SetFilePath")) {
				StringSelection stringSelection = new StringSelection(ctrlValue);
				Toolkit.getDefaultToolkit().getSystemClipboard().setContents(stringSelection, null);
				robot.delay(1000);
				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_V);
				robot.keyRelease(KeyEvent.VK_V);
				robot.keyRelease(KeyEvent.VK_CONTROL);

			} else if (controlName.equalsIgnoreCase("TAB")) {
				Thread.sleep(1000);
				try {
					webElementForROBOT.sendKeys(Keys.TAB);
				} catch (Exception ex) {
					log.error(ex.getMessage(), ex);
					log.info("Object was not available");
				}
			} else if (controlName.equalsIgnoreCase("SPACE")) {
				robot.keyPress(KeyEvent.VK_SPACE);
				robot.keyRelease(KeyEvent.VK_SPACE);
			} else if (controlName.equalsIgnoreCase("ENTER")) {
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
				Thread.sleep(3000);
			}
			break;

		case DB:
			switch (actionName) {
			case Write:
				String policyNo = currentdriver.findElement(By.xpath(controlName)).getText();
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
				WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo, colNo);
				break;
			default:
				break;
			}
			break;

		case WaitForEC:
			wait.until(CommonExpectedConditions.elementToBeClickable(webElement));
			break;

		case WaitToLoad:
			log.info("Inside WaitToLoad");
			int waitSeconds = 0;
			if (Config.timeOut == null || Config.timeOut.equals("")) {
				Automation.driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(100));
				waitSeconds = 300;
			} else {
				Automation.driver.manage().timeouts()
						.pageLoadTimeout(Duration.ofSeconds(Integer.parseInt(Config.timeOut)));

				waitSeconds = Integer.parseInt(Config.timeOut) * 4;
			}
			WebDriverWait waitToload = new WebDriverWait(Automation.driver, Duration.ofSeconds(waitSeconds));
			switch (actionName) {
			case I:
				if (!ctrlValue.equalsIgnoreCase("")) {
					waitToload.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
					waitToload.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(controlName)));
					waitToload.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));

				}
				break;

			case NC:
				waitToload.until(ExpectedConditions.presenceOfElementLocated(By.xpath(controlName)));
				waitToload.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(controlName)));
				waitToload.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
				break;
			default:
				break;

			}
			log.info("Outside WaitToLoad");
			break;

		case SikuliScreen:
			App.open(sikscreen);
			break;
		case SikuliType:
			log.info("in sikulitype");
			log.info("controlName is:" + controlName);
			break;

		case SikuliButton:
			log.info("in sikuliButton");
			log.info("controlName is:" + controlName);
			log.info("Done");
			break;
		case Slider:
			WebElement slider = currentdriver.findElement(By.xpath(controlName));
			Thread.sleep(3000);
			Actions moveSlider = new Actions(currentdriver);
			Action actionslider = moveSlider.dragAndDropBy(slider, 30, 0).build();
			actionslider.perform();
			break;
		case MaskedInputDate:
			if (!ctrlValue.equalsIgnoreCase("null")) {
				webElement.clear();
				webElement.click();
				((JavascriptExecutor) currentdriver)
						.executeScript("arguments[0].setAttribute('value', '" + ctrlValue + "')", webElement);
				webElement.clear();
				Thread.sleep(1000);
				webElement.sendKeys(ctrlValue);
				webElement.sendKeys(Keys.TAB);
			} else {
				webElement.clear();
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
			((JavascriptExecutor) currentdriver)
					.executeScript("arguments[0].setAttribute('value', '" + ctrlValue + "')", webElement);
			// For handling Object level issue 01 June webElement.clear();
			Thread.sleep(1000);
			webElement.sendKeys(ctrlValue);
			break;

		case SendFileUpload: // New Case added - 25/08/2020
			switch (actionName) {
			case I:
				if (!ctrlValue.isEmpty()) {
					((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
							webElement);
					webElement.sendKeys(Config.inputDataFilePath + "\\" + ctrlValue);
				}
				break;

			default:
				break;
			}
			break;

		case ScrollTo:
			Locatable element = (Locatable) webElement;
			Point p = element.getCoordinates().onScreen();
			JavascriptExecutor js = (JavascriptExecutor) currentdriver;
			js.executeScript("window.scrollTo(" + p.getX() + "," + (p.getY() + 150) + ");");
			break;

		case Freeze:
			switch (actionName) {
			case I:
				if (logicalName.equalsIgnoreCase("GroupNo"))
					group_no = ctrlValue;
				else if (logicalName.equalsIgnoreCase("AccountNo"))
					account_no = ctrlValue;
				else if (logicalName.equalsIgnoreCase("JobStatus"))
					job_status = ctrlValue;
				else if (logicalName.equalsIgnoreCase("ExpectedStatus")) {
					expected_status = ctrlValue;

					try {
						Connection conn = JDBCConnection.establishHTML5BillingDBConn();
						Statement st = conn.createStatement();

						if (group_no != null) {
							st.execute("UPDATE job_schedule SET job_status = '" + job_status + "' WHERE job_status = '"
									+ expected_status
									+ "' AND group_system_code IN (select system_entity_code from entity_register where source_system_entity_code IN ("
									+ group_no + "))");
						}
						if (account_no != null) {
							st.execute("UPDATE job_schedule SET job_status = '" + job_status + "' WHERE job_status = '"
									+ expected_status
									+ "' AND Account_system_code IN (select system_entity_code from entity_register where source_system_entity_code IN ("
									+ account_no + "))");
						}
						log.info("Group No Freezed :" + group_no + account_no);
						webDriver.getReport().setMessage("Group No freezed : " + group_no + account_no);
						if (Config.databaseType.equalsIgnoreCase("ORACLE")) {
							st.execute("commit");
						}
						st.close();
						JDBCConnection.closeConnection(conn);
					}

					catch (Exception e) {
						log.error(e.getMessage(), e);
						throw new Exception("Error in RunDBQueries : " + e.getMessage());
					}
				}

				break;
			default:
				break;
			}
			// Case for Apex Archive ---Varsha 05/09/2018
		case Apex_Archive:
			switch (actionName) {
			case I:
				if (logicalName.equalsIgnoreCase("LocationNo"))
					LocNo = ctrlValue;
				else if (logicalName.equalsIgnoreCase("PolicyNo")) {
					PolNo = ctrlValue;

					try {
						Connection conn = JDBCConnection.establishHTML5BillingDBConn();
						Statement st = conn.createStatement();

						String query1 = "UPDATE entity_register SET"
								+ " source_system_entity_code=concat(source_system_entity_code, 1)"
								+ "WHERE APEX_ENTITY_CODE  =" + "(SELECT APEX_ENTITY_CODE " + "FROM entity_register "
								+ "WHERE source_system_entity_code='" + LocNo + "')";

						String query2 = "UPDATE policy_register SET" + " policy_no=concat(policy_no, 1)"
								+ "WHERE policy_no LIKE '" + PolNo + "'";

						if (LocNo != null) {
							st.execute(query1);
						}
						if (PolNo != null) {
							st.execute(query2);
						}
						log.info("Apex Archived : " + LocNo + " and respective policies also archived.");

						// Updating Report
						webDriver.getReport().setMessage("Location and Policies archived.");

						if (Config.databaseType.equalsIgnoreCase("ORACLE")) {
							st.execute("commit");
						}
						st.close();
						JDBCConnection.closeConnection(conn);
					}

					catch (Exception e) {
						log.info(e.getLocalizedMessage());
						webDriver.getReport().setStatus("FAIL");
						webDriver.getReport().setMessage(e.getLocalizedMessage());
					}
				}

				break;
			default:
				break;
			}
			break;

		case Task_Load:
			switch (actionName) {

			case I:
				if (logicalName.trim().equalsIgnoreCase("ModuleNO"))
					Module_No = Integer.parseInt(ctrlValue);

				else if (logicalName.trim().equalsIgnoreCase("UserNO"))
					User_No = Integer.parseInt(ctrlValue);

				else if (logicalName.equalsIgnoreCase("UserID1")) {
					User_ID1 = ctrlValue;
					UserList.add(User_ID1);
				} else if (logicalName.trim().equalsIgnoreCase("UserID2")) {
					User_ID2 = ctrlValue;
					UserList.add(User_ID2);
				} else if (logicalName.trim().equalsIgnoreCase("UserID3")) {
					User_ID3 = ctrlValue;
					UserList.add(User_ID3);
				} else if (logicalName.trim().equalsIgnoreCase("UserID4")) {
					User_ID4 = ctrlValue;
					UserList.add(User_ID4);
				} else if (logicalName.trim().equalsIgnoreCase("UserID5")) {
					User_ID5 = ctrlValue;
					UserList.add(User_ID5);
				} else if (logicalName.trim().equalsIgnoreCase("UserID6")) {
					User_ID6 = ctrlValue;
					UserList.add(User_ID6);
				} else if (logicalName.trim().equalsIgnoreCase("UserID7")) {
					User_ID7 = ctrlValue;
					UserList.add(User_ID7);
				} else if (logicalName.trim().equalsIgnoreCase("UserID8")) {
					User_ID8 = ctrlValue;
					UserList.add(User_ID8);
				} else if (logicalName.trim().equalsIgnoreCase("UserID9")) {
					User_ID9 = ctrlValue;
					UserList.add(User_ID9);
				} else if (logicalName.trim().equalsIgnoreCase("AllocatedCount")) {

					StringBuilder UserNameStr = new StringBuilder();

					try {

						Allocated_Count = Integer.parseInt(ctrlValue);

						int ac = 1, uniqueNo;
						String User_ID_Value = null;

						System.out.print(UserList.size());

						Connection conn = JDBCConnection.establishHTML5BillingDBConn();
						Statement st = conn.createStatement();

						// Iteration for all users to load with Task
						for (int ul = 0; ul < UserList.size(); ul++) {

							// ***Retrieve USER_ID from user_master ***
							// Varsha
							ResultSet UId = st.executeQuery("Select USER_ID from USER_MASTER where DISPLAY_USER_ID ='"
									+ UserList.get(ul) + "'");

							if (UId.next()) {
								User_ID_Value = UId.getString("USER_ID");
							}

							uniqueNo = ac;

							// //Iteration for No of allocation Count
							for (ac = uniqueNo; ac < uniqueNo + Allocated_Count; ac++) {

								st.execute(
										"INSERT INTO [dbo].RS_WORKITEMCACHE (WIR_ID, [version], SPECID, SPECURI, SPECVERSION, TASKID, RESOURCESTATUS)"
												+ " VALUES (" + Module_No + User_No + ac
												+ ", 1, 'UID', 'Task', 1, 'TID','Allocated')");

								st.execute(
										"INSERT INTO [dbo].RS_WORKITEMPARTICIPANT(WIR_PART_ID, PARTICIPANT_ID, WIR_ID) "
												+ "	VALUES(" + Module_No + User_No + ac + ", '" + User_ID_Value + "','"
												+ Module_No + User_No + ac + "')");
							}

							// Converting User List values to string
							UserNameStr.append(UserList.get(ul));
							if (ul != UserList.size() - 1) {
								UserNameStr.append(", ");
							}
							UId.close();
						}
						if (Config.databaseType.equalsIgnoreCase("ORACLE")) {
							st.execute("commit");
						}
						st.close();
						JDBCConnection.closeConnection(conn);
						// Updating Report
						webDriver.getReport().setMessage(
								"Users : " + UserNameStr + " Loaded With Allocation Count " + Allocated_Count);
						// Clear Userlist
						UserList.clear();

					}

					catch (Exception e) {

						webDriver.getReport().setStatus("FAIL");
						webDriver.getReport().setMessage(e.getLocalizedMessage());

					}

				}

				break;
			default:
				break;
			}

			break;
		case AP_Outbound:
		case FlatFileResponse:
		case FlatFile:
			switch (actionName) {
			case I:
				if (logicalName.equalsIgnoreCase("Local_Path")) {
					file_cycledate = ctrlValue;
					file_cycledate = file_cycledate.replace("/", "_");
					local_path = Config.inputDataFilePath + TransactionMapping.directoryPathFileUpload.toString()
							+ "\\FlatFiles\\" + ctrlValue;
					local_path = local_path.replace("/", "_");
				} else if (logicalName.equalsIgnoreCase("Remote_Path_Inbound")) {
					if (!(Config.flatFilePath != null))
						remote_path_in = ctrlValue;
					else
						remote_path_in = Config.flatFilePath + ctrlValue;
				} else if (logicalName.equalsIgnoreCase("File_To_Be_Converted"))
					file_to_be_converted = ctrlValue;
				else if (logicalName.equalsIgnoreCase("Extension"))
					// extension = file_cycledate + ctrlValue;
					extension = ctrlValue;
				else if (logicalName.equalsIgnoreCase("Remote_Path_Outbound")) {
					if (!(Config.flatFilePath != null))
						remote_path_out = ctrlValue;
					else
						remote_path_out = Config.flatFilePath + ctrlValue;
				} else if (logicalName.equalsIgnoreCase("ArchiveYN"))
					archive = ctrlValue;
				else if (logicalName.equalsIgnoreCase("XFL_FileName")) {
					xfl_filename = Config.inputDataFilePath + TransactionMapping.directoryPathFileUpload.toString()
							+ "\\FlatFiles\\" + ctrlValue;
					xfl_filename = xfl_filename.replace("/", "_");
				} else if (logicalName.equalsIgnoreCase("ReadFromDatabase") && !(ctrlValue.isEmpty())) {
					String requests_xml = local_path + "\\" + file_to_be_converted;
					requests_xml = requests_xml.replace("/", "_");
					Thread.sleep(1000);
					WebService.setXMLAPTagValue(requests_xml, "CheckStatusUpdate", "DisbursementDetailSeq", "PaymentId",
							"PaymentStatusDate", ctrlValue, 0);
				} else if (logicalName.equalsIgnoreCase("UniqueNo")) {
					uniqueFFNo = ctrlValue;
				} else if (logicalName.equalsIgnoreCase("ReportFilePath")) {
					reportFilePath = ctrlValue;
				} else if (logicalName.equalsIgnoreCase("CopyReportsToLocal")) {
					if (ctrlValue.equalsIgnoreCase("Y")) {
						try {
							log.info("Ready to Download all Bill XML !!");
							GenerateOutputXML.main(null);
						} catch (Exception e) {
							log.error(e.getMessage(), e);
							e.printStackTrace();
						}
					}
				}
				break;

			case ArchiveFiles:
				if (ctrlValue.equalsIgnoreCase("Y")) {
					if (Config.EnvironmentType.equalsIgnoreCase("Linux")) {//Selenium 4- 23/10/2024
						try {
							ArchiveFiles af = new ArchiveFiles();
							if (!(reportFilePath.equals(null) | reportFilePath.equals(""))) {
								reportFilePath = Config.invoiceProcessingRemotePath.toString();
								af.archiveLinux(reportFilePath, extension);
							} else {
								af.archiveLinux(remote_path_out, extension);
							}
						} catch (Exception e) {
							log.error(e.getMessage(), e);
							e.printStackTrace();
						}
					} else {
						ArchiveFiles af = new ArchiveFiles();
						af.archiveWindows(remote_path_out, extension);
					}
				}
				break;

			case UpdateXML:
				if (!ctrlValue.trim().equals("") && ITAFWebDriver.isSuiteApplication()) {
					String request_xml_temp = local_path + file_to_be_converted + ".xml";

					String[] tagValuePair = new String[2];
					tagValuePair = ctrlValue.split(";;");
					if (tagValuePair[1].contains("pickFromUniqueNumbers")) {
						String curValue = WebHelperUtil.ReadFromExcel("pickFromUniqueNumbers", WebHelper.columnName);
						tagValuePair[1] = tagValuePair[1].replace("pickFromUniqueNumbers", curValue);
					}
					WebHelperUtil.updateXmlTagValue(request_xml_temp, tagValuePair[0], "NA", tagValuePair[1], 1);
				}
				break;

			case CopyFiles:

				if (ctrlValue.equalsIgnoreCase("Y")) {
					InboundFileTransfer cfr = new InboundFileTransfer();
					if (uniqueFFNo.equals("") | uniqueFFNo.equals(null))
						cfr.copyFilesToRemote_Nav(local_path, remote_path_in, file_to_be_converted,
								("_" + file_cycledate + extension), xfl_filename);
					else
						cfr.copyFilesToRemote_Nav(local_path, remote_path_in, file_to_be_converted,
								("_" + uniqueFFNo + extension), xfl_filename);
				}

				break;

			case ConvertToXml:
				if (ctrlValue.equalsIgnoreCase("Y")) {
					validateTag = ctrlValue1Cell.toString();
					validationMsg = ctrlValue2Cell.toString();
					OutboundTransferFiles cf = new OutboundTransferFiles();
					cf.copyFiles_Nav(remote_path_out, local_path, file_to_be_converted, extension, xfl_filename,
							archive, validateTag, validationMsg, file_cycledate);
				}
				break;
			default:
				break;
			}
			break;

		case CopyFlatFile_Text:
			switch (actionName)

			{
			case I:
				if (logicalName.equalsIgnoreCase("DestinationFolder"))
					destination_folder = ctrlValue;
				else if (logicalName.equalsIgnoreCase("FileName")) {
					file_names = ctrlValue;

					try {
						cycleDate = cycleDate.replace("/", "_");
						SourceFlatFile = (Config.inputDataFilePath + "FlatFile\\FlatFiles\\");
						DestinationFlatFile = (Config.flatFilePath + "\\" + destination_folder);
						String[] files = file_names.split("\\|");
						String filesnames = "";

						for (int k = 0; k < files.length; k++) {
							file_name = "";
							file_name = files[k];

							File src = new File(SourceFlatFile + "/" + cycleDate + "/" + file_name);
							File dest = new File(DestinationFlatFile + "/" + file_name);

							if (src.exists()) {
								FileUtils.copyFile(src, dest);
								filesnames = filesnames + " " + file_name;
							} else {
								webDriver.getReport().setMessage(
										"Directory: " + SourceFlatFile + "/" + cycleDate + "/" + "not found");
								log.error("Directory: " + SourceFlatFile + "/" + cycleDate + "/" + "not found");
								throw new IOException(
										"Directory: " + SourceFlatFile + "/" + cycleDate + "/" + "not found");

							}

						}
						log.info("Files Copied :" + filesnames);
						webDriver.getReport().setMessage("Files Copied :" + filesnames);

					} catch (IOException ff) {
						webDriver.getReport().setMessage(ff.getMessage());
						log.error("File not found in path: " + SourceFlatFile + "/" + cycleDate + "/" + file_name
								+ " <-|-> LocalizeMessage " + ff.getLocalizedMessage() + " <-|-> Message "
								+ ff.getMessage() + " <-|-> Cause " + ff.getCause(), ff);
						throw new IOException("File not found in path: " + SourceFlatFile + "/" + cycleDate + "/"
								+ file_name + " <-|-> LocalizeMessage " + ff.getLocalizedMessage() + " <-|-> Message "
								+ ff.getMessage() + " <-|-> Cause " + ff.getCause());
					}
				}
				break;
			default:
				break;
			}
			break;

		case CopyFlatFile:
			switch (actionName)

			{
			case I:
				if (logicalName.equalsIgnoreCase("DestinationFolder"))
					destination_folder = ctrlValue.trim();
				else if (logicalName.equalsIgnoreCase("FileName")) {
					file_names = ctrlValue.trim();

					if (Config.copyServerRemotePathLinux != null && !Config.copyServerRemotePathLinux.trim().isEmpty()) {
						try {
							cycleDate = cycleDate.replace("/", "_");
							SourceFlatFile = (Config.inputDataFilePath + "CopyFlatFile\\FlatFiles\\");
							DestinationFlatFile = (Config.copyServerRemotePathLinux + "/" + destination_folder);

							List<String> allFileNames = new ArrayList<>();
							String[] files = file_names.split("\\|");
							String filesnames = "";

							for (int k = 0; k < files.length; k++) {
								file_name = "";
								file_name = files[k];

								File src = new File(SourceFlatFile + "/" + cycleDate + "/" + file_name);
								// File dest = new File(DestinationFlatFile + "/" + file_name);

								log.info("Performing check to verify if file" + " " + file_name + " "
										+ "exists in the source folder...");
								if (src.exists()) {
									filesnames = filesnames + " " + file_name;
									allFileNames.add(file_name);

								} else {

									webDriver.getReport().setMessage("(" + " Directory: " + SourceFlatFile + cycleDate
											+ "\\" + file_name + ")" + "not found");
									log.info("(" + "Directory: " + SourceFlatFile + cycleDate + "\\" + file_name + ")"
											+ "not found");
									throw new IOException("(" + "Directory: " + SourceFlatFile + cycleDate + "\\"
											+ file_name + ")" + "not found");
								}
							}

							log.info("All files exists in the source folder");
							// log.info("Files Copied :" + allFileNames);

							log.info("Initiating file transfer process over the Linux Server...");

							FTPFileTransferLinux.FTPLinux(filesnames, DestinationFlatFile, cycleDate);

							webDriver.getReport().setMessage("Files Copied Sucessfully:" + filesnames);

						} catch (IOException ff) {

							webDriver.getReport().setMessage(ff.getMessage());

							log.error("File not found in path: " + SourceFlatFile + cycleDate + "\\" + file_name);
							throw new IOException(
									"File not found in path: " + SourceFlatFile + cycleDate + "\\" + file_name);

						}
					}
				} else {

					try {
						cycleDate = cycleDate.replace("/", "_");
						SourceFlatFile = (Config.inputDataFilePath + "CopyFlatFile\\FlatFiles\\");
						DestinationFlatFile = (Config.copyServerRemotePath + "\\" + destination_folder);
						String[] files = file_names.split("\\|");
						String filesnames = "";

						for (int k = 0; k < files.length; k++) {
							file_name = "";
							file_name = files[k];

							File src = new File(SourceFlatFile + "/" + cycleDate + "/" + file_name);
							File dest = new File(DestinationFlatFile + "/" + file_name);

							if (src.exists()) {
								FileUtils.copyFile(src, dest);
								filesnames = filesnames + " " + file_name;
							} else {

								webDriver.getReport().setMessage("(" + " Directory: " + SourceFlatFile + cycleDate
										+ "\\" + file_name + ")" + "not found");
								log.error("(" + "Directory: " + SourceFlatFile + cycleDate + "\\" + file_name + ")"
										+ "not found");
								throw new IOException("(" + "Directory: " + SourceFlatFile + cycleDate + "\\"
										+ file_name + ")" + "not found");
							}
						}

						log.info("Files Copied :" + filesnames);
						webDriver.getReport().setMessage("Files Copied :" + filesnames);

					} catch (IOException ff) {

						webDriver.getReport().setMessage(ff.getMessage());

						log.error("File not found in path: " + SourceFlatFile + cycleDate + "\\" + file_name);
						throw new IOException(
								"File not found in path: " + SourceFlatFile + cycleDate + "\\" + file_name);

					}
				}

				;
				break;
			default:
				break;
			}
			break;

		case CopyInterfaceFile:
			switch (actionName) {
			case I:
				if (logicalName.equalsIgnoreCase("InterfaceFolder")) {
					interface_folder = ctrlValue;

					try {
						cycleDate = cycleDate.replace("/", "_");
						String testCaseID = controller.controllerTestCaseID.toString();
						SourceFlatFile = (Config.InterfaceOutBoundPath + "\\" + interface_folder);
						DestinationFlatFile = (Config.inputDataFilePath + "OutBound\\" + interface_folder + cycleDate);

						File OutBoundFiles = new File(SourceFlatFile);
						File[] InterfaceFiles = OutBoundFiles.listFiles();

						for (i = 0; i < InterfaceFiles.length; i++) {
							String fName = InterfaceFiles[i].getName();
							File src = new File(SourceFlatFile + fName);
							File dest = new File(DestinationFlatFile + "\\" + testCaseID + "_Actual_" + i + ".csv");
							log.info("File name is: " + fName);
							if (fName.contains(".csv")) {
								FileUtils.copyFile(src, dest);
								log.info("File copied successfully:" + dest);
							}
						}
						log.info("Files copied in " + interface_folder + "folder successfully.");
						webDriver.getReport()
								.setMessage("Files copied in " + interface_folder + "folder successfully.");

					} catch (IOException ff) {
						webDriver.getReport().setMessage(ff.getMessage());
						log.error("Files not copied in " + interface_folder + "folder.");
						throw new IOException("Files not copied in " + interface_folder + "folder.");
					}

				} else if (logicalName.equalsIgnoreCase("CompareFile")) {
					compare_file = ctrlValue;
					log.info("CompareFile value is set to :" + compare_file);

				} else if (logicalName.equalsIgnoreCase("ServerFolder")) {
					server_folder = ctrlValue;

					try {
						String serverFlatFilePath = null;
						String localFlatFileLocation = null;
						String expectedFlatFile = null;

						String testCaseID = controller.controllerTestCaseID.toString();

						localFlatFileLocation = Config.inputDataFilePath + "Claim/CopyFlatFile/FlatFiles/" + cycleDate;
						serverFlatFilePath = Config.copyServerRemotePath + "\\" + server_folder;
						expectedFlatFile = Config.inputDataFilePath + "Claim/CopyFlatFile/FlatFiles/" + cycleDate + "\\"
								+ testCaseID + "_Expected.txt";

						File OutBoundFiles = new File(serverFlatFilePath);

						// Below given condition added to check if server folder access is available
						if (OutBoundFiles.exists() && OutBoundFiles.isDirectory()) {
							File[] InterfaceFiles = OutBoundFiles.listFiles();

							// Sort InterfaceFiles based on the last modified timestamp in descending order
							// (newest first)
							Arrays.sort(InterfaceFiles, Comparator.comparingLong(File::lastModified).reversed());

							if (InterfaceFiles.length > 0) {
								File latestFile = InterfaceFiles[0];
								String latestFileName = latestFile.getName();

								File src = new File(serverFlatFilePath + "\\" + latestFileName);
								File dest = new File(localFlatFileLocation + "\\" + latestFileName);

								// check if latest file name already exist
								if (dest.exists()) {
									// The file already exists in the destination path, so we need to rename it.
									String baseName = latestFileName.substring(0, latestFileName.lastIndexOf('.'));
									String extension = latestFileName.substring(latestFileName.lastIndexOf('.'));

									int count = 1;
									File renamedDest;
									while ((renamedDest = new File(
											localFlatFileLocation + "\\" + baseName + "_" + count + extension))
											.exists()) {
										count++;
									}

									// Perform the renaming of the existing file to the new name
									try {
										Path sourcePath = dest.toPath();
										Path targetPath = renamedDest.toPath();
										Files.move(sourcePath, targetPath, StandardCopyOption.REPLACE_EXISTING);
										log.info("Existing Flat File renamed successfully.");
									} catch (Exception e) {
										log.info("Error renaming the file: " + e.getMessage());
									}

									// Perform the file copying operation
									try {
										Files.copy(src.toPath(), dest.toPath(), StandardCopyOption.REPLACE_EXISTING);
										log.info(
												"Latest file as per time stamp copied successfully: " + dest.getName());
									} catch (Exception e) {
										log.info("Error copying the file: " + e.getMessage());
										webDriver.getReport().setMessage("Error copying the file: " + e.getMessage());
										webDriver.report.setStatus("FAIL");
									}
								}

								// Perform the file copying operation
								Files.copy(src.toPath(), dest.toPath(), StandardCopyOption.REPLACE_EXISTING);
								log.info("Latest file as per time stamp copied successfully: " + dest.getName());

								// Convert File object to String representation
								String destPath = dest.getAbsolutePath();

								// To remove Trailing spaces
								RemoveExtraSpaces.removeSpace(destPath);
								log.info("Trailing spaces removed successfully from Actual Flat File");

								// Convert the path to a Path object
								Path path = Paths.get(destPath);

								// Extract the filename (without the extension) from the 'destPath' string
								String filenameWithoutExt = path.getFileName().toString();
								int dotIndex = filenameWithoutExt.lastIndexOf('.');
								if (dotIndex > 0) {
									filenameWithoutExt = filenameWithoutExt.substring(0, dotIndex);
								}

								// Append '_New.txt' to the filename
								String newFilename = filenameWithoutExt + "_Actual.txt";

								// Construct the new path with the new filename
								Path newPath = path.resolveSibling(newFilename);

								// log.info(newPath.toString());

								if (compare_file.equalsIgnoreCase("Y")) {

									Path expectedFlatFilePath = Paths.get(expectedFlatFile);
									if (Files.exists(expectedFlatFilePath)) {

										String expectedfile = expectedFlatFilePath.toString();
										RemoveExtraSpaces.removeSpace(expectedfile);
										log.info("Trailing spaces removed successfully from Expected Flat File");

										// Compare expected FF with Actual FF
										FlatFileComparison.FlatFileCompare(newPath.toString(), expectedFlatFile,
												localFlatFileLocation, testCaseID);

									} else {
										// If the expectedFlatFile does not exist, log and report the message
										log.info("Expected flat file does not exist: " + expectedFlatFile);
										webDriver.getReport()
												.setMessage("Expected flat file does not exist: " + expectedFlatFile);
									}

								} else {
									log.info("File Comparison is set to 'N'");
									webDriver.getReport().setMessage(
											"Transaction PASS - Latest file copied successfully, as File Comparison is set to 'N' in the CopyFlatfile sheet Flat file comparison will not be performed");
									webDriver.report.setStatus("PASS");
								}
							} else {
								log.info("No files found in the server folder: " + server_folder);
								webDriver.getReport()
										.setMessage("No files found in the server folder: " + server_folder);
							}

						} else {
							log.info("Server folder does not exist or access permission are denied: "
									+ serverFlatFilePath);
							webDriver.getReport().setMessage("Server folder does not exist: " + serverFlatFilePath);
							webDriver.report.setStatus("FAIL");
							break;
						}

					}

					catch (IOException e) {
						webDriver.getReport().setMessage(e.getMessage());
						log.error("Error occurred while copying files from " + server_folder + " folder.");
						throw new IOException("Error occurred while copying files from " + server_folder + " folder.",
								e);
					}
				}

				break;
			default:
				break;
			}
			break;

		case ReplaceDynamicText:
			switch (actionName) {
			case I:
				if (logicalName.equalsIgnoreCase("RequestJSON"))
					requestJson = ctrlValue;
				else if (logicalName.equalsIgnoreCase("TagName"))
					xmlTagName = ctrlValue;
				else if (logicalName.equalsIgnoreCase("CharacterShiftCount"))
					charCount = Integer.parseInt(ctrlValue);
				else if (logicalName.equalsIgnoreCase("StartIndexSearch"))
					startIndexSearch = ctrlValue;
				else if (logicalName.equalsIgnoreCase("EndIndexSearch")) {
					endIndexSearch = ctrlValue;

					cycleDate = cycleDate.replace("/", "_");
					String expectedXmlFilePath = Config.inputDataFilePath + "Restful\\RestfulFiles\\" + cycleDate + "\\"
							+ requestJson + "_Expected.xml";
					WebHelperUtil.replaceText(expectedXmlFilePath, xmlTagName, charCount, startIndexSearch,
							endIndexSearch);
					String responseXmlFilePath = Config.inputDataFilePath + "Restful\\RestfulFiles\\" + cycleDate + "\\"
							+ requestJson + "_Response.xml";
					WebHelperUtil.replaceText(responseXmlFilePath, xmlTagName, charCount, startIndexSearch,
							endIndexSearch);
				}
				break;
			default:
				break;
			}

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
			switch (actionName) {
			case I:
				if (logicalName.equalsIgnoreCase("WSDL_URL")) {
					wsdl_url = ctrlValue;
					log.info("WSDL Name:" + wsdl_url);
					if (wsdl_url.toLowerCase().contains("${host:port}")) {
						wsdl_url = wsdl_url.replace("${host:port}", Config.appHostAppPort);
					}

				}
				if (logicalName.equalsIgnoreCase("REQUEST_URL")) {
					request_url = ctrlValue;
					log.info("Request XML Name :" + request_url);

				}
				if (logicalName.equalsIgnoreCase("ValidateResponseCheck")) {
					validateResponseFlag = ctrlValue;
					if (validateResponseFlag.startsWith("Y")) {
						validateResponse = true;
						log.info("validateResponseFlag is :" + validateResponseFlag);
					} else {
						validateResponse = false;
						log.info("validateResponseFlag is : N");
					}

				}
				if (logicalName.equalsIgnoreCase("REQUEST_XML")) {
					request_xml = ctrlValue;
					log.info("Request XML Name :" + request_xml);

					// if(TransactionType.toString().equalsIgnoreCase("WebServiceC"))
					if (!TransactionType.toString().contains("WebServiceV")) {

						String request_xml = Config.inputDataFilePath
								+ TransactionMapping.directoryPathFileUpload.toString() + "\\WebserviceFiles\\"
								+ cycleDate + "\\" + ctrlValue + ".xml";
						request_xml = request_xml.replace("/", "_");
						Thread.sleep(1000);
						log.info("Generating new unique SourceSystemRequestNo... ");
						WebService.setXMLResponseTagValue(request_xml, "RequestHeader", "SourceSystemRequestNo", 0);
						log.info("Completed... ");
						// break;
					} else {
						log.info("Transaction is - " + TransactionType);
					}
				}

				// For Check Status Update
				if (logicalName.equalsIgnoreCase("ReadFromDatabase") && !(ctrlValue.isEmpty())) {

					String requests_xml = Config.inputDataFilePath
							+ TransactionMapping.directoryPathFileUpload.toString() + "\\WebserviceFiles\\" + cycleDate
							+ "\\" + request_xml + ".xml";
					requests_xml = requests_xml.replace("/", "_");
					Thread.sleep(1000);
					WebService.setXMLAPTagValue(requests_xml, "CheckStatusUpdate", "DisbursementDetailSeq", "PaymentId",
							"PaymentStatusDate", ctrlValue, 0);
				}

				break;

			case UpdateXML:

				if (!ctrlValue.trim().equals("")) {

					String request_xml_temp = Config.inputDataFilePath
							+ TransactionMapping.directoryPathFileUpload.toString() + "\\WebserviceFiles\\" + cycleDate
							+ "\\" + request_xml + ".xml";

					request_xml_temp = request_xml_temp.replace("/", "_");

					if (!ctrlValue.contains("pickFromUniqueNumbers")) {
						// if (!logicalName.isEmpty()) {
						String TagName = logicalName;
						String Value = ctrlValue;
						log.info("TageName is:" + TagName);
						log.info("Value is:" + Value);

						WebHelperUtil.updateXmlTagValue(request_xml_temp, TagName, "NA", Value, 1);

					} else {

						String[] tagValuePair = new String[2];
						tagValuePair = ctrlValue.split(";;");

						try {

							if (tagValuePair[0].contains("pickFromUniqueNumbers")) {

								String curValue = WebHelperUtil.ReadFromExcel("pickFromUniqueNumbers",
										WebHelper.columnName);
								tagValuePair[0] = tagValuePair[0].replace("pickFromUniqueNumbers", curValue);

								String tagValuePairValue = tagValuePair[0];
								// tagValuePairValue = tagValuePairValue.substring(1, tagValuePairValue.length()
								// - 1); // Remove the square brackets []

								String[] splitValues = tagValuePairValue.split(";");
								String variable1 = splitValues[0];
								String variable2 = splitValues[1];

								// log.info(variable1);
								// log.info(variable2);

								WebHelperUtil.updateXmlTagValue(request_xml_temp, variable1, "NA", variable2, 1);
							} else {
								log.info("Not enough elements in the array");
							}

						} catch (Exception e) {
							log.info(e);
							log.info("Text 'pickFromUniqueNumbers' not found ");

						}

					}

				} else {
					log.info("Mistmacth between logical Name and Column header in structure sheet ");

				}

				break;

			case T:
				faultstring = false;
				success = false;
				description = null;
				pathtoNode = ctrlValue1;
				column_Name = ctrlValue2;

				if ((pathtoNode == null && column_Name == null) || (pathtoNode == "" && column_Name == "")) {
					log.info("PathtoNode and ColumnName are empty, If validate Response Flag is set to 'Y'"
							+ "Expected Response with be compared with Actual....");
				} else {

					log.info("path to Node:" + pathtoNode);
					log.info("column names :" + column_Name);
				}

				log.info("Generating XML Response...");
				responseXml = WebService.callWebService(wscycledate, wsdl_url, request_xml, Config.user,
						Config.password);

				if (validateResponse == true) {

					if ((pathtoNode != null && column_Name != null) && (pathtoNode != "" && column_Name != "")) {

						String responsexmlFilePath = Config.inputDataFilePath
								+ TransactionMapping.directoryPathFileUpload.toString() + "\\WebserviceFiles\\"
								+ cycleDate + "\\" + request_xml + "_Response.xml";

						responsexmlFilePath = responsexmlFilePath.replace("/", "_");
						String responseJsonPath = XML2JsonCompare.convertXmlToJson(responsexmlFilePath);
						String responseJson = "";

						responseJson = new String(Files.readAllBytes(Paths.get(responseJsonPath)));
						boolean jsonValidationSuccess = false;

						jsonValidationSuccess = JsonUtility.validateFailedJsonContent(responseJson, pathtoNode,
								column_Name);
						if (jsonValidationSuccess == true) {

							restErrorResDesc = "  => Response Validation Successful"
									+ " => All messages from expected response - " + " ' " + column_Name + " ' "
									+ " are present in actual response";

							webDriver.report.setStatus("PASS");
							webDriver.report.setMessage(restErrorResDesc);
							log.info("Response Validation Successful");

							pathToNode = null;
							nodeDescription = null;
							validateResponse = false;
						} else {

							restErrorResDesc = "  => Expected values not matching Actual ";

							webDriver.report.setStatus("FAIL");
							webDriver.report.setMessage(restErrorResDesc);
							log.info("Response Validation Failed");
						}

					} else {

						String comparisonResultFilePath = null;
						@SuppressWarnings("unused")
						boolean xmlComparisonSuccess = false;
						// validate response and compare it with expected response json file.
						List<String> validateJsonIgnoreKeyListWebservice = null;
						String expectedxmlFilePath = Config.inputDataFilePath
								+ TransactionMapping.directoryPathFileUpload.toString() + "\\WebserviceFiles\\"
								+ cycleDate + "\\" + request_xml + "_Expected.xml";

						expectedxmlFilePath = expectedxmlFilePath.replace("/", "_");

						String responsexmlFilePath = Config.inputDataFilePath
								+ TransactionMapping.directoryPathFileUpload.toString() + "\\WebserviceFiles\\"
								+ cycleDate + "\\" + request_xml + "_Response.xml";

						responsexmlFilePath = responsexmlFilePath.replace("/", "_");

						comparisonResultFilePath = (Config.inputDataFilePath
								+ TransactionMapping.directoryPathFileUpload.toString() + "\\WebserviceFiles\\"
								+ cycleDate + "\\" + request_xml + "_Result.xlsx").replace("/", "_");

						if (Config.validateResponseIgnoreKeysWebservice != null)
							validateResponseIgnoreKeys = Config.validateResponseIgnoreKeysWebservice;
						validateJsonIgnoreKeyListWebservice = Arrays
								.asList(validateResponseIgnoreKeys.trim().split("\\s*,\\s*"));
						log.info("Ignore parameters are : " + validateJsonIgnoreKeyListWebservice);
						
						validateResponseIgnoreKeys = "";
						File expectedXmlFile = new File(expectedxmlFilePath);
						if (!expectedXmlFile.exists()) {
							log.info("Expected XML file does not exist.");
							// Add any necessary clean-up code or additional handling

							webDriver.getReport().setMessage("Transaction Failed As Expected XML is Missing");
							webDriver.getReport().setStatus("FAIL");
							// ...
							break; // Break and come out of the code
						}
						log.info("Converting Expected XML File to Expected Json File... ");
						String expectedJsonPath = XML2JsonCompare.convertXmlToJson(expectedxmlFilePath);
						log.info("Converting Actual Response XML File to Actual Response Json File... ");
						String responseJsonPath = XML2JsonCompare.convertXmlToJson(responsexmlFilePath);

						boolean jsonValidationSuccess = false;
						if (!StringUtils.isEmpty(responseJsonPath)) {
							if (pathToNode != null && nodeDescription != null) {

								jsonValidationSuccess = JsonUtility.validateFailedJsonContent(restResponse, pathToNode,
										nodeDescription);
							} else {
								// success = JsonUtility.validateRestJson(restResponse);
								try {
									log.info(
											"Comparing Expected Json to Actual Response Json, Ignore parameters Included if mentioned in the config sheet");
									jsonValidationSuccess = ValidateJson.main(expectedJsonPath, responseJsonPath,
											validateJsonIgnoreKeyListWebservice);
									log.info("Comparison result: " + jsonValidationSuccess);
								} catch (Exception e) {
									log.info("An error occurred: " + e.getMessage());
								}
							}
							// jsonValidationSuccess = true;
						} else {
							log.info("Rest response returned is null or empty");
						}
						if (jsonValidationSuccess == true) {

							restErrorResDesc = "  => Response Validation Successful"
									+ " => All messages from expected response are present in actual response" + " => "
									+ "List Of Ignored Keys :" + validateJsonIgnoreKeyListWebservice;

							webDriver.report.setStatus("PASS");
							webDriver.report.setMessage(restErrorResDesc);
							log.info("COMPARISON PASSED");
						} else {

							restErrorResDesc = "  => Expected results not matching Actual: " + " => "
									+ "ComparisonResultFilePath :" + " => " + comparisonResultFilePath;
							webDriver.report.setStatus("FAIL");
							webDriver.report.setMessage(restErrorResDesc);
							log.info("COMPARISON FAILED");
						}
					}

				} else {
					log.info("validateResponse is set to false...");
					
					//New condition added for Billing WebserviceV transactions as they are only meant for comparison
					 if (controller.controllerTransactionType.toString().startsWith("WebServiceV") && !validateResponse) {
				            webDriver.getReport().setMessage("Marking the " + controller.controllerTransactionType + " transaction Fail as the ValidateResponsecheck tag is missing or is set to N in the Inputsheet");
				            webDriver.getReport().setStatus("FAIL");
				            resetVariables();// Sel4
				            break; // Exit the case block
				        }

					String InBoundInstrumentInformation = "InBoundInstrumentInformation";
					if (InBoundInstrumentInformation.contains(wsdl_url)
							&& ((!pathtoNode.equalsIgnoreCase("null") && !pathtoNode.equalsIgnoreCase(""))
									|| (!column_Name.equalsIgnoreCase("null") && !column_Name.equalsIgnoreCase(""))))

						WebService.getXMLResponseData(ctrlValue1, ctrlValue2, testCase, wscycledate, responseXml);

					if (!controller.controllerTransactionType.toString().contains("WebServiceV")) {
						// String Tag_Name = "EntityResponse";
						String Tag_Name = "*";
						String[] Node_Value = new String[2];

						if ((ctrlValue1 != "" && ctrlValue2 != "") && (ctrlValue1 != null && ctrlValue2 != null)) {
							Node_Value[0] = ctrlValue1;
							Node_Value[1] = ctrlValue2;
						} else {
							Node_Value[0] = "ProcessStatusFlag";
							Node_Value[1] = "SuccessFlag";
						}
						String Node_Value2 = "faultstring";
						String Node_Value3 = "Description";
						int index = 0;
						WebServiceResponse = WebService.getXMLResponseStatus(responseXml, Tag_Name, Node_Value, index);

						if (WebServiceResponse.equalsIgnoreCase("SUCCESS")) {
							webDriver.getReport().setMessage("SUCCESS : Matched -- " + " ' " + Node_Value[1] + " ' ");// Mandar
							webDriver.getReport().setStatus("PASS");
							success = true;
							description = "SUCCESS";
						}

						else if (WebServiceResponse.equalsIgnoreCase("FAILED") && success != true) {
							FailedResponseTagValue = WebService.getXMLResponseTagValue(responseXml, Tag_Name,
									Node_Value3, index);
							if (webDriver.getReport().getMessage() == null
									|| webDriver.getReport().getMessage() == "") {
								// to do//Mandar --20/09/2017
							}
							failed = true;
						} else if (WebServiceResponse == null && failed != true) {
							faultstring = true;
							FailedResponseTagValue = WebService.getXMLResponseTagValue(responseXml, Tag_Name,
									Node_Value2, index);
							webDriver.getReport()
									.setMessage("REQUEST FAILED : Error Msg displayed -- " + FailedResponseTagValue);// Mandar--
							webDriver.getReport().setStatus("FAIL");
							description = FailedResponseTagValue;
						}

						else if (WebServiceResponse.equalsIgnoreCase("BLANK") && nullvalue != true
								&& WebService.isNoResponseFileTrue()) {

							webDriver.getReport().setMessage("BLANK WEBSERVICE RESPONSE");
							webDriver.getReport().setStatus("FAIL");

							description = "BLANK WEBSERVICE RESPONSE";
						}
						if (WebServiceResponse == null || WebServiceResponse.equalsIgnoreCase("FAILED")
								|| WebServiceResponse.equalsIgnoreCase("BLANK")) {
						}
						success = false;

					} else if (controller.controllerTransactionType.toString().contains("WebServiceVI")) {
						String actual_XML = responseXml;
						String expected_XML = actual_XML.replaceAll("_Response.xml", "_expected.xml");
						String details = "TransactionID:" + webDriver.getReport().getTestcaseId() + "|" + "CycleDate:"
								+ cycleDate + "|" + "TransactionType:" + controller.controllerTransactionType.toString()
								+ "|" + "ExpectedValue:" + expected_XML + "|" + "ActualValue:" + actual_XML;
						String f_Details = System.getProperty("user.dir") + "\\" + Config.projectName
								+ "\\Resources\\XML_Comparison_File_Helper.txt";
						try (PrintWriter writeDetailsToFile = new PrintWriter(f_Details)) {
							writeDetailsToFile.write(details);
						} catch (Exception e) {
							log.error(e.getMessage(), e);
							log.info("Failed while writing file to " + f_Details);
						}

						try {
							Thread.sleep(5000);
						} catch (InterruptedException e1) {
							log.error(e1.getMessage(), e1);
						}

						Jacob.main(Config.webserviceComparisonUtilityPath, "WebserviceVerificationIntermideate");

						try {
							Thread.currentThread().wait(2000);
							CalendarSnippet.killProcess("EXCEL.EXE");
						} catch (Exception e) {
							log.error(e.getMessage(), e);
						}

						StringBuilder comparedResult = null;
						try (BufferedReader ReaderDetailsFromFile = new BufferedReader(new FileReader(f_Details))) {
							Boolean createStringBuilder = false;
							// String line =
							// ReaderDetailsFromFile.readLine();
							String line = null;
							while ((line = ReaderDetailsFromFile.readLine()) != null) {
								if (createStringBuilder == false) {
									comparedResult = new StringBuilder();
									createStringBuilder = true;
								}
								comparedResult.append(line);
								comparedResult.append(System.lineSeparator());
								log.info(line);
							}

						} catch (Exception e) {
							log.error(e.getMessage(), e);
							log.info("Failed while Reading file from " + f_Details);
						}

						if (comparedResult != null && !comparedResult.toString().equals("")) {
							// If comparedResult is not null and not empty
							webDriver.getReport().setMessage(comparedResult.toString());
							webDriver.getReport().setStatus("FAIL"); // Set status to "FAIL"
						} else {
							// If comparedResult is null or empty
							webDriver.getReport().setStatus("PASS"); // Set status to "PASS"
						}
					} else {
						break;
					}

					pathToNode = null;
					nodeDescription = null;

				}

				break;

			case V:
				currentValue = WebService.getXMLTagValue(controlName);
				break;

			case Write:
				if (ctrlValue != "") {
					WebHelperUtil.writeToExcel(ctrlValue, webElement, controlId, controlType, controlName, rowNo,
							colNo);
				}
				break;

			case Read:

				if (ctrlValue != "") {

					uniqueNumber = WebHelperBilling.ReadFromExcel(ctrlValue);
					String[] tag_Names = controlName.split(";");
					String req_xml = Config.inputDataFilePath + TransactionMapping.directoryPathFileUpload.toString()
							+ "\\WebserviceFiles\\" + cycleDate + "\\" + request_xml + ".xml";
					req_xml = req_xml.replace("/", "_");
					WebService.setReadValueXML(req_xml, tag_Names[0], tag_Names[1], 0, uniqueNumber);

				}
				break;

			default:
				break;
			}
			break;
		case Restful: {

			log.info("---------Rest call reached");
			switch (actionName) {
			case I:
				if (logicalName.equalsIgnoreCase("RestURL")) {
					log.info("rest Url is: " + ctrlValue);
					restUrl = ctrlValue;
				}

				if (logicalName.equalsIgnoreCase("RequestMethod")) {
					requestMethod = ctrlValue;
					log.info("requestMethod is: " + ctrlValue);
				}
				if (logicalName.equalsIgnoreCase("ContentType")) {
					contentType = ctrlValue;
					log.info("contentType is: " + ctrlValue);
				}
				if (logicalName.equalsIgnoreCase("RequestJSON")) {
					requestJson = ctrlValue;
					log.info("requestJson is: " + ctrlValue);
				}
				if (logicalName.equalsIgnoreCase("AuthType")) {
					authType = ctrlValue;
					log.info("AuthType is: " + ctrlValue);
				}
				if (logicalName.equalsIgnoreCase("PathParam")) {
					pathParam = ctrlValue;
					log.info("PathParam is: " + ctrlValue);
				}
				if (logicalName.equalsIgnoreCase("QueryParam")) {
					queryParam = ctrlValue;
					log.info("QueryParam is: " + ctrlValue);
				}
				if (logicalName.equalsIgnoreCase("ReadFromResponse")) {
					readFromResponse = ctrlValue;
					log.info("ReadFromResponse is: " + ctrlValue);
				}
				if (logicalName.equalsIgnoreCase("PathToNode")) {
					pathToNode = ctrlValue;
					log.info("pathToNode is: " + ctrlValue);
				}
				if (logicalName.equalsIgnoreCase("NodeDescription")) {
					nodeDescription = ctrlValue;
					log.info("nodeDescription is: " + ctrlValue);
				}
				if (logicalName.equals("Validate")) {
					if ("Y".equals(ctrlValue))
						validateResponse = true;
					else
						validateResponse = false;
				}

				break;

			case T:
				String xmlResponse = null;

				if (TransactionType.toString().startsWith("RestServiceJSON")) {
					log.info("REST JSON POST request");
					String requestJsonFile = System.getProperty("user.dir") + "\\" + Config.projectName
							+ "\\Resources\\Input\\Restful\\RestfulFiles\\" + cycleDate + "\\" + requestJson + ".json";
					requestJsonFile = requestJsonFile.replace("/", "_");
					String inputJson = new String(Files.readAllBytes(Paths.get(requestJsonFile)));

					Response response = RestService.getRestResponse(contentType, HttpMethod.valueOf(requestMethod),
							restUrl, inputJson, authType, "", "", "", "");
					// restUrl, inputJson, "Auth2.0", "", "", "", "");//Sel4 testing -25\06\24

					restResponse = response.asString();
				}

				if (!StringUtils.isEmpty(restResponse)) {
					log.debug("Writing rest response to a file.");
					String responseFilePath = System.getProperty("user.dir") + "\\" + Config.projectName
							+ "\\Resources\\Input\\Restful\\RestfulFiles\\" + cycleDate + "\\" + requestJson
							+ "_Response.json";
					responseFilePath = responseFilePath.replace("/", "_");
					File responseFile = new File(responseFilePath);
					/* try{ */
					restResponse = CommonUtils.toPrettyFormat(restResponse);// Sel4 - pretty format check?...
					CommonUtils.WriteToFile(responseFile, restResponse.getBytes());

					// Actual response xml
					JSONObject resJson = new JSONObject(restResponse);
					xmlResponse = XML.toString(resJson);
					String soapnode = "<soap:Envelope xmlns:soap= \"http://schemas.xmlsoap.org/soap/envelope/\"> \n"
							+ "<soap:Body> \n"
							+ "<ns2:serviceResponse xmlns:ns2=\"http://com/majescomastek/stgicd/ws/meta/entityinterface\"> \n"
							+ " <return> \n ";
					soapnode = soapnode + xmlResponse;
					soapnode = soapnode + " \n </return> \n </ns2:serviceResponse> \n </soap:Body> \n </soap:Envelope>";

					File xmlFilePath = new File(responseFilePath.replace(".json", ".xml"));
					CommonUtils.WriteToFile(xmlFilePath, soapnode.getBytes());

					// Expected Respons xml
					String expectedJsonFilePath = System.getProperty("user.dir") + "\\" + Config.projectName
							+ "\\Resources\\Input\\Restful\\RestfulFiles\\" + cycleDate + "\\" + requestJson
							+ "_Expected.json";
					expectedJsonFilePath = expectedJsonFilePath.replace("/", "_");

					File expectedJsonFile = new File(expectedJsonFilePath);

					String expectedJson = "";
					if (expectedJsonFile.exists()) {
						expectedJson = new String(Files.readAllBytes(Paths.get(expectedJsonFilePath)));
					}

					if (expectedJson.contains("{")) {
						JSONObject expJson = null;
						try {
							expJson = new JSONObject(expectedJson);
						} catch (Exception e) {
							log.info("error while parsing expected JSON: " + e.getMessage());
						}
						String expectedXML = XML.toString(expJson);

						String soapexpnode = "<soap:Envelope xmlns:soap= \"http://schemas.xmlsoap.org/soap/envelope/\"> \n"
								+ "<soap:Body> \n"
								+ "<ns2:serviceResponse xmlns:ns2=\"http://com/majescomastek/stgicd/ws/meta/entityinterface\"> \n"
								+ " <return> \n ";
						soapexpnode = soapexpnode + expectedXML;
						soapexpnode = soapexpnode
								+ " \n </return> \n </ns2:serviceResponse> \n </soap:Body> \n </soap:Envelope>";

						File expectedFilePath = new File(expectedJsonFilePath.replace(".json", ".xml"));
						CommonUtils.WriteToFile(expectedFilePath, soapexpnode.getBytes());
					} else {
						log.info("Expected Json File does not exist or It is present but its Empty");
					}

				}

				boolean success = false;
				if (!StringUtils.isEmpty(restResponse)) {//Updated on 7/11/24
				    // Check if pathToNode and nodeDescription are both non-null and non-empty
				    if (pathToNode != null && !pathToNode.isEmpty() && nodeDescription != null && !nodeDescription.isEmpty()) {
				    	log.info(" Verifying validateFailedJsonContent");
				        success = JsonUtility.validateFailedJsonContent(restResponse, pathToNode, nodeDescription);
				    } else {
				    	log.info(" Verifying validateRestJson");
				        success = JsonUtility.validateRestJson(restResponse);
				    }
				} else {
				    log.info("Rest response returned is null or empty");
				}

				if (success) {
					webDriver.report.setStatus("PASS");
					webDriver.report.setMessage("SuccessFlag : SUCCESS");
					log.info("Rest response returned");
				} else {
					webDriver.report.setStatus("FAIL");
					if (restResponse == null) {
						webDriver.report.setMessage(
								"Rest Response is null(Internal Server Error:500 OR BAD REQUEST Error:400)");
						log.info("Rest Response is null(Internal Server Error:500 OR BAD REQUEST Error:400)");
					} else {
						if (restErrorResDesc.equals("")) {
							webDriver.report.setMessage("SuccessFlag : FAILED");
							log.info(restErrorResDesc);

						} else {
							webDriver.report.setMessage(restErrorResDesc);
							log.info(restErrorResDesc);
						}
					}
				}
				restErrorResDesc = "";
				pathToNode = null;
				nodeDescription = null;

				break;

			default:
				log.info("Rest action not configured.");
			}
			break;
		}

		case OpenAPI: {
			log.info("---------OpenAPI  service call reached");
			switch (actionName) {
			case I:
				if (logicalName.equalsIgnoreCase("OASURL")) {

					if (Config.openapiHost != null && Config.openapidynamicOASURL != null
							&& Config.openapidynamicOASURL.equalsIgnoreCase("TRUE")) {
						if (!ctrlValue.startsWith("http")) {
							restUrl = Config.openapiHost.concat(ctrlValue.trim());
							log.info("Dynamic OpenAPI URL is" + " " + restUrl);
						} else {
							log.info(
									"Can not create Dynamic OpenAPI URL, where ctrlValue has static URL starting with http... Current ctrlValue is: "
											+ ctrlValue);
							log.info(
									"This transaction WILL NOT BE performed. Remove the ENV details from controlvalue.");
							break;
						}
					} else {
						restUrl = ctrlValue.trim();
						log.info("Static OpenAPI URL is" + " " + restUrl);
					}

				}
				if (logicalName.equalsIgnoreCase("RequestMethod")) {
					requestMethod = ctrlValue;
				}
				if (logicalName.equalsIgnoreCase("ContentType")) {
					contentType = ctrlValue;
				}
				if (logicalName.equalsIgnoreCase("RequestJSON")) {
					requestJson = ctrlValue;
				}
				if (logicalName.equalsIgnoreCase("FolderPath")) {
					folderPath = ctrlValue;
				}
				if (logicalName.equalsIgnoreCase("QueryParam")) {
					queryParam = ctrlValue;
				}
				if (logicalName.equalsIgnoreCase("AuthType")) {
					oasAuthType = ctrlValue;
				}
				if (logicalName.equalsIgnoreCase("PathParam")) {
					pathParam = ctrlValue;
				}
				if (logicalName.equalsIgnoreCase("Headers")) {
					reqHeader = ctrlValue;
				}
				if (logicalName.equalsIgnoreCase("ReadFromResponse")) {
					readFromResponse = ctrlValue;
				}
				if (logicalName.equalsIgnoreCase("UpdateReqBody")) {
					updateReqBody = ctrlValue;
				}
				if (logicalName.equalsIgnoreCase("Multipart")) {
					multiPart = ctrlValue;
				}
				if (logicalName.equalsIgnoreCase("PathToNode")) {
					pathToNode = ctrlValue;
				}
				if (logicalName.equalsIgnoreCase("NodeDescription")) {
					nodeDescription = ctrlValue;
				}

				if (logicalName.equalsIgnoreCase("ValidateResponseIgnoreKeys")) {
					validateResponseIgnoreKeys = ctrlValue;
				}
				if (logicalName.equals("ValidateResponse")) {
					validateResponseFlag = ctrlValue;
					if (validateResponseFlag.startsWith("Y"))
						validateResponse = true;
					else
						validateResponse = false;
				}
				if (logicalName.equalsIgnoreCase("UpdateExpectedResponseBody")) {
					updateExpResponseBody = ctrlValue;
				}

				break;

			case T:
				String oasFolderName, tmpFolderPath;

				if (Config.executionApproach.equalsIgnoreCase("Linear"))
					oasFolderName = testcaseID.toString();
				else
					oasFolderName = cycleDate;

				Boolean pathNOTEmpty = true;
				try {
					if (Objects.isNull(folderPath))
						pathNOTEmpty = false;
					else if (folderPath.isEmpty())
						pathNOTEmpty = false;
				} catch (Exception e) {
					pathNOTEmpty = false;
				}

				// xmlResponse was commented previously, added again
				String xmlResponse = null;
				Response response = null;
				if (TransactionType.toString().startsWith("OASServiceJSON")) {
					log.info("Inside OAS Service, here json request will be selected and processed....");

					String inputJson = "";

					String requestJsonFile = ""; // Initialize the variable at the beginning
					File reqJSONFile = null; // Initialize the File object

					if (Config.openapidynamicFolder != null) {

						if (Config.openapidynamicFolder.equalsIgnoreCase("TRUE") && pathNOTEmpty) {
							log.info("openapidynamicFolder is set to TRUE in the config sheet");
							tmpFolderPath = folderPath.toString();
							requestJsonFile = Config.inputDataFilePath + "OpenAPI\\OpenAPIFiles\\" + tmpFolderPath
									+ "\\" + oasFolderName + "\\" + requestJson + ".json";
							log.info("Dynamic Folder Path for json request is: " + requestJsonFile);
						} else {
							log.info("openapidynamicFolder is set to FALSE in the config sheet");
							requestJsonFile = Config.inputDataFilePath + "OpenAPI\\OpenAPIFiles\\" + oasFolderName
									+ "\\" + requestJson + ".json";
							log.info("Static Folder Path for json request is: " + requestJsonFile);
						}
					} else {
						log.info("openapidynamicFolder parameter is not defined in the config sheet");
						requestJsonFile = Config.inputDataFilePath + "OpenAPI\\OpenAPIFiles\\" + oasFolderName + "\\"
								+ requestJson + ".json";
						log.info("Static Folder Path for json request is: " + requestJsonFile);
					}

					requestJsonFile = requestJsonFile.replace("/", "_");

					// Check if the request JSON file exists
					reqJSONFile = new File(requestJsonFile);

					log.info("Verifying if TransactionType contains POST, GET, PUT or DELETE");

					/*
					 * if (!TransactionType.toString().contains("GET") &&
					 * !TransactionType.toString().contains("_GET") &&
					 * !TransactionType.toString().contains("PUT") &&
					 * !TransactionType.toString().contains("_PUT") &&
					 * !TransactionType.toString().contains("DELETE") &&
					 * !TransactionType.toString().contains("_DELETE")) {
					 */

					String transactionTypeLower = TransactionType.toString().toLowerCase();

					if (!transactionTypeLower.contains("get") && !transactionTypeLower.contains("_get")
							&& !transactionTypeLower.contains("put") && !transactionTypeLower.contains("_put")
							&& !transactionTypeLower.contains("delete") && !transactionTypeLower.contains("_delete")) {

						log.info("TransactionType contains 'POST', checking if request file exists...");

						if (!reqJSONFile.exists()) {
							log.error("Request JSON file does not exist at path: " + requestJsonFile);
							resetVariables();// Sel4
							throw new FileNotFoundException("Request JSON file not found at path: " + requestJsonFile);
						}

						log.info("File exists");
					} else {
						log.info("TransactionType contains either GET,PUT or DELETE : " + TransactionType);
					}

					try {
						if (!updateReqBody.equals("")) {
							log.info("Updating JSON request body...");
							JsonUtility.updateReqJsonFile(requestJsonFile, updateReqBody);
							log.info("JSON request body completed");
						}
					} catch (Exception e) {
						log.error("Error occurred while updating JSON request body: " + e.getMessage());
						resetVariables();// Sel4
						throw new Exception(e);
						// break;
					}

					// update URLs of all OpenAPI requests as per requirement of user.
					if (Config.openapiHost != null && restUrl != null) {
						if (Config.openapiHost.contains(",")) {
							String[] hostArray = Config.openapiHost.split(",");
							restUrl = restUrl.replace(hostArray[0].trim(), hostArray[1].trim());
						} else
							restUrl = restUrl.replace(restUrl.substring(0, StringUtils.ordinalIndexOf(restUrl, "/", 3)),
									Config.openapiHost);
					}

					if (reqJSONFile.exists()) {
						inputJson = new String(Files.readAllBytes(Paths.get(requestJsonFile)));
					}

					log.info("Execute the REST request and get response...");
					response = RestService.getRestResponse(contentType, HttpMethod.valueOf(requestMethod), restUrl,
							inputJson, oasAuthType, queryParam, pathParam, multiPart, reqHeader);
					restResponse = response.prettyPrint();
					multiPart = pathParam = queryParam = oasAuthType = updateReqBody = reqHeader = "";
					log.info("response captured...");
				} else {
					log.info("Transaction name does not starts with OASServiceJSON & hence will not be executed.");
				}

				if (!StringUtils.isEmpty(restResponse)) {
					// log.debug("Writing rest response to a file.");
					log.info("Writing rest response to a file.");
					String responseFilePath = "";

					if (Config.openapidynamicFolder != null) {
						if (Config.openapidynamicFolder.equalsIgnoreCase("TRUE") && pathNOTEmpty) {
							log.info("openapidynamicFolder is set to TRUE in the config sheet");
							tmpFolderPath = folderPath.toString();
							responseFilePath = Config.inputDataFilePath + "OpenAPI\\OpenAPIFiles\\" + tmpFolderPath
									+ "\\" + oasFolderName + "\\" + requestJson + ".json";
							log.info("Dynamic Folder Path to store response json is: " + responseFilePath);
						} else {
							responseFilePath = Config.inputDataFilePath + "OpenAPI\\OpenAPIFiles\\" + oasFolderName
									+ "\\" + requestJson + "_Response.json";
							log.info("Static Folder Path to store response json is: " + responseFilePath);
						}
					} else {
						responseFilePath = Config.inputDataFilePath + "OpenAPI\\OpenAPIFiles\\" + oasFolderName + "\\"
								+ requestJson + "_Response.json";
						log.info("Static Folder Path to store response json is: " + responseFilePath);
					}

					responseFilePath = responseFilePath.replace("/", "_");
					File responseFile = new File(responseFilePath);
					CommonUtils.WriteToFile(responseFile, restResponse.getBytes());

					// Actual response xml
					if (restResponse.startsWith("{")) {
						JSONObject resJson = new JSONObject(restResponse);
						xmlResponse = XML.toString(resJson);
					} else if (restResponse.startsWith("[")) {
						JSONArray resJson = new JSONArray(restResponse);
						xmlResponse = XML.toString(resJson);
					} else {
						log.info("Json Response is Empty or Invalid!!");
					}

					String soapnode = "<soap:Envelope xmlns:soap= \"http://schemas.xmlsoap.org/soap/envelope/\"> \n"
							+ "<soap:Body> \n"
							+ "<ns2:serviceResponse xmlns:ns2=\"http://com/majescomastek/stgicd/ws/meta/entityinterface\"> \n"
							+ " <return> \n ";
					soapnode = soapnode + xmlResponse;
					soapnode = soapnode + " \n </return> \n </ns2:serviceResponse> \n </soap:Body> \n </soap:Envelope>";

					log.info("Converting response json to XML...");
					File xmlFilePath = new File(responseFilePath.replace(".json", ".xml"));
					CommonUtils.WriteToFile(xmlFilePath, soapnode.getBytes());
					log.info("Completed Successfully");

					try {// try and catch added for invalid expected json format - you cannot have
							// duplicate keys in an object
						String expectedJsonFilePath = Config.inputDataFilePath + "OpenAPI\\OpenAPIFiles\\"
								+ oasFolderName + "\\" + requestJson + "_Expected.json";
						expectedJsonFilePath = expectedJsonFilePath.replace("/", "_");

						File expectedJsonFile = new File(expectedJsonFilePath);

						String expectedJson = "";
						if (expectedJsonFile.exists()) {
							expectedJson = new String(Files.readAllBytes(Paths.get(expectedJsonFilePath)));
							String expectedXML = "";
							if (expectedJson.startsWith("{")) {
								JSONObject expJson = new JSONObject(expectedJson);
								expectedXML = XML.toString(expJson);

							} else if (expectedJson.startsWith("[")) {
								JSONArray expJson = new JSONArray(expectedJson);
								expectedXML = XML.toString(expJson);

							} else {
								log.info("Expected XML provided is not valid or its empty!!!");
							}

							String soapexpnode = "<soap:Envelope xmlns:soap= \"http://schemas.xmlsoap.org/soap/envelope/\"> \n"
									+ "<soap:Body> \n"
									+ "<ns2:serviceResponse xmlns:ns2=\"http://com/majescomastek/stgicd/ws/meta/entityinterface\"> \n"
									+ " <return> \n ";
							soapexpnode = soapexpnode + expectedXML;
							soapexpnode = soapexpnode
									+ " \n </return> \n </ns2:serviceResponse> \n </soap:Body> \n </soap:Envelope>";

							log.info("Converting expected json to XML...");
							File expectedFilePath = new File(expectedJsonFilePath.replace(".json", ".xml"));
							CommonUtils.WriteToFile(expectedFilePath, soapexpnode.getBytes());
							log.info("Completed Successfully");
						}
					} catch (JSONException jsonException) {
						// Customized exception for invalid JSON
						log.info("Invalid expected JSON file: " + jsonException.getMessage(), jsonException);
						throw new RuntimeException("Invalid expected JSON file: " + jsonException.getMessage(),
								jsonException);
					} catch (IOException ioException) {
						// Handle IOException
						ioException.printStackTrace(); // Or log the exception
					}
				}
				if (isDMApplication) {
					if (!readFromResponse.equals("")) {
						log.info("readFromResponse: " + readFromResponse);
						JsonUtility.readFromResponseJson(restResponse, readFromResponse);
					}

					if (!updateExpResponseBody.equals("")) {
						String expResJsonFile = Config.inputDataFilePath + "OpenAPI\\OpenAPIFiles\\" + oasFolderName
								+ "\\" + requestJson + "_Expected.json";

						expResJsonFile = expResJsonFile.replace("/", "_");
						JsonUtility.updateReqJsonFile(expResJsonFile, updateExpResponseBody);
					}
				}

				// pre-requisites for json comparison
				String actualResponseFilePath = null, expectedResponseFilePath = null, comparisonResultFilePath = null;
				boolean readyForJsonComparison = false;
				List<String> validateJsonIgnoreKeyList = null;
				if (validateResponse) {
					actualResponseFilePath = (Config.inputDataFilePath + "OpenAPI\\OpenAPIFiles\\" + oasFolderName
							+ "\\" + requestJson + "_Response.json").replace("/", "_");
					expectedResponseFilePath = (Config.inputDataFilePath + "OpenAPI\\OpenAPIFiles\\" + oasFolderName
							+ "\\" + requestJson + "_Expected.json").replace("/", "_");
					comparisonResultFilePath = (Config.inputDataFilePath + "OpenAPI\\OpenAPIFiles\\" + oasFolderName
							+ "\\" + requestJson + "_Result.xlsx").replace("/", "_");

					if (new File(actualResponseFilePath).exists() && new File(expectedResponseFilePath).exists())
						readyForJsonComparison = true;

					if (Config.validateResponseIgnoreKeys != null)// Changed typo 'IngoreKeys' to 'IgnoreKeys'
						validateResponseIgnoreKeys = validateResponseIgnoreKeys + ","
								+ Config.validateResponseIgnoreKeys;

					validateJsonIgnoreKeyList = Arrays.asList(validateResponseIgnoreKeys.trim().split("\\s*,\\s*"));
					validateResponseIgnoreKeys = "";
				}
				// pre-requisites for json comparison END

				// openAPI validation starts [status code validation + Response Json validation]
				boolean openapiSuccess = false; // [Default] based on response status code
				boolean jsonValidationSuccess = false; // based on output of json validation.

				if (!StringUtils.isEmpty(restResponse)) {

					if (!(pathToNode == null || pathToNode.isEmpty())
							&& (!(nodeDescription == null || nodeDescription.isEmpty()))) {
						log.info("pathToNode and nodeDescription present... ");

						openapiSuccess = JsonUtility.validateFailedJsonContent(restResponse, pathToNode,
								nodeDescription);

						if (openapiSuccess) {
							webDriver.report.setStatus("PASS");
							webDriver.report.setMessage(restErrorResDesc);
						} else {
							webDriver.report.setStatus("FAIL");
							webDriver.report.setMessage(restErrorResDesc);
							log.info("FAILED");
						}
						restErrorResDesc = "";
						resetVariables();// Sel4
						break;

					} else {
						// validate request based on its status code
						openapiSuccess = JsonUtility.validateOASJson(response);
					}
					// Setting error message if files are absent for comparison
					if ((validateResponse && readyForJsonComparison == false)) {
						restErrorResDesc = restErrorResDesc
								+ "  => Expected/Response Json file was not present for Json Response Validation.";
						if (validateResponseFlag.equalsIgnoreCase("Y")
								|| validateResponseFlag.equalsIgnoreCase("Y(FAIL)"))
							openapiSuccess = false;
					}

					// for passed and failed requests (Full Json Validation)
					if (validateResponse && readyForJsonComparison
							&& (validateResponseFlag.equalsIgnoreCase("Y(FAIL)") || openapiSuccess)) {

						try {
							// validate response and compare it with expected response json file.
							jsonValidationSuccess = ValidateJson.main(expectedResponseFilePath, actualResponseFilePath,
									validateJsonIgnoreKeyList);

							// Set description message of Response Validation.
							if (jsonValidationSuccess)
								// restErrorResDesc = restErrorResDesc + " => Response Validation Successful" +
								// " => All messages from expected response are present in actual response" + "
								// => " + "List Of Ignored Keys :"+ validateJsonIgnoreKeyList;
								restErrorResDesc = "  => Response Validation Successful"
										+ " => All messages from expected response are present in actual response"
										+ " => " + "List Of Ignored Keys :" + validateJsonIgnoreKeyList;
							else
								// restErrorResDesc = restErrorResDesc + " => Response Validation Failed: " + "
								// => " + "List Of Ignored Keys :"+ validateJsonIgnoreKeyList +
								// comparisonResultFilePath;

								restErrorResDesc = "  => Expected results not matching Actual: " + " => "
										+ "ComparisonResultFilePath :" + " => " + comparisonResultFilePath;
							// END Set description message of Response Validation

							// Deciding Final Status of transaction based on Response Status Code and
							// ValidateResponse Flag.

							if (validateResponseFlag.equalsIgnoreCase("Y"))
								openapiSuccess = openapiSuccess && jsonValidationSuccess; // for [ValidateResponse = Y]
							else if (validateResponseFlag.equalsIgnoreCase("Y(FAIL)"))
								openapiSuccess = jsonValidationSuccess; // for failed requests[ValidateResponse =
							// Y(FAIL)]
							else if (validateResponseFlag.equalsIgnoreCase("Y(IGNORE)")) {

								openapiSuccess = openapiSuccess && jsonValidationSuccess;
								// [Optional] for [ValidateResponse = Y(IGNORE)]

								// openapiSuccess = openapiSuccess; //--- Test/29-02-24 -- need to verify this
							}
						} catch (Exception e) {
							if (validateResponseFlag.equalsIgnoreCase("Y")
									|| validateResponseFlag.equalsIgnoreCase("Y(FAIL)"))
								openapiSuccess = false;
							log.info("Exception occured while comparing JSON: " + e.getMessage());
							restErrorResDesc = restErrorResDesc
									+ "  => Exception occured while validating Response JSON file: " + e.getMessage();
						}
					}
				} else {
					log.info("Rest response returned is null or empty");
					restErrorResDesc = "Rest Response is null(OR Internal Server Error:500)";
				}
				// openAPI validation END

				if (openapiSuccess) {
					webDriver.report.setStatus("PASS");
					webDriver.report.setMessage(restErrorResDesc);

					// for read tag values from response and store it into unique number sheet.
					if (!isDMApplication || ITAFWebDriver.isPASApplication() || ITAFWebDriver.isCSApplication()) {
						if (!readFromResponse.equals("")) {
							JsonUtility.readFromResponseJson(restResponse, readFromResponse);
							readFromResponse = "";
							restResponse = "";
						}
					}
				} else {
					webDriver.report.setStatus("FAIL");
					webDriver.report.setMessage(restErrorResDesc);
					log.info("FAILED");
				}
				restErrorResDesc = "";

				if (isDMApplication) {
					readFromResponse = "";
					restResponse = "";
					updateExpResponseBody = "";
				}
				resetVariables();// Sel4
				break;

			default:

				log.info("Rest action not configured.");
			}
			break;
		}

		case OutPutForm:
			// ---------OutPut Form verification start------------------
			switch (actionName) {
			case I:
				if (logicalName.equalsIgnoreCase("PolicyNo"))
					PolicyNo = ctrlValue;
				else if (logicalName.equalsIgnoreCase("AccountNo"))
					AccountNo = ctrlValue;
				else if (logicalName.equalsIgnoreCase("BrokerNo"))
					BrokerNo = ctrlValue;
				else if (logicalName.equalsIgnoreCase("OutPutFormCode"))
					OutPutFormCode = ctrlValue;
				else if (logicalName.equalsIgnoreCase("OutPutForm_XML")) {
					OutPutForm_XML = ctrlValue;
					String request_xml = System.getProperty("user.dir") + "\\" + Config.projectName
							+ "\\Resources\\Input\\OutPutForm\\OutPutFormXMLFile\\" + cycleDate + "\\" + ctrlValue
							+ ".xml";
					request_xml = request_xml.replace("/", "_");
				}
				break;
			case T:
				String OutputFormName = null;
				ResultSet OutputRecords = null;
				Connection conn = null;
				Statement st = null;

				conn = JDBCConnection.establishHTML5BillingCoreDBConn();
				st = conn.createStatement();
				if (Config.databaseType.equalsIgnoreCase("MsSQL")) {

					OutputRecords = st.executeQuery("SELECT OPFS.OUTPUT_RESPONSE, OPFS.POLICY_TERM_ID FROM "
							+ Config.jbeamdatabaseusername + ".dbo.LOG CORE, " + Config.applicationdatabaseusername
							+ ".dbo.JOB_SCHEDULE BASE , " + Config.applicationdatabaseusername
							+ ".dbo.OUTPUT_FORMS_SCHEDULE OPFS WHERE CORE.BE_SEQ_NO=BASE.JOB_SEQ AND OPFS.JOB_SEQ=BASE.JOB_SEQ AND OPFS.OUTPUT_FORM_CODE = '"
							+ OutPutFormCode
							+ "' AND CORE.TASK_NAME is not null AND CORE.BATCH_NO = (Select max(batch_no)-5 from "
							+ Config.jbeamdatabaseusername
							+ ".dbo.BATCH) AND (BASE.POLICY_NO is not null OR BASE.ACCOUNT_SYSTEM_CODE is not null OR BASE.BROKER_SYSTEM_CODE is not null) AND OPFS.OUTPUT_RESPONSE is not null AND (BASE.POLICY_NO ='"
							+ PolicyNo + "' OR BASE.BROKER_SYSTEM_CODE = (select source_system_entity_code from "
							+ Config.applicationdatabaseusername
							+ ".dbo.entity_register where entity_type= 'BROKER' and SOURCE_SYSTEM_ENTITY_CODE = '"
							+ BrokerNo + "') OR BASE.ACCOUNT_SYSTEM_CODE= (select source_system_entity_code from "
							+ Config.applicationdatabaseusername
							+ ".dbo.entity_register where entity_type= 'ACCOUNT' and SOURCE_SYSTEM_ENTITY_CODE = '"
							+ AccountNo + "'))");

				} else if (Config.databaseType.equalsIgnoreCase("Oracle")) {
					OutputRecords = st.executeQuery("SELECT OPFS.OUTPUT_RESPONSE, OPFS.POLICY_TERM_ID FROM "
							+ Config.jbeamdatabaseusername + ".LOG CORE, " + Config.applicationdatabaseusername
							+ ".JOB_SCHEDULE BASE , " + Config.applicationdatabaseusername
							+ ".OUTPUT_FORMS_SCHEDULE OPFS WHERE CORE.BE_SEQ_NO=BASE.JOB_SEQ AND OPFS.JOB_SEQ=BASE.JOB_SEQ AND OPFS.OUTPUT_FORM_CODE = '"
							+ OutPutFormCode
							+ "' AND CORE.TASK_NAME is not null AND CORE.BATCH_NO = (Select max(batch_no) from "
							+ Config.jbeamdatabaseusername
							+ ".BATCH) AND (BASE.POLICY_NO is not null OR BASE.ACCOUNT_SYSTEM_CODE is not null OR BASE.BROKER_SYSTEM_CODE is not null) AND OPFS.OUTPUT_RESPONSE is not null AND (BASE.POLICY_NO ='"
							+ PolicyNo + "' OR BASE.BROKER_SYSTEM_CODE = (select source_system_entity_code from "
							+ Config.applicationdatabaseusername
							+ ".entity_register where entity_type= 'BROKER' and SOURCE_SYSTEM_ENTITY_CODE = '"
							+ BrokerNo + "') OR BASE.ACCOUNT_SYSTEM_CODE= (select source_system_entity_code from "
							+ Config.applicationdatabaseusername
							+ ".entity_register where entity_type= 'ACCOUNT' and SOURCE_SYSTEM_ENTITY_CODE = '"
							+ AccountNo + "'))");
				} else {
					log.error("Databse is not selected ");
				}

				while (OutputRecords.next()) {
					try {
						OutputFormName = OutputRecords.getString(1).toString();
					} catch (NullPointerException e) {
						log.error(e.getMessage(), e);
						log.info("OutputFormName is null for output Object " + OutPutFormCode + "Batch # " + BatchNo);
					}
				}
				OutputRecords.close();
				st.close();
				JDBCConnection.closeConnection(conn);

				responseXml = request_xml.replace(".xml", "_Response" + ".xml");
				responseXml = responseXml.replace("/", "_");
				File chkDir = new File(System.getProperty("user.dir") + "\\" + Config.projectName
						+ "\\Resources\\Input\\OutPutForm\\OutPutFormXMLFile\\" + wscycledate + "\\");
				if (chkDir.mkdir()) {
					log.info("New directory is created : " + chkDir);
				} else {
					log.info("Directory is already Exists : " + chkDir);
				}
				String actual_XML = responseXml;
				String expected_XML = request_xml;
				File file = new File(responseXml);
				log.info("file object created");

				String content = OutputFormName.toString();
				log.info("content object created");
				FileOutputStream fop = new FileOutputStream(file);
				log.info("FileOutputStream object created");
				if (!file.exists()) {
					file.createNewFile();
					log.info("New response file created");
				}
				byte[] contentInBytes = content.getBytes();
				log.info("contentInBytes done");
				fop.write(contentInBytes);
				log.info("contentInBytes written");
				fop.flush();
				log.info("flushed");
				fop.close();

				String details = "TransactionID:" + webDriver.getReport().getTestcaseId() + "|" + "CycleDate:"
						+ cycleDate + "|" + "TransactionType:" + controller.controllerTransactionType.toString() + "|"
						+ "ExpectedValue:" + expected_XML + "|" + "ActualValue:" + actual_XML;
				String f_Details = System.getProperty("user.dir") + "\\" + Config.projectName
						+ "\\Resources\\XML_Comparison_File_Helper.txt";
				try (PrintWriter writeDetailsToFile = new PrintWriter(f_Details)) {
					writeDetailsToFile.write(details);
				} catch (Exception e) {
					log.error(e.getMessage(), e);
					log.info("Failed while writing file to " + f_Details);
				}

				try {
					Thread.sleep(5000);
				} catch (InterruptedException e) {
					log.error(e.getMessage(), e);
				}

				Jacob.main(Config.webserviceComparisonUtilityPath, "WebserviceVerificationIntermideate");

				try {
					Thread.currentThread().wait(2000);
					CalendarSnippet.killProcess("EXCEL.EXE");
				} catch (Exception e) {
					log.error(e.getMessage(), e);
				}

				StringBuilder comparedResult = new StringBuilder();
				try (BufferedReader ReaderDetailsFromFile = new BufferedReader(new FileReader(f_Details))) {
					String line = null;
					while ((line = ReaderDetailsFromFile.readLine()) != null) {
						comparedResult.append(line);
						comparedResult.append(System.lineSeparator());
						log.info(line);
					}

				} catch (Exception e) {
					log.error(e.getMessage(), e);
					log.info("Failed while Reading file from " + f_Details);
				}

				// Sel4
				if (comparedResult != null && !comparedResult.toString().equals("")) {
					webDriver.getReport().setMessage(comparedResult.toString());
					webDriver.getReport().setStatus("FAIL");
				} else {
					webDriver.getReport().setMessage("");
					webDriver.getReport().setStatus("PASS");
				}

				break;
			// ---------OutPut Form verification end------------------
			default:
				break;
			}
			break;

		case DownloadDocument:

			try {
				// deletes all zip and pdf files from runtime download folder
				String path = Config.runtimeFileDownloadFolder;
				File dir = new File(path);
				@SuppressWarnings("unused")
				List<File> fileList;
				File[] files = dir.listFiles();
				log.info("File download location: " + path);
				for (File file : files) {
					if (file.getName().endsWith("zip") || file.getName().endsWith("pdf")
							|| file.getName().endsWith("xls") || file.getName().endsWith("xlsx")) {
						log.info("Deleting file: " + file.getName());
						file.delete();
					}
				}
			} catch (Exception e) {
				log.info("No file present & hence unable to delete");
			}

			int loop = 5;
			while (loop > 1) {

				try {

					if (webElement.isDisplayed()) {

						webElement = WebHelperPAS.getElementByType(controlId, controlName, WebHelper.control, imageType,
								ctrlValue);
						webElement.click();
						log.info("DownloadDocument clicked");
						log.info("DownloadDocument->clicked");
						break;
					}
				}

				catch (Exception e) {
					log.error(e.getMessage(), e);
					try {

						webElement = WebHelperPAS.getElementByType(controlId, controlName, WebHelper.control, imageType,
								ctrlValue);
						Thread.sleep(2000);
						webElement.click();
						log.info("DownloadDocument catchclicked");
						break;

					} catch (Exception e1) {
						// Below click downloads the document
						Actions MousebuilderClick1 = new Actions(Automation.driver);
						// highlightElement(webElement);
						Action MouseclickAction1 = MousebuilderClick1.moveToElement(webElement).clickAndHold().release()
								.build();
						MouseclickAction1.perform();
						break;
					}
				}
			}
			break;

		case DownloadIEDocument:
			Thread.sleep(20000);
			String DataDir = Config.inputDataFilePath + TransactionMapping.directoryPathFileUpload.toString();
			log.info("On Filedownloader");
			Runtime.getRuntime().exec(DataDir + "\\FileDownload.exe");
			log.info("File download started");
			Thread.sleep(5000);
			break;

		case WaitForElementToVisible:
			if (ctrlValue.equals("") || ctrlValue == null)
				break;

			String temp1 = Automation.driver.findElement(By.xpath(controlName)).getText();
			log.info("Inside WaitForElementToVisible | Current value:-" + temp1 + " | Expected value:-" + ctrlValue);

			int x = 500;
			while (x > 0) {
				if (temp1.equals(ctrlValue)) {
					break;
				} else {
					String temp2 = Automation.driver.findElement(By.xpath(controlName)).getText();
					temp1 = temp2;
					Thread.sleep(2000);
				}
				x--;
			}
			log.info("Out of WaitForElementToVisible | Current value:-" + temp1 + " | Expected value:-" + ctrlValue);

			break;

		case WaitTillFileDownload:
			String[] temp3 = ctrlValue.split(";");
			String dowFilePath = Config.runtimeFileDownloadFolder;
			String DownloadedFile = dowFilePath + "\\" + temp3[0];
			File DownloadFilePath = new File(DownloadedFile);
			for (int z = 1; z <= 120; z++) {
				if (DownloadFilePath.exists()) {
					log.info("File successfully downloaded: " + DownloadFilePath);
					break;
				} else {
					Thread.sleep(1000);
					if (z == 120) {
						log.info("File not found: " + DownloadFilePath);
					}
				}
			}

			break;

		case CopyFromServerLocation:
			String serverLocStr = Config.serverFileLocation;
			File serverLoc = new File(serverLocStr);
			File[] filesLst = serverLoc.listFiles();
			File fileToCopy = null;
			long lastModifiedTime = Long.MIN_VALUE;
			String finalPath = "";

			switch (controlId) {
			case "PDF":
				String[] temp = ctrlValue.split(";");
				finalPath = Config.actualPdfDownloadPath;
				for (File file : filesLst) {
					if (file.getName().endsWith("pdf")) {
						if (file.lastModified() > lastModifiedTime) {
							fileToCopy = file;
							lastModifiedTime = file.lastModified();
						}
					}
				}
				WebHelperUtil.writeToExcel(fileToCopy.getName(), webElement, "writeCtrlValueAsIs", controlType,
						controlName, rowNo, colNo);
				String existingSameFile = finalPath + "\\" + fileToCopy.getName();
				File copiedFile = new File(existingSameFile);
				if (copiedFile.exists()) {
					copiedFile.delete();
					log.info("Old file " + temp[0] + " deleted successfully");
				}
				com.google.common.io.Files.copy(fileToCopy, copiedFile);
				break;
			}
			break;

		case SaveAsDocument:

			try {
				String path = Config.runtimeFileDownloadFolder;
				File dir = new File(path);
				// List<File> fileList;
				File[] files = dir.listFiles();
				for (File file : files) {
					if (file.getName().endsWith("zip") || file.getName().endsWith("pdf")) {
						file.delete();
					}
				}
			} catch (Exception e) {
				log.info("No file present & hence unable to delete");
			}

			Thread.sleep(2000);
			Runtime.getRuntime().exec(Config.externalscriptsfolder + "\\FileDownload.exe "
					+ Config.runtimeFileDownloadFolder + "\\" + testcaseID.toString() + ".pdf");
			Thread.sleep(2000);
			break;

		case MoveDocument:
			// delete file with same name from destination folder
			String[] temp = ctrlValue.split(";");
			if (temp[0].trim().equals("")) {
				String downloadedFileName2 = Config.runtimeFileDownloadFolder;
				File moveFilePath = new File(downloadedFileName2);
				File[] filesList = moveFilePath.listFiles();
				@SuppressWarnings("unused")
				File fileToMove = null;
				@SuppressWarnings("unused")
				String fileName = null;
				for (File file : filesList) {
					log.info("Filename: " + file.getName());
					if (file.getName().endsWith("pdf")) {
						fileToMove = file;
						temp[0] = file.getName();
						WebHelperUtil.writeToExcel(temp[0], webElement, controlId, controlType, controlName, rowNo,
								colNo);
					}
				}
			}
			// Handle failure if file name is not stored in uniqueNumber excel
			if (temp[0].trim().equals("")) {
				throw new Exception("Case MoveDocument: Desired file not present or not downloaded.");
			} else {
				// move downloaded file to final PDF subfolder
				String existingSameFile = Config.actualPdfDownloadPath + "\\" + temp[0];
				File file1 = new File(existingSameFile);

				if (file1.exists()) {
					file1.delete();
					log.info("Old file " + temp[0] + " deleted successfully");
				}

				// move downloaded file to final PDF subfolder
				String downloadedFileName1 = Config.runtimeFileDownloadFolder + "\\" + temp[0];
				File moveFile = new File(downloadedFileName1);

				String distinationFileName = Config.actualPdfDownloadPath + "\\" + temp[0];

				if (moveFile.renameTo(new File(distinationFileName))) {
					// if file copied successfully then delete the original file
					if (moveFile.exists()) {
						moveFile.delete();
					}
					log.info("File " + downloadedFileName1 + " moved successfully to " + distinationFileName);
				} else {
					log.info("Failed to move the file " + downloadedFileName1);
				}
			}

			break;

		case RenameDocument:

			String mainWindow = Automation.driver.getWindowHandle();
			String temp2[] = ctrlValue.split(";");

			if (temp2[0].trim().equals("")) {
				temp2[0] = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
			}

			// Handle failure if file name is not stored in uniqueNumber excel
			if (temp2[0].trim().equals("")) {
				throw new Exception("Case RenameDocument: Desired file not present or not downloaded.");
			} else {
				String downloadedFileName = Config.actualPdfDownloadPath + "\\" + temp2[0];
				String renameFileNameTo = Config.actualPdfDownloadPath + "\\" + temp2[1];
				String ext = FilenameUtils.getExtension(temp2[1]); // returns file ext
				ext = "." + ext;

				File oldFile = new File(downloadedFileName);
				File newFile = new File(renameFileNameTo);

				if (newFile.exists()) {
					newFile.delete();
					log.info("Old file " + temp2[1] + " deleted successfully");
				}

				oldFile.renameTo(newFile);
				String timeStampForFileBackup = new SimpleDateFormat("ddMMMMyyyy_HH_mm_ss").format(new Date());
				String backUpFileName = renameFileNameTo.replace(ext, "") + "_" + timeStampForFileBackup + ext;
				File backUpFile = new File(backUpFileName);
				com.google.common.io.Files.copy(newFile, backUpFile);
			}

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

		case IgnoreString:
			PDFComparisonUtil.setStringToIgnore(ctrlValue);
			break;

		case PDFComparisonMode:
			PDFComparisonUtil.setPDFComparisonMode(ctrlValue);
			break;

		case PDFDocumentCompare:
			if (!ctrlValue.isEmpty()) {
				frmDate = new Date();
				Object[] result = PDFComparisonUtil.PDFCompare(logicalName, controlTypeEnum.toString(), ctrlValue);
				PDFResultBean pdfResultBean = new PDFResultBean();

				if ((ITAFWebDriver.isPASApplication() || ITAFWebDriver.isDMApplication()
						|| ITAFWebDriver.isCSApplication()) && result.length != 0) {

					webDriver.report = WebHelperPAS.WriteToDetailResults(result[1].toString(), result[2].toString(),
							logicalName);
					try {
						pdfResultBean = (PDFResultBean) result[0];
						if (pdfResultBean.isError()) {
							PDFComparisonUtil.writeToPDFComparisonSummary(controller.controllerGroupName.toString(),
									controller.controllerTestCaseID.toString(),
									controller.controllerTransactionType.toString(),
									Config.dtFormat.format(WebHelper.frmDate), result);
						}
					} catch (Exception e) {
						log.error("Error while writing to PDFComparisonSummary " + e.getMessage());
						e.printStackTrace();
					}
				} else if (ITAFWebDriver.isClaimsApplication() || ITAFWebDriver.isBillingApplication()) // Pradeep-Added
				{
					webDriver.setReport(WebHelperUtil.WriteToDetailResults(result[1].toString(), result[2].toString(),
							logicalName));
				}
				// Write code to pass result to claim/billing reports
			}
			break;

		// Controltypes for xml comparison
		case MoveXMLDocument:
			// move downloaded file to final XML subfolder
			String downloadedFileName2 = Config.runtimeFileDownloadFolder;
			File moveFilePath = new File(downloadedFileName2);
			File[] filesList = moveFilePath.listFiles();
			File fileToMove = null;
			String fileName = null;
			for (File file : filesList) {
				if (file.getName().endsWith("zip")) {
					fileToMove = file;
					fileName = file.getName();
				}
			}
			// delete if file exists
			if (fileToMove.exists()) {
				String distinationFileName1 = Config.actualXMLDownloadPath + "\\" + fileName;
				// String downloadedXMLFilename = fileName;
				File downloadedFile = new File(distinationFileName1);

				if (downloadedFile.exists()) {
					downloadedFile.delete();
					log.info("Old file " + fileName + " deleted successfully");
				}

				if (fileToMove.renameTo(new File(distinationFileName1))) {
					if (fileToMove.exists()) {
						fileToMove.delete();
					}
					log.info("File " + fileToMove.getAbsolutePath() + " moved successfully to " + distinationFileName1);
				} else {
					log.info("Failed to move the file " + fileToMove.getAbsolutePath());
				}
			}
			break;

		case MoveXMLDocumentExpected:
			// move downloaded file to final XML subfolder
			String downloadedfileName13 = Config.runtimeFileDownloadFolder;
			File moveFilePath1 = new File(downloadedfileName13);
			File[] filesList1 = moveFilePath1.listFiles();
			File fileToMove1 = null;
			String fileName2 = null;
			String NewFile = ctrlValue;
			for (File file : filesList1) {
				if (file.getName().endsWith("xml") || file.getName().endsWith("zip")) { /// Handled for files other than
					/// .zip
					fileToMove1 = file;
					fileName2 = file.getName();
				}
			}
			// delete if file exists
			if (fileToMove1.exists()) {
				String distinationfileName11 = Config.xmldownloads + "\\" + NewFile;
				// String downloadedXMLfileName1 = fileName1;
				File RenameFile = new File(distinationfileName11);

				if (RenameFile.exists()) {
					RenameFile.delete();
					log.info("Old file " + fileName2 + " deleted successfully");
				}

				if (fileToMove1.renameTo(new File(distinationfileName11))) {
					if (fileToMove1.exists()) {
						fileToMove1.delete();
					}
					log.info("File " + fileToMove1.getAbsolutePath() + " moved successfully to "
							+ distinationfileName11);
				} else {
					log.info("Failed to move the file " + fileToMove1.getAbsolutePath());
				}
			}
			break;

		case UnzipFolderAndRename:

			String mainWindow1 = Automation.driver.getWindowHandle();
			String newFileName = ctrlValue;

			String renameXMLFileNameTo = Config.actualXMLDownloadPath + "\\" + newFileName;

			File newXMLFile = new File(renameXMLFileNameTo);
			// log.info(oldXMLFile.getParentFile());

			if (newXMLFile.exists()) {
				newXMLFile.delete();
				log.info("Old file " + newFileName + " deleted successfully");
			}

			// Unzip zipped file
			String downloadedFileName4 = Config.actualXMLDownloadPath;
			XmlComparisonUtil.unZipXmlFolder(downloadedFileName4);

			// get the xml path from unzipped folder
			File dirExport = new File(downloadedFileName4);
			File[] exportFiles = dirExport.listFiles();

			for (File exportFile : exportFiles) {
				if (!exportFile.getName().endsWith("zip")) {
					if (exportFile.isDirectory()) {
						XmlComparisonUtil.setActualXmlPath(downloadedFileName4 + "\\" + exportFile.getName());
					}
				}
			}

			// move the xml from unzipped folder to actual download path
			String downloadedXMLFileName = XmlComparisonUtil.ActualXmlPath;
			File oldXMLFile = new File(downloadedXMLFileName);

			log.info("Unzip Folder " + oldXMLFile.renameTo(newXMLFile));

			String timeStampForXMLFileBackup = new SimpleDateFormat("ddMMMMyyyy_HH_mm_ss").format(new Date());
			String backUpXMLFileName = renameXMLFileNameTo.replace(".xml", "") + "_" + timeStampForXMLFileBackup
					+ ".xml";

			File backUpXMLFile = new File(backUpXMLFileName);

			com.google.common.io.Files.copy(newXMLFile, backUpXMLFile);

			// delete the zipped and unzipped folder
			String downloadedFileName3 = Config.actualXMLDownloadPath;
			File file2 = new File(downloadedFileName3 + "\\");
			File[] filesDeleteList = file2.listFiles();
			for (File file : filesDeleteList) {
				if (file.isDirectory()) {
					FileUtils.deleteDirectory(file);
				} else if (file.getName().endsWith("zip")) {
					file.delete();
				}
			}

			Set<String> handlers2 = null;
			handlers2 = Automation.driver.getWindowHandles();
			for (String handler3 : handlers2) {
				if (!mainWindow1.equalsIgnoreCase(handler3)) {
					Automation.driver.switchTo().window(handler3);
					Automation.driver.close();
				}
			}
			Automation.driver.switchTo().window(mainWindow1);

			break;

		case RenameXML: // Added for renaming the downloaded xml's

			@SuppressWarnings("unused")
			String mainWindow2 = Automation.driver.getWindowHandle();
			String newFileName1 = ctrlValue;

			String renameXMLFileNameTo1 = Config.actualXMLDownloadPath + "\\" + newFileName1;

			File newXMLFile1 = new File(renameXMLFileNameTo1);

			if (newXMLFile1.exists()) {
				newXMLFile1.delete();
				log.info("Old file " + newFileName1 + " deleted successfully");
			}

			String downloadedFileName5 = Config.actualXMLDownloadPath;
			File dirExport1 = new File(downloadedFileName5);
			File[] exportFiles1 = dirExport1.listFiles();
			long lastMod1 = Long.MIN_VALUE;
			File choice = null;
			for (File exportFile : exportFiles1) {
				if (exportFile.getName().endsWith("xml")) {
					if (exportFile.lastModified() > lastMod1) {
						choice = exportFile;
						lastMod1 = exportFile.lastModified();
					}
					XmlComparisonUtil.ActualXmlPath = choice.getAbsolutePath();
				}
			}
			String downloadedXMLFileName2 = XmlComparisonUtil.ActualXmlPath;
			File oldXMLFile1 = new File(downloadedXMLFileName2);

			log.info("Unzip Folder " + oldXMLFile1.renameTo(newXMLFile1));

			String timeStampForXMLFileBackup1 = new SimpleDateFormat("ddMMMMyyyy_HH_mm_ss").format(new Date());
			String backUpXMLFileName1 = renameXMLFileNameTo1.replace(".xml", "") + "_" + timeStampForXMLFileBackup1
					+ ".xml";

			File backUpXMLFile1 = new File(backUpXMLFileName1);

			com.google.common.io.Files.copy(newXMLFile1, backUpXMLFile1);

			break;

		case XMLIgnoreTags_AttributeLevel:

			XmlComparisonUtil.setAttributeLevelTagsToIgnore(ctrlValue);

			break;

		case XMLIgnoreTags_NodeLevel:

			XmlComparisonUtil.setNodeLevelTagsToIgnore(ctrlValue);

			break;

		case XMLCompare:

			if (!ctrlValue.isEmpty()) {
				String[] result = XmlComparisonUtil.XMLCompare(logicalName, ctrlValue);
				if (ITAFWebDriver.isPASApplication() && result.length != 0) {
					webDriver.report = WebHelperPAS.WriteToDetailResults(result[1], result[2], logicalName);
					// Write code to pass result to claim/billing reports
				} else if (ITAFWebDriver.isClaimsApplication() || ITAFWebDriver.isBillingApplication()) // Added for
				// Biling to
				// update the
				// DetailsResults.csv
				{
					webDriver.setReport(WebHelperUtil.WriteToDetailResults(result[1].toString(), result[2].toString(),
							logicalName));
				}
			}
			break;

		case RenameFile:
			if (!ctrlValue.isEmpty()) {
				String[] tempVar = ctrlValue.split(";");
				String fileType = tempVar[0];
				String folderPath = tempVar[1];
				String renameFromFileName = tempVar[2];
				String renameToFileName = tempVar[3];

				File fileFrom = null;
				File fileTo = null;

				if (folderPath.equals("")) {
					folderPath = Config.runtimeFileDownloadFolder;
				}
				if (renameFromFileName.equals("")) {
					File filePath = new File(folderPath);
					File[] listOfFiles = filePath.listFiles();
					String fileName1 = null;
					long lastMod = Long.MIN_VALUE;

					for (File file : listOfFiles) {
						if (file.lastModified() > lastMod && file.getName().endsWith(fileType)) {
							fileFrom = file;
							lastMod = file.lastModified();
							fileName1 = file.getName();
						}
					}
					renameFromFileName = fileName1;
				} else {
					fileFrom = new File(folderPath + "\\" + renameFromFileName);
				}

				String fileToStr = folderPath + "\\" + renameToFileName;
				fileTo = new File(fileToStr);

				if (fileTo.exists()) {
					fileTo.delete();
					log.info("Old file " + fileTo + " deleted successfully");
				}

				fileFrom.renameTo(fileTo);
				log.info("File '" + fileFrom.getAbsolutePath() + "' successfully renamed to " + "'"
						+ fileTo.getAbsolutePath() + "'.");

			}

			break;

		case MoveFile:

			if (!ctrlValue.isEmpty()) {
				String[] tempVar = ctrlValue.split(";");
				String fileNameToMove = tempVar[0];
				String moveFromFolderPath = tempVar[1];
				String moveToFolderPath = tempVar[2];
				if (moveFromFolderPath.equals("")) {
					moveFromFolderPath = Config.runtimeFileDownloadFolder;
				}

				String checkSameFileInDest = moveToFolderPath + "\\" + fileNameToMove;
				File fileCheck = new File(checkSameFileInDest);

				if (fileCheck.exists()) {
					fileCheck.delete();
					log.info("Old file " + fileCheck + " deleted successfully");
				}

				// move file
				String renamedFileName1 = moveFromFolderPath + "\\" + fileNameToMove;
				File moveFile1 = new File(renamedFileName1);

				String distinationFileName1 = moveToFolderPath + "\\" + fileNameToMove;

				if (moveFile1.renameTo(new File(distinationFileName1))) {
					if (moveFile1.exists()) {
						moveFile1.delete();
					}
					log.info("File " + renamedFileName1 + " moved successfully to " + distinationFileName1);
				} else {
					log.info("Failed to move the file " + renamedFileName1);
				}
			}

			break;

		case TPA_Archive:
			switch (actionName)

			{
			case I:
				try {
					String TPANo = null;
					if (logicalName.trim().equalsIgnoreCase("TPANo")) {
						TPANo = ctrlValue;

						String query = "UPDATE entity_register SET"
								+ " source_system_entity_code=concat(source_system_entity_code, 1)"
								+ "WHERE source_system_entity_code='" + TPANo + "'";

						Connection conn = JDBCConnection.establishHTML5BillingDBConn();
						Statement st = conn.createStatement();
						st.execute(query);
						if (Config.databaseType.equalsIgnoreCase("ORACLE")) {
							st.execute("commit");
						}
						st.close();
						JDBCConnection.closeConnection(conn);

						// Updating Report
						webDriver.getReport().setMessage("TPA Archived : " + TPANo);

					}
				} catch (Exception e) {
					log.info(e.getLocalizedMessage());
					webDriver.getReport().setStatus("FAIL");
					webDriver.getReport().setMessage(e.getLocalizedMessage());
				}

				break;
			default:
				break;
			}
			break;
			
		/*
		 * case ClearCache:
		 * 
		 * switch(action) { case "ClearCache": // Clear browser cache by deleting all
		 * cookies currentdriver.manage().deleteAllCookies(); Thread.sleep(1000);
		 * System.out.println("Browser cache cleared"); break;
		 * 
		 * // You can add more cases here case "RefreshPage": // Refresh the browser
		 * page currentdriver.navigate().refresh();
		 * System.out.println("Page refreshed"); break;
		 * 
		 * default: System.out.println("No matching action found"); break; }
		 */

		case DeriveBusinessDate:

			if (ctrlValue.equalsIgnoreCase("pickFromUniqueNumbers")) {
				ctrlValue = WebHelperUtil.ReadFromExcel(ctrlValue, WebHelper.columnName);
			}

			DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/uuuu");
			LocalDate d = LocalDate.parse(ctrlValue, formatter);
			LocalDate businessDate;
			if (d.getDayOfWeek().toString().equalsIgnoreCase("monday")) {
				businessDate = d.minusDays(3);
			} else if (d.getDayOfWeek().toString().equalsIgnoreCase("sunday")) {
				businessDate = d.minusDays(2);
			} else {
				businessDate = d.minusDays(1);
			}
			currentValue = businessDate.format(formatter);
			break;

		default:
			log.info("U r in Default");
			break;
		}

		return currentValue;

	}

	private static void resetVariables() {
		// Resetting or releasing variables
		requestMethod = "";
		contentType = "";
		requestJson = "";
		folderPath = "";
		queryParam = "";
		oasAuthType = "";
		pathParam = "";
		reqHeader = "";
		readFromResponse = "";
		updateReqBody = "";
		multiPart = "";
		pathToNode = "";
		nodeDescription = "";
		validateResponseIgnoreKeys = "";
		validateResponseFlag = "";
		validateResponse = false; // Resetting boolean value to false
		updateExpResponseBody = "";

	}

	@SuppressWarnings("unused")
	private static void print(String variable1) {
		// TODO Auto-generated method stub

	}

	@SuppressWarnings("unused")
	private static void SwitchtoChildWindow() throws InterruptedException {
		String currWin = currentdriver.getWindowHandle();
		Set<String> allWindows = currentdriver.getWindowHandles();
		Thread.sleep(2000);
		List<String> winsList = new ArrayList<>(allWindows);
		log.info("Total Window Count: " + winsList.size());
		String window1 = winsList.get(0);
		String window2 = winsList.get(1);
		if (currWin.equals(window1))
			Automation.driver.switchTo().window(window2);
		else
			Automation.driver.switchTo().window(window1);

		/*
		 * String parentWindow = currentdriver.getWindowHandle();
		 * log.info("Title of the parentWindow: " +currentdriver.getTitle());
		 * log.info("Parent window handle: " + parentWindow); Set<String> allWindows =
		 * currentdriver.getWindowHandles(); log.info("No of Windows =" +
		 * allWindows.size()); String anotherWidow = null; for (String curWindow :
		 * allWindows) { if (!curWindow.equals(parentWindow)) {
		 * Automation.driver.switchTo().window(curWindow); anotherWidow = curWindow;
		 * Thread.sleep(1000); log.info("Title of the new window: "
		 * +currentdriver.getTitle()); } }
		 */
	}
	
	// Method to clear cache
    public static void clearCache(WebDriver driver) {
    	try {
            // Retrieve all cookies
            Set<Cookie> cookies = driver.manage().getCookies();

            // Print details of each cookie
            if (cookies.isEmpty()) {
                log.info("No cookies found.");
            } else {
                System.out.println("Cookies to be deleted:");
                for (Cookie cookie : cookies) {
                    log.info("Name: " + cookie.getName());
                    log.info("Domain: " + cookie.getDomain());
                    log.info("Path: " + cookie.getPath());
                    log.info("Expiry: " + cookie.getExpiry());
                    log.info("Is Secure: " + cookie.isSecure());
                    log.info("-----------------------------");
                }
            }

            // Clear the browser cookies
            driver.manage().deleteAllCookies();
            Thread.sleep(1000); // Optional sleep for waiting
            log.info("Browser cache cleared");

        } catch (InterruptedException e) {
            e.printStackTrace(); // Handle the InterruptedException
        } catch (Exception e) {
            log.error("An error occurred while clearing the cache: " + e.getMessage());
        }
    }

    // Additional methods can be added here (e.g., refreshPage, etc.)

    
}

    
    