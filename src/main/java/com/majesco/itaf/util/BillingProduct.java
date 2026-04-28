package com.majesco.itaf.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.interactions.Action;
import com.majesco.itaf.main.Automation;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.ITAFWebDriver;
import com.majesco.itaf.main.MainController;
import com.majesco.itaf.main.ObjectFactory;
import com.majesco.itaf.main.WebHelper;
import com.majesco.itaf.main.WebHelperUtil;
import com.majesco.itaf.recovery.StartRecovery;

public class BillingProduct {

	private final static Logger log = LogManager.getLogger(BillingProduct.class.getName());
	public static LinkedHashMap<String, Object> ColumnheaderIndex = new LinkedHashMap<String, Object>();
	public static LinkedHashMap<String, Object> Operatectrlvalues = new LinkedHashMap<String, Object>();
	public static LinkedHashMap<String, Object> Searchctrlvalues = new LinkedHashMap<String, Object>();
	public static LinkedHashMap<String, Object> OperateControlType = new LinkedHashMap<String, Object>();
	public static DataFormat TIformat = null;
	private static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	private static MainController controller = ObjectFactory.getMainController();
	public static File file;

	public static void findAction(WebElement tableFound, String controlName, String logicalName, Row rowValues,
			HashMap<String, Integer> valuesHeader) throws Exception {
		try {
			ColumnheaderIndex.clear();
			Operatectrlvalues.clear();
			Searchctrlvalues.clear();
			OperateControlType.clear();
			//log.info("Find line 59");

			String Attempt = "";
			//System.out.println("logicalName is : " + logicalName);
			log.info("logicalName is : " + logicalName);
			String currentHeader = null;
			String finalXpathTemp;
			int tempPosition;
			List<WebElement> table_Th = tableFound.findElements(By.tagName("th"));
			WebElement firstelement = table_Th.get(0);
			((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView(true);", firstelement);

			for (int i = 0; i < table_Th.size(); i++) {
				currentHeader = (table_Th.get(i).getText()).toUpperCase();

				if (ColumnheaderIndex.containsKey(currentHeader) == false) {
					ColumnheaderIndex.put(currentHeader, i + 1);
				}
			}
			log.info("Application column headers and their indexes are: " + ColumnheaderIndex);

			String[] ExcelLogicalName = logicalName.split("\\|");
			String[] excelHeader = null;
			String[] actualHeader = null;
			String Sctrl = "";
			excelHeader = ExcelLogicalName[0].split("\\:");
			actualHeader = ExcelLogicalName[1].split("\\:");

			int arrlen = ExcelLogicalName.length;
			if (((arrlen - 2) % 3) == 0) {
			} else {
				String AttemptColName = ExcelLogicalName[arrlen - 1];
				Cell ctrlValuecell111 = rowValues
						.getCell(Integer.parseInt(valuesHeader.get(AttemptColName).toString()));
			

				if (ctrlValuecell111 != null) {
					try {
						FormulaEvaluator evaluator = rowValues.getSheet().getWorkbook().getCreationHelper()
								.createFormulaEvaluator();
						DataFormatter fmt = new DataFormatter();
						// CellValue cellValue = evaluator.evaluate(ctrlValuecell111);
						Attempt = fmt.formatCellValue(ctrlValuecell111, evaluator);
						log.info("Attempt value: " + Attempt);
					} catch (Exception e) {
						// Handle any exceptions here
						//e.printStackTrace();
						log.error("Error occurred during formula evaluation: " + e.getMessage());
					}
				} else {
					Attempt = "";
				}

				System.out.println(" Attempt no " + Attempt);
			}

			Workbook workbook = rowValues.getSheet().getWorkbook();
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

			int actualHeaderLen = excelHeader.length;
			String ctrlValue = null;
			for (int i = 0; i < actualHeaderLen; i++) {
				Cell ctrlValuecell = rowValues.getCell(valuesHeader.get(excelHeader[i]));

				if (ctrlValuecell != null) {
					try {
						DataFormatter fmt = new DataFormatter();
						// CellValue cellValue = evaluator.evaluate(ctrlValuecell);
						ctrlValue = fmt.formatCellValue(ctrlValuecell, evaluator);
						Searchctrlvalues.put(actualHeader[i], ctrlValue);
						Sctrl = Sctrl + ctrlValue;// Only concatenate if ctrlValue is not null
					} catch (Exception e) {
						// Handle any exceptions here
						e.printStackTrace();
						log.error("Error occurred during formula evaluation: " + e.getMessage());
					}
				} else {
			        // Handle null value here, maybe set it to an empty string
			        Searchctrlvalues.put(actualHeader[i], ""); // or Searchctrlvalues.put(actualHeader[i], null);
			    }

				//Searchctrlvalues.put(actualHeader[i], ctrlValue);
				//Sctrl = Sctrl + ctrlValue;
			}

			if (Sctrl.equals("")) {
				log.info(
						"As per value sheet, there is nothing to be searched on application and hence coming out of FIND function");
				return;
			}
			// log.info("Concatenated string in Values : " + Sctrl);
			log.info("Values to be searched on application table are : " + Searchctrlvalues);

			int logicalnameLen = ExcelLogicalName.length;
			for (int i = 2; i < logicalnameLen; i += 3) // for getting target
			// columns and values
			{
				Cell ctrlValuecell = rowValues.getCell(valuesHeader.get(ExcelLogicalName[i + 2]));
				DataFormatter fmt = new DataFormatter();
				
				
				if (ctrlValuecell == null || fmt.formatCellValue(ctrlValuecell).isEmpty()) {
					ctrlValue = "";
				}

				else  {//if (ctrlValuecell != null)
					try {
						CellValue cellValue = evaluator.evaluate(ctrlValuecell);

						if (cellValue != null) {

							ctrlValue = fmt.formatCellValue(ctrlValuecell, evaluator);
						}else {
							ctrlValue = "";
						}
					} catch (Exception e) {
						// Handle any exceptions here
						e.printStackTrace();
						log.error("Error occurred during formula evaluation: " + e.getMessage());
					}

				}

				Operatectrlvalues.put(ExcelLogicalName[i], ctrlValue);
				OperateControlType.put(ExcelLogicalName[i], ExcelLogicalName[i + 1]);

			}
			log.info("Target column ctrlvalues are : " + Operatectrlvalues);
			log.info("Target column ctrltypes are: " + OperateControlType);

			// for creating XPath
			String finalXpath = controlName + "//tr";
			Set<String> keys = Searchctrlvalues.keySet();
			int counter = 1;
			for (String key : keys) // For loop
			{
				int positionint;
				String position;

				try {
					positionint = (Integer.parseInt(key));
					position = "" + positionint;
				} catch (Exception pe) {
					position = (String) ColumnheaderIndex.get((key).toUpperCase()).toString().trim();
				}
				
				Object value = Searchctrlvalues.get(key);
				//String ctrlvalue = (value != null) ? value.toString() : "";
				String ctrlvalue;
				if (value != null) {
				    ctrlvalue = value.toString();
				} else {
				    ctrlvalue = ""; // Assigning empty string if value is null
				}
				
				//String ctrlvalue = (String) Searchctrlvalues.get(key).toString();
				//String ctrlvalue = (String) Searchctrlvalues.get(key);
				
				log.info("Search key , position , value is : " + key + ": " + position + ": " + ctrlvalue);
				String datePattern = "\\d{4}-\\d{2}-\\d{2}";
				if (!ctrlvalue.equalsIgnoreCase("")) {
					if (ctrlvalue.matches(datePattern)) {
						if (counter == 1) {
							finalXpath = finalXpath + "/td[position()=" + position + " and (.//@title='"
									+ (ctrlvalue.trim()) + "' or .//text()='" + (ctrlvalue.trim()) + "')]";
						} else {
							finalXpath = finalXpath + "/../td[position()=" + position + " and (.//@title='"
									+ (ctrlvalue.trim()) + "' or .//text()='" + (ctrlvalue.trim()) + "')]";
						}
					} else if (!ctrlvalue.equalsIgnoreCase("BLANK")) {
						if (counter == 1) {
							finalXpath = finalXpath + "/td[position()=" + position + " and (.//@value='"
									+ (ctrlvalue.trim()) + "' or .//text()='" + (ctrlvalue.trim()) + "')]";
						}

						else {
							finalXpath = finalXpath + "/../td[position()=" + position + " and (.//@value='"
									+ (ctrlvalue.trim()) + "' or .//text()='" + (ctrlvalue.trim()) + "')]";
						}
					} else {
						if (counter == 1) {
							finalXpath = finalXpath + "/td[position()=" + position
									+ " and (.//@value='N' or  @title='')]";
						} else {
							finalXpath = finalXpath + "/../td[position()=" + position
									+ " and (.//@value='N' or  @title='')]";
						}
					}
					counter++;
				}
			}

			log.info("finalXpath is : " + finalXpath);
			finalXpathTemp = finalXpath;
			Set<String> Operatekeys = Operatectrlvalues.keySet();
			String targetXpath = "";

			for (String Oprkey : Operatekeys) {
				String controlType = (String) OperateControlType.get(Oprkey).toString();
				String controlValue = (String) Operatectrlvalues.get(Oprkey).toString();

				int positionint;
				String position;

				try {
					positionint = (Integer.parseInt(Oprkey));
					position = "" + positionint;
				} catch (Exception pe) {
					position = (String) ColumnheaderIndex.get((Oprkey).toUpperCase()).toString().trim();

				}
				tempPosition = (Integer.parseInt(position)) + 1;
				log.info("Target key , position , value , type : " + Oprkey + ": " + position + ": " + controlValue
						+ ": " + controlType);

				targetXpath = finalXpath + "/../td[position()=" + position + "]";

				if (controlType.equalsIgnoreCase("WebEdit")) {
					String XPath = targetXpath + "//div/div/input";
					log.info("XPath of target WebEdit is : " + XPath);
					WebElement newelement = Automation.driver.findElement(By.xpath(XPath));
					if (Automation.browserType.toString().toUpperCase().contains("CHROME")
							|| Automation.browserType.toString().toUpperCase().contains("MSEDGE")) {
						log.info("inside chrome/MsEdge browser code");

						newelement = WebHelper.currentdriver.findElement(By.xpath(XPath));
						Actions xAct = new Actions(WebHelper.currentdriver);
						xAct.moveToElement(newelement).click().perform();
						log.info("newelement got:" + newelement);
					}

					if (controlValue.equalsIgnoreCase("BLANK")) {
						// newelement.click();
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								newelement);
						Thread.sleep(100);
						newelement.clear();
					}

					else if (controlType != "" && controlValue != "") {

						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								newelement);
						Thread.sleep(1000);
						newelement.clear();

						Thread.sleep(200);
						if (newelement.isDisplayed()) {
							try {
								WebElement OKwebButton = Automation.driver.findElement(By.xpath(".//*[@name='Ok']"));
								((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
										OKwebButton);
								log.info("OK Button clicked");
							} catch (Exception e) {
								// log.error(e.getMessage(), e);
								log.info(
										"Error msg with 'OK' button not displayed as in the case of WriteOffUndoWriteOff trnx");

							}
						}
						Thread.sleep(500);
						// newelement.sendKeys(controlValue);
						newelement.sendKeys(Keys.chord(Keys.CONTROL, "a"), controlValue);
						// newelement.sendKeys(Keys.chord(Keys.CONTROL,"a", Keys.DELETE)
						Thread.sleep(100);
						newelement.sendKeys(Keys.TAB);
						Thread.sleep(100);
					}

				} else if (controlType.equalsIgnoreCase("WebButton")) {
					if (controlValue != "") {
						String XPath = targetXpath + "//div/div/button";
						log.info("XPath of target WebButton is : " + XPath);
						WebElement newelement = Automation.driver.findElement(By.xpath(XPath));
						// Thread.sleep(100);
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								newelement);
						Thread.sleep(100);
						// newelement.click();
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();", newelement);// Testing
					}
				} else if (controlType.equalsIgnoreCase("WebLink")) {
					if (controlValue != "") {

						if (controller.controllerTransactionType.toString().equalsIgnoreCase("ModifyAccountCurrent")
								&& (tempPosition == 12)) {
							WebElement headers = Automation.driver.findElement(
									By.xpath(".//*[@id='dgAddTransactionTemp']//*[@title='Transaction']/div"));
							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();", headers);
							Thread.sleep(100);

							Boolean hclicked = false;

							while (!hclicked) {
								try {

									WebElement headers_asc = Automation.driver.findElement(By.xpath(
											".//*[contains(@id,'dgAddTransactionTemp')]/tbody/tr[1]/td[3]/div/div/span/span/span/span"));
									String first_value = headers_asc.getAttribute("title");
									log.info(first_value);
									if (first_value.equals("New Business")) {
										hclicked = true;
									} else {
										((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
												headers);
										first_value = "";
									}

								}

								catch (Exception ex) {
									log.error(ex.getMessage(), ex);
									log.info("Header element not found");
								}
							}
							Thread.sleep(15000);
						}
						WebElement newelement;
						if (Automation.browserType.toString().toUpperCase().contains("CHROME")
								|| Automation.browserType.toString().toUpperCase().contains("MSEDGE")) {
							String linkXPath = targetXpath + "//div//span/span";

							try {
								newelement = Automation.driver.findElement(By.xpath(linkXPath));
								log.info("XPath of target WebLink is : " + linkXPath);
							} catch (Exception e) {
								linkXPath = targetXpath;
								newelement = Automation.driver.findElement(By.xpath(linkXPath));
								log.info("XPath of target WebLink is : " + linkXPath);
							}
						} else {
							String linkXPath = targetXpath + "//div/div";
							log.info("XPath of target WebLink is : " + linkXPath);
							newelement = Automation.driver.findElement(By.xpath(linkXPath));
						}

						if (controller.controllerTransactionType.toString().equalsIgnoreCase("ModifyAccountCurrent")) {
							log.info("Do not scroll");
						}

						else if (controller.controllerTransactionType.toString().equalsIgnoreCase("UndoWriteOff")
								|| controller.controllerTransactionType.toString()
										.equalsIgnoreCase("CreateNotification")) {
							log.info("Do not scroll");

						} else {
							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
									newelement);

						}
						Thread.sleep(100);
						if (Automation.browserType.toString().toUpperCase().contains("CHROME")
								|| Automation.browserType.toString().toUpperCase().contains("MSEDGE")) {
							Actions actions = new Actions(Automation.driver);
							actions.moveToElement(newelement).click().perform();
							Thread.sleep(3000);
							log.info("First Click Applied");

							try {
								if (newelement.isDisplayed()) {
									Actions actions2 = new Actions(Automation.driver);
									actions2.moveToElement(newelement).click().perform();
									// newelement.click();
									Thread.sleep(4000);
									log.info("Second Click Applied");

								}
							} catch (Exception e) {
								log.info("Element is no longer displayed");
							}

							Thread.sleep(1000);
						} else {
							newelement.click();
							Thread.sleep(2000);
							log.info("First Click Applied");

							try {
								if (newelement.isDisplayed()) {
									newelement.click();
									Thread.sleep(2000);
									log.info("Second Click Applied");

								}
							} catch (Exception e) {
								log.info("Element is no longer displayed");
							}
						}

						Thread.sleep(100);
						log.info("Checks");

						// ******For Modify Acc Current ***
						int loopcntr = 0;
						if (controller.controllerTransactionType.toString().equalsIgnoreCase("ModifyAccountCurrent")) {
							String tempValue;

							if (tempPosition == 12) {
								finalXpathTemp = finalXpathTemp + "/../td[position()=" + (tempPosition) + "]";
								WebElement tempEdit = Automation.driver.findElement(By.xpath(finalXpathTemp));

								tempValue = tempEdit.getAttribute("title");

								// while (tempValue.isEmpty())
								while (tempValue.isEmpty() && loopcntr < 50) // meghna--Loop
								// Control
								{
									if (newelement.isDisplayed()) {
										// tempEdit.sendKeys(Keys.UP);
										// newelement.click();
										// to test with javascript
										((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
												newelement);

									}
									tempValue = tempEdit.getAttribute("title");
									loopcntr = loopcntr + 1;
								}
							}

						}

					}

					if (controller.controllerTransactionType.toString().equalsIgnoreCase("APR_UI")
							&& Oprkey.equalsIgnoreCase("Remarks")) // To

					{
						WebElement RemWebedit = Automation.driver.findElement(By.xpath(".//*[@id='addRemarks']"));
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								RemWebedit);
						RemWebedit.sendKeys(controlValue);
						Thread.sleep(100);
						WebElement RemButton = Automation.driver.findElement(By.xpath(".//*[@id='okRemarks']"));
						RemButton.click();
						Thread.sleep(100);
					}

				}

				else if (controlType.equalsIgnoreCase("Radio")) {
					String XPath = targetXpath + "//div/div/input";
					log.info("XPath of target Radio is : " + XPath);
					WebElement newelement = Automation.driver.findElement(By.xpath(XPath));
					Thread.sleep(100);
					((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
							newelement);
					Thread.sleep(100);
					((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();", newelement);// Testing
																												// 27/10/21

				} else if (controlType.equalsIgnoreCase("CheckBox")) {
					if (controlValue != "") {
						Thread.sleep(3000);
						String XPath = targetXpath + "//div/div/input";
						log.info("XPath of target CheckBox is : " + XPath);
						WebElement newelement = Automation.driver.findElement(By.xpath(XPath));
						if (controlValue.equalsIgnoreCase("Y") || controlValue.equalsIgnoreCase("Yes")) {
							if (!newelement.isSelected()) {
								// newelement.click();
								Thread.sleep(100);
								((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
										newelement);

								((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
										newelement);// Testing 27/10/21

								// newelement.click();
								Thread.sleep(100);
								if (!newelement.isSelected()) {
									((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
											newelement);
								}
								// ((JavascriptExecutor)Automation.driver).executeScript("arguments[0].scrollIntoView();",
								// newelement);
								Thread.sleep(100);
							}
						} else if (controlValue.equalsIgnoreCase("N") || controlValue.equalsIgnoreCase("No")) {
							if (newelement.isSelected()) {

								Thread.sleep(100);
								((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
										newelement);
								newelement.click(); // Mangesh--For unchecking
								// the checkbox
								Thread.sleep(100);
								// newelement.click();

							}
						} else if (controlValue.equalsIgnoreCase("") || StringUtils.isEmpty(controlValue)) {
							// return;
						}
					}

				} else if (controlType.equalsIgnoreCase("WebList")) {
					if (controlValue != "") {
						Thread.sleep(3000);
						String XPath = targetXpath + "//div/div/select";
						log.info("XPath of target WebList is : " + XPath);
						WebElement newelement = Automation.driver.findElement(By.xpath(XPath));
						ExpectedCondition<Boolean> isTextPresent = CommonExpectedConditions
								.textToBePresentInElement(newelement, controlValue);
						if (isTextPresent != null) {
							Thread.sleep(100);
							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
									newelement);
							if (Automation.browserType.toString().toUpperCase().contains("CHROME")
									|| Automation.browserType.toString().toUpperCase().contains("MSEDGE")) { // sonali
								// -chrome
								Actions actions = new Actions(Automation.driver);
								actions.moveToElement(newelement).click().perform();

								Select dropdown1 = new Select(newelement);
								Thread.sleep(100);
								dropdown1.selectByVisibleText(controlValue);
								Thread.sleep(100);
								if (!"GroupBilling".equalsIgnoreCase(Config.productTeam)) {
									try {
										if (newelement.isDisplayed()) {
											actions.moveToElement(newelement).click().perform();
										}
										if (controller.controllerTransactionType.toString()
												.equalsIgnoreCase("ChangeInstallmentSchedule")) {

											((JavascriptExecutor) Automation.driver)
													.executeScript("arguments[0].click();", newelement);
										}
									}

									catch (Exception ex) {
										log.error(ex.getMessage(), ex);
										log.info("Dropdown Selected");
									}
								}
							} else if (controlType.equalsIgnoreCase("ReplaceDefault")) {
								if (controlValue != "") {
									// String XPath1 = targetXpath + "//div/div/input";//Sel4
									log.info("XPath of target WebEdit is : " + XPath);
									WebElement newelement1 = Automation.driver.findElement(By.xpath(XPath));
									((JavascriptExecutor) Automation.driver)
											.executeScript("arguments[0].scrollIntoView();", newelement1);
									Thread.sleep(500);

									int counter1 = 1;
									String temp = "";
									while (!temp.equals(controlValue)) {

										JavascriptExecutor js1 = (JavascriptExecutor) Automation.driver;
										js1.executeScript("arguments[0].value = '';", newelement1);
										newelement1.sendKeys(controlValue);
										Thread.sleep(500);
										newelement1.sendKeys(Keys.TAB);
										newelement1.sendKeys(Keys.TAB);
										Thread.sleep(1000);
										temp = newelement1.getAttribute("value");
										if (temp.equals("")) {
											temp = newelement1.getText();
										}

										counter1++;
										if (counter1 == 2) {
											// System.out.println("Could not set correct value");
											break;
										}
									}
								}
							} else {
								newelement.click();
								Select dropdown1 = new Select(newelement);
								Thread.sleep(100);
								dropdown1.selectByVisibleText(controlValue);
								Thread.sleep(100);
								if (!"GroupBilling".equalsIgnoreCase(Config.productTeam)) {
									try {
										if (newelement.isDisplayed())// Meghna
										{
											newelement.click();
										}
									}

									catch (Exception ex) {
										log.error(ex.getMessage(), ex);
										log.info("Dropdown Selected");
									}
									// new
									// Select(newelement).selectByVisibleText(controlValue);
								}
							}
						}

					}
				}
			}
		} catch (Exception ex) {
			String errorMessage = ex.getMessage();
			String firstLine = errorMessage.split("\\R")[0];
			log.error("Error: " + firstLine);
			webDriver.getReport().setMessage("Error in getting the row in FIND : " + ex.getLocalizedMessage());
			webDriver.getReport().setStatus("FAIL");
			//StartRecovery.initiateRecovery();//Not required
			throw new Exception("Failed in FIND: " + controlName + firstLine );
			//throw new Exception("Error : " + ex.getMessage());
		}
	}

	// Below given code added for basebilling team to handle xpath for
	// FINDaction in an editable field***//
	public static void findActionEditable(WebElement tableFound, String controlName, String logicalName, Row rowValues,
			HashMap<String, Integer> valuesHeader) throws Exception {
		try {
			ColumnheaderIndex.clear();
			Operatectrlvalues.clear();
			Searchctrlvalues.clear();
			OperateControlType.clear();
			String Attempt = "";
			System.out.println("logicalName is : " + logicalName);
			log.info("logicalName is : " + logicalName);

			String currentHeader = null;
			String finalXpathTemp;
			int tempPosition;
			List<WebElement> table_Th = tableFound.findElements(By.tagName("th"));
			WebElement firstelement = table_Th.get(0);
			((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView(true);", firstelement);
			// to scroll table to the left

			for (int i = 0; i < table_Th.size(); i++) {
				currentHeader = (table_Th.get(i).getText()).toUpperCase(); // Meghna

				if (ColumnheaderIndex.containsKey(currentHeader) == false) {
					ColumnheaderIndex.put(currentHeader, i + 1);

				}
			}
			log.info("Application column headers and their indexes are : " + ColumnheaderIndex);

			String[] ExcelLogicalName = logicalName.split("\\|");
			String[] excelHeader = null;
			String[] actualHeader = null;
			String Sctrl = "";
			// int a = ColumnheaderIndex.get(currentHeader)

			excelHeader = ExcelLogicalName[0].split("\\:");
			actualHeader = ExcelLogicalName[1].split("\\:");

			// ***For GB - 24/07/2018***Deduction -- START***//
			int arrlen = ExcelLogicalName.length;
			if (((arrlen - 2) % 3) == 0) {
			} else {
				String AttemptColName = ExcelLogicalName[arrlen - 1];
				Cell ctrlValuecell111 = rowValues
						.getCell(Integer.parseInt(valuesHeader.get(AttemptColName).toString()));

				// ***Sel 4 Change***
				if (ctrlValuecell111 != null) {
					try {
						FormulaEvaluator evaluator = rowValues.getSheet().getWorkbook().getCreationHelper()
								.createFormulaEvaluator();
						DataFormatter fmt = new DataFormatter();
						// CellValue cellValue = evaluator.evaluate(ctrlValuecell111);
						Attempt = fmt.formatCellValue(ctrlValuecell111, evaluator);
						log.info("Attempt value: " + Attempt);
					} catch (Exception e) {
						// Handle any exceptions here
						//e.printStackTrace();
						log.error("Error occurred during formula evaluation: " + e.getMessage());
					}
				} else {
					Attempt = "";
				}
				// ***
				System.out.println(" Attempt no " + Attempt);
			}

			// ***Sel 4 Change***
			Workbook workbook = rowValues.getSheet().getWorkbook();
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

			int actualHeaderLen = excelHeader.length;
			//String ctrlValue = null;//Mandar
			
			
			for (int i = 0; i < actualHeaderLen; i++) {
				
				String ctrlValue = null;//Mandar - 18/07/2025
				
				Cell ctrlValuecell = rowValues.getCell(valuesHeader.get(excelHeader[i]));

				/*
				 * DataFormatter fmt = new DataFormatter(); if (ctrlValuecell == null) {
				 * ctrlValue = ""; }
				 * 
				 * else { CellType type = ctrlValuecell.getCellType(); switch (type) { case
				 * BLANK: ctrlValue = ""; break; case NUMERIC: ctrlValue =
				 * fmt.formatCellValue(ctrlValuecell); break; case STRING: ctrlValue =
				 * ctrlValuecell.getStringCellValue(); break; case BOOLEAN: ctrlValue =
				 * Boolean.toString(ctrlValuecell.getBooleanCellValue()); break; case ERROR:
				 * ctrlValue = "error"; break; case FORMULA: ctrlValue =
				 * ctrlValuecell.getCellFormula(); break; case _NONE: break; default: break; } }
				 */

				// ***Sel 4 Change***

				if (ctrlValuecell != null) {
					try {
						DataFormatter fmt = new DataFormatter();
						// CellValue cellValue = evaluator.evaluate(ctrlValuecell);
						ctrlValue = fmt.formatCellValue(ctrlValuecell, evaluator);
						Searchctrlvalues.put(actualHeader[i], ctrlValue);
						// Sctrl = Sctrl + ctrlValue;
					} catch (Exception e) {
						// Handle any exceptions here
						//e.printStackTrace();
						log.error("Error occurred during formula evaluation: " + e.getMessage());
					}
				}

				Searchctrlvalues.put(actualHeader[i], ctrlValue);
				Sctrl = Sctrl + ctrlValue;
			}

			if (Sctrl.equals("")) {
				log.info(
						"As per value sheet, there is nothing to be searched on application and hence coming out of FIND function");
				return;
			}
			log.info("Concatenated string in Values : " + Sctrl);
			log.info("Values to be searched on application table are : " + Searchctrlvalues);

			int logicalnameLen = ExcelLogicalName.length;
			for (int i = 2; i < logicalnameLen; i += 3) // for getting target  // columns and values
			{
				String ctrlValue = "";  // <-- declare and initialize here
				Cell ctrlValuecell = rowValues.getCell(valuesHeader.get(ExcelLogicalName[i + 2]));
				DataFormatter fmt = new DataFormatter();
				// Sel4 //Mandar - 18/7/2025
				/*
				 * if (ctrlValuecell == null) { ctrlValue = ""; }
				 */
				if (ctrlValuecell != null) {
					// Workbook workbook = rowValues.getSheet().getWorkbook();
					// FormulaEvaluator evaluator =
					// workbook.getCreationHelper().createFormulaEvaluator();
					try {
						CellValue cellValue = evaluator.evaluate(ctrlValuecell);

						if (cellValue != null) {

							ctrlValue = fmt.formatCellValue(ctrlValuecell, evaluator);
						}
					} catch (Exception e) {
						// Handle any exceptions here
						//e.printStackTrace();
						log.error("Error occurred during formula evaluation: " + e.getMessage());
					}
				}

				Operatectrlvalues.put(ExcelLogicalName[i], ctrlValue);
				OperateControlType.put(ExcelLogicalName[i], ExcelLogicalName[i + 1]);

			}
			log.info("Target column ctrlvalues are : " + Operatectrlvalues);
			log.info("Target column ctrltypes are : " + OperateControlType);

			// for creating XPath
			// String finalXpath = controlName+"//tr";
			String finalXpath = controlName + "/tbody/tr";
			Set<String> keys = Searchctrlvalues.keySet();
			int counter = 1;
			for (String key : keys) // For loop
			{
				// int position = (int) ColumnheaderIndex.get(key);
				int positionint;
				String position;

				try {
					positionint = (Integer.parseInt(key));
					position = "" + positionint;
				} catch (Exception pe) {

					position = (String) ColumnheaderIndex.get((key).toUpperCase()).toString().trim(); // Meghna-Updated

				}

				String ctrlvalue = (String) Searchctrlvalues.get(key).toString();
				log.info("Search key , position , value is : " + key + ": " + position + ": " + ctrlvalue);

				String datePattern = "\\d{4}-\\d{2}-\\d{2}";

				if (!ctrlvalue.equalsIgnoreCase(""))

				{

					if (ctrlvalue.matches(datePattern)) {
						if (counter == 1) {
							finalXpath = finalXpath + "/td[position()=" + position + " and (.//@title='"
									+ (ctrlvalue.trim()) + "' or .//text()='" + (ctrlvalue.trim()) + "')]";

						}

						else {
							finalXpath = finalXpath + "/../td[position()=" + position + " and (.//@title='"
									+ (ctrlvalue.trim()) + "' or .//text()='" + (ctrlvalue.trim()) + "')]";

						}
					}

					else if (!ctrlvalue.equalsIgnoreCase("BLANK")) {
						if (counter == 1) {
							finalXpath = finalXpath + "/td[position()=" + position + " and (.//@value='"
									+ (ctrlvalue.trim()) + "' or .//@title='" + (ctrlvalue.trim()) + "')]";
						}

						else {
							finalXpath = finalXpath + "/../td[position()=" + position + " and (.//@value='"
									+ (ctrlvalue.trim()) + "' or .//@title='" + (ctrlvalue.trim()) + "')]";
						}
					} else {
						if (counter == 1) {

							finalXpath = finalXpath + "/td[position()=" + position
									+ " and (.//@value='N' or  @title='')]";
						} else {
							finalXpath = finalXpath + "/../td[position()=" + position
									+ " and (.//@value='N' or  @title='')]";
						}
					}
					counter++;
				}
			}

			log.info("finalXpath is : " + finalXpath);

			try {
			    // Check if the original XPath exists
			    List<WebElement> elements = Automation.driver.findElements(By.xpath(finalXpath));

			    // Check if the finalXpath contains a date in MM/DD/YYYY format
			    String datePatternMMDDYYYY = "\\d{2}/\\d{2}/\\d{4}";
			    Pattern datePattern = Pattern.compile(datePatternMMDDYYYY);
			    Matcher matcher = datePattern.matcher(finalXpath);

			    // If no elements found initially, print a message about not finding the original XPath
			    if (elements.size() == 0) {
			        log.warn("Original XPath not found: " + finalXpath);
			    }

			    // Only proceed if the original XPath is NOT found and the date pattern is detected
			    if (elements.size() == 0 && matcher.find()) {
			        log.info("Date in MM/DD/YYYY format found in XPath. Attempting conversion to YYYY-MM-DD...");

			        // Extract the matched date from XPath
			        String dateInXpath = matcher.group(); // MM/DD/YYYY format

			        // Convert date format from MM/DD/YYYY to YYYY-MM-DD
			        String[] dateParts = dateInXpath.split("/");
			        if (dateParts.length == 3) {
			            String convertedDate = dateParts[2] + "-" + dateParts[0] + "-" + dateParts[1];

			            // Replace the original date in finalXpath with the converted format
			            String updatedXpath = finalXpath.replace(dateInXpath, convertedDate);
			            log.info("Converted XPath: " + updatedXpath);

			            // Check if the converted XPath exists
			            List<WebElement> updatedElements = Automation.driver.findElements(By.xpath(updatedXpath));

			            if (updatedElements.size() > 0) {
			                log.info("Converted XPath found successfully! Using updated XPath.");
			                finalXpath = updatedXpath; // Update finalXpath with the new XPath
			            } else {
			                log.info("Converted XPath not found. Proceeding with the original XPath.");
			            }
			        }
			    }

			    // Check finalXpath again after possible conversion
			    elements = Automation.driver.findElements(By.xpath(finalXpath));

			    if (elements.size() > 0) {
			        log.info("Element found successfully with the XPath: " + finalXpath);
			    } else {
			        log.warn("Element not found with the XPath: " + finalXpath);
			    }
			} catch (Exception e) {
			    log.error("Error while checking and converting date format: " + e.getMessage());
			}

			
			
			finalXpathTemp = finalXpath;

			// for operating on Target column
			Set<String> Operatekeys = Operatectrlvalues.keySet();
			String targetXpath = "";

			for (String Oprkey : Operatekeys) {
				String controlType = (String) OperateControlType.get(Oprkey).toString();
				String controlValue = (String) Operatectrlvalues.get(Oprkey).toString();

				int positionint;
				String position;

				try {
					positionint = (Integer.parseInt(Oprkey));
					position = "" + positionint;
				} catch (Exception pe) {
					position = (String) ColumnheaderIndex.get((Oprkey).toUpperCase()).toString().trim(); // Meghna-Updated
				}

				tempPosition = (Integer.parseInt(position)) + 1;
				log.info("Target key , position , value , type : " + Oprkey + ": " + position + ": " + controlValue
						+ ": " + controlType);

				targetXpath = finalXpath + "/../td[position()=" + position + "]";
				// System.out.println("targetXpath is : " + targetXpath);

				
				
				
				if (controlType.equalsIgnoreCase("WebEdit")) {
					String XPath = targetXpath + "//div/div/input";
					log.info("XPath of target WebEdit is : " + XPath);
					
					
					//WebElement newelement = Automation.driver.findElement(By.xpath(XPath));//Testing - 18/07/2025
					
					    WebElement newelement = Automation.driver.findElement(By.xpath(XPath));

					    if (newelement.isEnabled()) {
					        log.info("Element found and is enabled state");
					       // break;  // Exit from WebEdit block
					    
				
					// To clear a field in webtable--04/12/2017
					if (controlValue.equalsIgnoreCase("BLANK")) {
						// newelement.click();
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();", newelement);// Testing
																													// 27/10/21
						Thread.sleep(100);
						newelement.clear();
					}
					// To clear a field in webtable--04/12/2017

					else if (controlType != "" && controlValue != "") {
						// Thread.sleep(100);
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								newelement);
						Thread.sleep(100);// for EST_PR issue
						// newelement.click();
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();", newelement);// Testing
																													// 27/10/21
						Thread.sleep(100);
						newelement.clear();
						Thread.sleep(100);
						((JavascriptExecutor) Automation.driver).executeScript(
								"arguments[0].setAttribute('value', '" + controlValue + "')", newelement); // For
						newelement.clear();
						Thread.sleep(100);
						newelement.sendKeys(controlValue);

						Thread.sleep(100);
						// newelement.click();
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();", newelement);// Testing
																													// 27/10/21

						newelement.sendKeys(Keys.TAB);
						Thread.sleep(100);
					}
					    }
					    
					    else {
					    	 log.info("Element found but it is disabled. Skipping WebEdit");
					    	
					    	
					    }
				
					
					
				} else if (controlType.equalsIgnoreCase("WebButton")) {
					if (controlValue != "") {
						String XPath = targetXpath + "//div/div/button";
						log.info("XPath of target WebButton is : " + XPath);
						WebElement newelement = Automation.driver.findElement(By.xpath(XPath));
						// Thread.sleep(100);
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								newelement);
						Thread.sleep(100);
						// newelement.click();
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();", newelement);// Testing
																													// 27/10/21
					}
				} else if (controlType.equalsIgnoreCase("WebLink")) {
					if (controlValue != "") {

						if (controller.controllerTransactionType.toString().equalsIgnoreCase("ModifyAccountCurrent")
								&& (tempPosition == 12)) {
							WebElement headers = Automation.driver.findElement(
									By.xpath(".//*[@id='dgAddTransactionTemp']//*[@title='Transaction']/div"));
							// ((JavascriptExecutor)Automation.driver).executeScript("arguments[0].click();",
							// headers);
							Thread.sleep(100);
							headers.click();

							Boolean hclicked = false;

							while (!hclicked) {
								try {

									WebElement headers_asc = Automation.driver.findElement(By.xpath(
											".//*[contains(@id,'dgAddTransactionTemp')]/tbody/tr[1]/td[3]/div/div/span/span/span/span"));
									String first_value = headers_asc.getAttribute("title");

									log.info(first_value);

									if (first_value.equals("New Business")) {
										hclicked = true;
									} else {
										headers.click();
										first_value = "";
									}

								}

								catch (Exception ex) {
									log.error(ex.getMessage(), ex);
									log.info("Header element not found");
								}
							}
							Thread.sleep(15000);
						}

						String linkXPath = targetXpath + "//div/div/span";

						log.info("XPath of target WebLink is : " + linkXPath);
						WebElement newelement = Automation.driver.findElement(By.xpath(linkXPath));
						if (controller.controllerTransactionType.toString().equalsIgnoreCase("ModifyAccountCurrent")) {
							log.info("Do not scroll");
						}

						// ***Below given condition added for Basebilling Team -
						// To avoid scroll13-03-10-2019 ***
						else if (controller.controllerTransactionType.toString().equalsIgnoreCase("UndoWriteOff")) {
							log.info("Do not scroll");
							// ((JavascriptExecutor)Automation.driver).executeScript("arguments[0].click();",
							// newelement);
						} else {
							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
									newelement);

						}
						Thread.sleep(100);
						// newelement.click();
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();", newelement);// Testing
																													// 27/10/21
						Thread.sleep(500);
						log.info("Checks");

						// ******For Modify Acc Current --***
						int loopcntr = 0; // -Loop Control
						if (controller.controllerTransactionType.toString().equalsIgnoreCase("ModifyAccountCurrent")) {
							String tempValue;

							if (tempPosition == 12) {
								finalXpathTemp = finalXpathTemp + "/../td[position()=" + (tempPosition) + "]";
								WebElement tempEdit = Automation.driver.findElement(By.xpath(finalXpathTemp));

								tempValue = tempEdit.getAttribute("title");

								// while (tempValue.isEmpty())
								while (tempValue.isEmpty() && loopcntr < 50) // meghna--Loop
								// Control
								{
									if (newelement.isDisplayed()) {
										((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
												newelement);

									}
									tempValue = tempEdit.getAttribute("title");
									loopcntr = loopcntr + 1;
								}
							}

						}

					}
					// ******************* Below given code is to handle remarks
					// popup in the Find Action*************
					if (controller.controllerTransactionType.toString().equalsIgnoreCase("APR_UI")
							&& Oprkey.equalsIgnoreCase("Remarks")) // To
					{
						WebElement RemWebedit = Automation.driver.findElement(By.xpath(".//*[@id='addRemarks']"));
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								RemWebedit);
						RemWebedit.sendKeys(controlValue);
						Thread.sleep(100);
						WebElement RemButton = Automation.driver.findElement(By.xpath(".//*[@id='okRemarks']"));
						RemButton.click();
						Thread.sleep(100);
						// *************************//
					}

				}

				else if (controlType.equalsIgnoreCase("Radio")) {
					String XPath = targetXpath + "//div/div/input";
					log.info("XPath of target Radio is : " + XPath);
					WebElement newelement = Automation.driver.findElement(By.xpath(XPath));
					Thread.sleep(100);
					((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
							newelement);
					Thread.sleep(100);
					// newelement.click();
					((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();", newelement);// Testing
																												// 27/10/21
				} else if (controlType.equalsIgnoreCase("CheckBox")) {
					if (controlValue != "") {
						log.info("controlValue is:" + controlValue);
						Thread.sleep(3000);
						String XPath = targetXpath + "//div/div/input";
						log.info("XPath of target CheckBox is : " + XPath);
						WebElement newelement = Automation.driver.findElement(By.xpath(XPath));
						if (controlValue.equalsIgnoreCase("Y") || controlValue.equalsIgnoreCase("Yes")) {
							if (!newelement.isSelected()) {
								// newelement.click();
								Thread.sleep(100);
								((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
										newelement);
								// newelement.click();
								((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
										newelement);// Testing 27/10/21
								Thread.sleep(100);
								if (!newelement.isSelected()) {
									((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
											newelement);
								}
								// ((JavascriptExecutor)Automation.driver).executeScript("arguments[0].scrollIntoView();",
								// newelement);
								Thread.sleep(100);
							}
						} else if (controlValue.equalsIgnoreCase("N") || controlValue.equalsIgnoreCase("No")) {
							if (newelement.isSelected()) {

								Thread.sleep(100);
								((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
										newelement);
								((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
										newelement);// Testing 27/10/21
								Thread.sleep(100);
								// newelement.click();
							}
						} 
					}

				} else if (controlType.equalsIgnoreCase("WebList")) {
					if (controlValue != "") {
						Thread.sleep(3000);
						String XPath = targetXpath + "//div/div/select";
						log.info("XPath of target WebList is : " + XPath);
						WebElement newelement = Automation.driver.findElement(By.xpath(XPath));
						ExpectedCondition<Boolean> isTextPresent = CommonExpectedConditions
								.textToBePresentInElement(newelement, controlValue);
						if (isTextPresent != null) {
							Thread.sleep(100);
							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
									newelement);
							if (Automation.browserType.toString().toUpperCase().contains("CHROME")
									|| Automation.browserType.toString().toUpperCase().contains("MSEDGE")) {
								Actions actions = new Actions(Automation.driver);
								actions.moveToElement(newelement).click().perform();

								Select dropdown1 = new Select(newelement);
								Thread.sleep(100);
								dropdown1.selectByVisibleText(controlValue);
								Thread.sleep(100);
								if (!"GroupBilling".equalsIgnoreCase(Config.productTeam)) {
									try {
										if (newelement.isDisplayed()) {
											actions.moveToElement(newelement).click().perform();
										}
									}

									catch (Exception ex) {
										log.error(ex.getMessage(), ex);
										log.info("Dropdown Selected");
									}
									// new
									// Select(newelement).selectByVisibleText(controlValue);
								}
							} else if (controlType.equalsIgnoreCase("ReplaceDefault")) {
								if (controlValue != "") {
									// String XPath1 = targetXpath + "//div/div/input";//Sel4
									log.info("XPath of target WebEdit is : " + XPath);
									WebElement newelement1 = Automation.driver.findElement(By.xpath(XPath));
									((JavascriptExecutor) Automation.driver)
											.executeScript("arguments[0].scrollIntoView();", newelement1);
									Thread.sleep(500);

									int counter1 = 1;
									String temp = "";
									while (!temp.equals(controlValue)) {

										JavascriptExecutor js1 = (JavascriptExecutor) Automation.driver;
										js1.executeScript("arguments[0].value = '';", newelement1);
										newelement1.sendKeys(controlValue);
										Thread.sleep(500);
										newelement1.sendKeys(Keys.TAB);
										newelement1.sendKeys(Keys.TAB);
										Thread.sleep(1000);
										temp = newelement1.getAttribute("value");
										if (temp.equals("")) {
											temp = newelement1.getText();
										}

										counter1++;
										if (counter1 == 2) {
											// System.out.println("Could not set correct value");
											break;
										}
									}
								}
							} else {
								newelement.click();
								Select dropdown1 = new Select(newelement);
								Thread.sleep(100);
								dropdown1.selectByVisibleText(controlValue);
								Thread.sleep(100);
								if (!"GroupBilling".equalsIgnoreCase(Config.productTeam)) {
									try {
										if (newelement.isDisplayed()) {
											newelement.click();
										}
									}

									catch (Exception ex) {
										log.error(ex.getMessage(), ex);
										log.info("Dropdown Selected");
									}
									// new
									// Select(newelement).selectByVisibleText(controlValue);
								}
							}
						}

					}
				}

				/* Added to download pdf through find logic--07/05/2020--start - BaseBilling */
				else if (controlType.equalsIgnoreCase("DownloadDocument")) {

					String path = Config.runtimeFileDownloadFolder;
					File dir = new File(path);
					// List<File> fileList;//Sel4
					File[] files = dir.listFiles();
					for (File file : files) {
						if (file.getName().endsWith("zip") || file.getName().endsWith("pdf")
								|| file.getName().endsWith("xls") || file.getName().endsWith("xlsx")) {
							file.delete();
						}
					}

					String linkXPath = targetXpath + "//div/div/span";
					WebElement newelement = Automation.driver.findElement(By.xpath(linkXPath));
					Actions MousebuilderClick1 = new Actions(Automation.driver);
					Action MouseclickAction1 = MousebuilderClick1.moveToElement(newelement).clickAndHold().release()
							.build();

					MouseclickAction1.perform();
					break;
				}
				/* Added to download pdf through find logic--07/05/2020--End */
			}

		} catch (Exception ex) {
			log.error(ex.getMessage(), ex);
			// iTAFSeleniumWeb.WebDriver.report.setStrMessage(ex.getLocalizedMessage());
			webDriver.getReport().setMessage("Error in getting the row in FIND : " + ex.getLocalizedMessage());
			webDriver.getReport().setStatus("FAIL");
			StartRecovery.initiateRecovery();
			throw new Exception("Failed in FIND: " + controlName + " <-|-> LocalizeMessage " + ex.getLocalizedMessage()
					+ " <-|-> Message " + ex.getMessage() + " <-|-> Cause " + ex.getCause());

		}
	}

	public static void TableInputAction(WebElement tableFound, String controlName, String logicalName, Row rowValues12,
			HashMap<String, Integer> valuesHeader, ArrayList<Integer> valuesheetrowsnum) throws Exception {
		try {
			log.info("logicalName is : " + logicalName);
			ColumnheaderIndex.clear();
			Operatectrlvalues.clear();
			OperateControlType.clear();

			String rowvalue = "";
			String currentHeader = null;

			// For ChecksEntry Edit Flow--01/12/2017
			// int TIRow = 1;
			int TIRow;
			if (controller.controllerTransactionType.toString().equalsIgnoreCase("ChecksEntry")
					&& logicalName.contains("Edit_")) {
				TIRow = 2; // Only for Edit specific scenario--
			} else {
				TIRow = 1; // for all TableInput scenarios--
			}
			// For ChecksEntry Edit Flow--01/12/2017

			InputStream TImyXls = new FileInputStream(WebHelper.TIFilePath);
			Workbook TIworkBook = null; // new XSSFWorkbook(TImyXls);
			if (WebHelper.TIFilePath.endsWith(".xls")) {
				TIworkBook = new HSSFWorkbook(TImyXls);
			} else if (WebHelper.TIFilePath.endsWith(".xlsx") | WebHelper.TIFilePath.endsWith(".xlsm")) {
				TIworkBook = new XSSFWorkbook(TImyXls);
			}

			TIformat = TIworkBook.createDataFormat();
			Sheet TIsheetStructure = TIworkBook.getSheet("Values");

			List<WebElement> table_Th = tableFound.findElements(By.tagName("th"));
			for (int i = 0; i < table_Th.size(); i++) {
				// currentHeader = table_Th.get(i).getText();
				currentHeader = (table_Th.get(i).getText()).toUpperCase();
				if (ColumnheaderIndex.containsKey(currentHeader) == false) {
					ColumnheaderIndex.put(currentHeader, i + 1);
				}
			}
			log.info("Application column headers and their indexes are : " + ColumnheaderIndex);

			String[] ExcelLogicalName = logicalName.split("\\|");
			int logicalnameLen = ExcelLogicalName.length;
			String ctrlValue = null;

			for (int valuesheetrow : valuesheetrowsnum) {
				Operatectrlvalues.clear();
				OperateControlType.clear();
				for (int i = 0; i < logicalnameLen; i += 3) // for getting
				// target columns
				// and values
				{
					// System.out.println("XXX");
					ctrlValue = WebHelperUtil.getCellData((ExcelLogicalName[i + 2].toString()), TIsheetStructure,
							valuesheetrow, valuesHeader);

					Operatectrlvalues.put(ExcelLogicalName[i], ctrlValue);
					OperateControlType.put(ExcelLogicalName[i], ExcelLogicalName[i + 1]);
					rowvalue = rowvalue + ctrlValue;

				}

				log.info("Target column ctrlvalues are : " + Operatectrlvalues);
				log.info("Target column ctrltypes are : " + OperateControlType);
				if (rowvalue.equals("")) {
					log.info(
							"As per value sheet, there is nothing to input on application further and hence coming out of TableInput function");
					return;
				}
				rowvalue = "";
				Set<String> Operatekeys = Operatectrlvalues.keySet();
				String targetXpath = "";

				for (String Oprkey : Operatekeys) {
					String controlType = (String) OperateControlType.get(Oprkey).toString();
					String controlValue = (String) Operatectrlvalues.get(Oprkey).toString();

					if (controlValue.contains("${RDATETIME}")) {
						for (int i = 1; i <= 10; i++) {
							if (controlValue.startsWith("POL") && controlValue.contains("${RDATETIME}" + i)) {
								controlValue = WebHelperUtil.ReadFromExcel("", "PAS_PolicyNum" + i);
							} else if (controlValue.startsWith("AGT") && controlValue.contains("${RDATETIME}" + i)) {
								controlValue = WebHelperUtil.ReadFromExcel("", "ProducerCode" + i);
							} else if (controlValue.startsWith("ACC") && controlValue.contains("${RDATETIME}" + i)) {
								controlValue = WebHelperUtil.ReadFromExcel("", "AccountNum" + i);
							} else {
								String uniqueNum = WebHelperUtil.ReadFromExcel("", "UniqueNum");
								controlValue = controlValue.replace("${RDATETIME}", uniqueNum);
							}
						}
					}

					int positionint;
					String position;

					try {
						positionint = (Integer.parseInt(Oprkey));
						position = "" + positionint;
					} catch (Exception pe) {
						// position = (String)
						// ColumnheaderIndex.get(Oprkey).toString().trim();
						position = (String) ColumnheaderIndex.get((Oprkey).toUpperCase()).toString().trim();
						// log.error(pe.getMessage(), pe);
					}

					log.info("Target key , position , value , type : " + Oprkey + ": " + position + ": " + controlValue
							+ ": " + controlType);
					String actualcontrolname = controlName;
					targetXpath = actualcontrolname + "/tbody/tr[" + TIRow + "]/td[" + position + "]";

					if (controlType.equalsIgnoreCase("WebEdit")) // operate on
					// targeted
					// column
					{
						String XPath = targetXpath + "//div/div/input";
						log.info("XPath of target WebEdit is : " + XPath);
						WebElement newelement = Automation.driver.findElement(By.xpath(XPath));
						log.info("controlType is : " + controlType + " && " + "controlValue is : " + controlValue);// Mandar

						// 01/12/2017--Ui Validation-- To clear fields in
						// webtable
						if (controlValue.equalsIgnoreCase("BLANK")) {
							newelement.click();
							Thread.sleep(100);
							newelement.clear();
						}
						// 01/12/2017--Ui Validation-- To clear fields in
						// webtable

						else if (controlType != "" && controlValue != "") {
							Thread.sleep(100);
							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
									newelement);
							Thread.sleep(100);
							newelement.click();
							Thread.sleep(100);
							newelement.clear();
							Thread.sleep(100);
							newelement.sendKeys(controlValue);
							Thread.sleep(2000);
							// newelement.click();// to avoid remarks
							// click on payment screen (UI)
							// newelement.sendKeys(Keys.TAB);
							// Thread.sleep(2000);
							((JavascriptExecutor) Automation.driver).executeScript(
									"arguments[0].setAttribute('value', '" + controlValue + "')", newelement); // For
							newelement.clear();
							Thread.sleep(100);
							newelement.sendKeys(controlValue);
							Thread.sleep(100);
							newelement.click();
							newelement.sendKeys(Keys.TAB);
							Thread.sleep(100);
							// newelement.sendKeys(Keys.ESCAPE);// for
							// back to back date calendars--Commented as this
							// was causing failure in ChecksEntry
							if (controller.controllerTransactionType.toString().equalsIgnoreCase("CreateDeposit")
									&& Oprkey.equalsIgnoreCase("*Apply To #"))

							{
								try {
									WebElement ApplyToWebedit = Automation.driver
											.findElement(By.xpath(".//*[contains(@class,'modal-title')]"));
									((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
											ApplyToWebedit);
									ApplyToWebedit.sendKeys(controlValue);
									Thread.sleep(100);
									WebElement ApplyToButton = Automation.driver
											.findElement(By.xpath(".//*[@name='Ok']"));
									ApplyToButton.click();
									Thread.sleep(100);
								} catch (Exception e) {
									log.error(e.getMessage(), e);
									log.info("Warning message on Apply To # handle and tLocalizedMessage"
											+ e.getLocalizedMessage());
								}
							}
						}

					} else if (controlType.equalsIgnoreCase("WebButton")) {
						String XPath = targetXpath + "//div/div";
						log.info("XPath of target WebButton is : " + XPath);
						WebElement newelement = Automation.driver.findElement(By.xpath(XPath));
						Thread.sleep(100);
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								newelement);
						Thread.sleep(100);
						newelement.click();
					} else if (controlType.equalsIgnoreCase("WebLink")) {

						if (controlValue != "") {
							String linkXPath = targetXpath + "//div/div/span";
							log.info("XPath of target WebLink is : " + linkXPath);
							WebElement newelement = Automation.driver.findElement(By.xpath(linkXPath));
							Thread.sleep(100);
							Thread.sleep(100);
							((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
									newelement);
							Thread.sleep(100);
							newelement.click();
							if ((controller.controllerTransactionType.toString().equalsIgnoreCase("ChecksEntry")
									&& Oprkey.equalsIgnoreCase("Remarks"))
									|| (controller.controllerTransactionType.toString()
											.equalsIgnoreCase("CreateDeposit") && Oprkey.equalsIgnoreCase("Remarks"))) // To

							{
								WebElement RemWebedit = Automation.driver
										.findElement(By.xpath(".//*[@id='addRemarks']"));
								((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
										RemWebedit);
								RemWebedit.clear();
								Thread.sleep(100);
								RemWebedit.sendKeys(controlValue);
								((JavascriptExecutor) Automation.driver).executeScript(
										"arguments[0].setAttribute('value', '" + controlValue + "')", RemWebedit); // For
								RemWebedit.clear(); // -- newelement
								// Changed to RemWebedit
								Thread.sleep(100);
								RemWebedit.sendKeys(controlValue);

								Thread.sleep(100);
								WebElement RemButton = Automation.driver.findElement(By.xpath(".//*[@id='okRemarks']"));
								RemButton.click();
								Thread.sleep(100);
							}

						}
					} else if (controlType.equalsIgnoreCase("Radio")) {
						String XPath = targetXpath + "//div/div/input";
						log.info("XPath of target Radio is : " + XPath);
						WebElement newelement = Automation.driver.findElement(By.xpath(XPath));
						Thread.sleep(100);
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
								newelement);
						Thread.sleep(100);
						newelement.click();
					} else if (controlType.equalsIgnoreCase("CheckBox")) {
						String XPath = targetXpath + "//div/div/input";
						log.info("XPath of target CheckBox is : " + XPath);
						WebElement newelement = Automation.driver.findElement(By.xpath(XPath));
						if (controlValue.equalsIgnoreCase("Y") || controlValue.equalsIgnoreCase("Yes")) {
							if (!newelement.isSelected()) {
								// newelement.click();
								Thread.sleep(100);
								((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
										newelement);
								newelement.click();
								Thread.sleep(100);
								if (!newelement.isSelected()) {
									((JavascriptExecutor) Automation.driver).executeScript("arguments[0].click();",
											newelement);
								}
								// ((JavascriptExecutor)Automation.driver).executeScript("arguments[0].scrollIntoView();",
								// newelement);
								Thread.sleep(100);
							}
						} else if (controlValue.equalsIgnoreCase("N") || controlValue.equalsIgnoreCase("No")) {
							if (newelement.isSelected()) {

								Thread.sleep(100);
								((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
										newelement);
								Thread.sleep(100);
								newelement.click();
							}
						} else if (controlValue.equalsIgnoreCase("") || StringUtils.isEmpty(controlValue)) {
							// return;
						}
					} else if (controlType.equalsIgnoreCase("WebList")) {
						if (controlValue != "") {
							String XPath = targetXpath + "//div/div/select";
							log.info("XPath of target WebList is : " + XPath);
							WebElement newelement = Automation.driver.findElement(By.xpath(XPath));
							ExpectedCondition<Boolean> isTextPresent = CommonExpectedConditions
									.textToBePresentInElement(newelement, controlValue);
							if (isTextPresent != null) {
								Thread.sleep(100);
								((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
										newelement);
								if (Automation.browserType.toString().toUpperCase().contains("CHROME")
										|| Automation.browserType.toString().toUpperCase().contains("MSEDGE")) {
									Actions actions = new Actions(Automation.driver);
									actions.moveToElement(newelement).click().perform();

									Select dropdown1 = new Select(newelement);
									Thread.sleep(100);
									dropdown1.selectByVisibleText(controlValue);
									Thread.sleep(100);
									if (!"GroupBilling".equalsIgnoreCase(Config.productTeam)) {
										try {
											if (newelement.isDisplayed())// Meghna
											{
												actions.moveToElement(newelement).click().perform();
											}
										}

										catch (Exception ex) {
											log.error(ex.getMessage(), ex);
											log.info("Dropdown Selected");
										}
										// new
										// Select(newelement).selectByVisibleText(controlValue);
									}
								} else {
									newelement.click();
									Select dropdown1 = new Select(newelement);
									Thread.sleep(100);
									dropdown1.selectByVisibleText(controlValue);
									Thread.sleep(100);

									newelement.click();
									Thread.sleep(100);
									// new
									// Select(newelement).selectByVisibleText(controlValue);
								}
							}
						}
					}
				}

				if ((controller.controllerTransactionType.toString().equalsIgnoreCase("ChecksEntry"))
						|| (controller.controllerTransactionType.toString().equalsIgnoreCase("ManualEntry"))
						|| (controller.controllerTransactionType.toString().equalsIgnoreCase("CreateDeposit")))

				{
					TIRow = TIRow + 1;
				}

			}

		}

		catch (Exception ex) {
			log.error(ex.getMessage(), ex);
			webDriver.getReport().setMessage("Error in TableInputAction : " + ex.getLocalizedMessage());
			webDriver.getReport().setStatus("FAIL");
			StartRecovery.initiateRecovery();
			throw new Exception("Failed in TableInput: " + controlName + " <-|-> LocalizeMessage "
					+ ex.getLocalizedMessage() + " <-|-> Message " + ex.getMessage() + " <-|-> Cause " + ex.getCause());
		}

	}

}
