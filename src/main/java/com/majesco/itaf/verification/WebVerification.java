package com.majesco.itaf.verification;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import com.majesco.itaf.main.Automation;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.ITAFWebDriver;
import com.majesco.itaf.main.ITAFWebDriverPAS;
import com.majesco.itaf.main.ITAFWebDriverDM;
import com.majesco.itaf.main.MainController;
import com.majesco.itaf.main.MainControllerClaims;
import com.majesco.itaf.main.MainControllerSuite;
import com.majesco.itaf.main.ObjectFactory;
import com.majesco.itaf.main.WebHelper;
import com.majesco.itaf.main.WebHelperBilling;
import com.majesco.itaf.main.WebHelperClaims;
import com.majesco.itaf.main.WebHelperPAS;
import com.majesco.itaf.main.WebHelperUtil;
import com.majesco.itaf.util.Constants;
import com.majesco.itaf.util.DMProduct;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.util.JDBCConnection;
import com.majesco.itaf.vo.Reporter;

@SuppressWarnings("unused")
public class WebVerification {

	private final static Logger log = LogManager.getLogger(WebVerification.class.getName());
	private static HashMap<String, Integer> vTableListMap = new HashMap<String, Integer>();
	private static HashMap<String, Integer> templateMap = new HashMap<String, Integer>();
	private static List<String> columns = new ArrayList<String>();
	private static List<String> columnsData = new ArrayList<String>();
	private static Date vdate = null;
	private static HashMap<String, Integer> inputHashTable = new HashMap<String, Integer>();
	private static List<List<String>> rows = new ArrayList<List<String>>();
	public static int currentRowIndex = 0;
	private static boolean isTableFound = false;
	public static boolean isFromVerification = false;
	private static Date dt = null;
	private static String failedMsg = null;
	private static final String EXCEL_FILE_EXTENSTION = ".xlsx";
	private static boolean isClaimsApplication = ITAFWebDriver.isClaimsApplication();
	private static boolean isBillingApplication = ITAFWebDriver.isBillingApplication();
	public static String tempStatus = null;
	public static int failedcount = 0;
	private static MainController controller = ObjectFactory.getMainController();

	// ***below given performVerification is used for Billing***
	public static void performVerificationBilling(String transactionType, String testcaseID, String operationtype,
			String cycleDate) throws IOException, Exception {

		if (ITAFWebDriver.isSuiteApplication() && MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims"))
			isClaimsApplication = true;
		if (ITAFWebDriver.isSuiteApplication() && MainControllerSuite.transactionSilo.equalsIgnoreCase("Billing"))
			isBillingApplication = true;

		Sheet vTableSheet = WebHelperUtil.getSheet(Config.verificationTableISTPath, "VerificationTables");
		int rowCount = vTableSheet.getLastRowNum() + 1;
		vTableListMap = WebHelperUtil.getValueFromHashMap(vTableSheet);
		Reporter report = new Reporter();
		Sheet actualSheet = null;
		WebElement tableElement = null;
		List<WebElement> rowElements = null;
		String ActualPath = null;
		String duplicateActualPath = null;
		String expectedSheetPath = null;
		String SheetName = "ActualValues";
		vdate = new Date();
		WebHelper.frmDate = new Date();
		for (int rowIndex = 1; rowIndex < rowCount && !controller.pauseExecution; rowIndex++) {
			try {
				Row vRow = vTableSheet.getRow(rowIndex);
				String executeFlag = WebHelperUtil.getCellData("Verify", vTableSheet, rowIndex, inputHashTable);
				String vTransaction = WebHelperUtil.getCellData("TransactionType", vTableSheet, rowIndex,
						inputHashTable);

				if (isTableFound == true && !vTransaction.equalsIgnoreCase(transactionType)) {
					break;
				}

				if (executeFlag.toString().equalsIgnoreCase("Y")
						&& vTransaction.toString().equalsIgnoreCase(transactionType.toString())) {
					isTableFound = true;
					String functionalFlag = WebHelperUtil.getCellData("Functional", vTableSheet, rowIndex,
							inputHashTable);
					String templateDir = WebHelperUtil.getCellData("TemplateAdditionalPath", vTableSheet, rowIndex,
							inputHashTable);
					String templateSheet = WebHelperUtil.getCellData("TemplateSheet", vTableSheet, rowIndex,
							inputHashTable);
					String expectedDirPath = WebHelperUtil.getCellData("ExpectedDataAdditionalPath", vTableSheet,
							rowIndex, inputHashTable);
					String expectedSheet = WebHelperUtil.getCellData("Expected", vTableSheet, rowIndex, inputHashTable);
					String templatePath = Config.verificationTemplatePath + templateDir.toString() + "\\"
							+ templateSheet.toString();
					log.info(templatePath);

					ActualPath = Config.actualValuesValuesPath + expectedDirPath + "\\" + transactionType + "_Actual"
							+ EXCEL_FILE_EXTENSTION;
					duplicateActualPath = Config.actualValuesValuesPath + expectedDirPath + "\\" + transactionType
							+ "_Actual_duplicate" + EXCEL_FILE_EXTENSTION;

					expectedSheetPath = Config.expectedValuesValuesPath + expectedDirPath.toString() + "\\"
							+ expectedSheet.toString();
					Sheet layoutSheet = WebHelperUtil.getSheet(templatePath, "Layout");
					int templateRowCount = layoutSheet.getLastRowNum() + 1;
					templateMap = WebHelperUtil.getValueFromHashMap(layoutSheet);

					for (int templateIndex = 1; templateIndex < templateRowCount
							&& !controller.pauseExecution; templateIndex++) {
						Row layoutRow = layoutSheet.getRow(templateIndex);
						String tableID = WebHelperUtil.getCellData("TableID", layoutSheet, templateIndex,
								inputHashTable);
						String tableType = WebHelperUtil.getCellData("TableType", layoutSheet, templateIndex,
								inputHashTable);
						String tableIDType = WebHelperUtil.getCellData("TableIDType", layoutSheet, templateIndex,
								inputHashTable);
						String startRow = WebHelperUtil.getCellData("StartRow", layoutSheet, templateIndex,
								inputHashTable);
						String endRow = WebHelperUtil.getCellData("EndRow", layoutSheet, templateIndex, inputHashTable);
						String columnName = WebHelperUtil.getCellData("ColumnName", layoutSheet, templateIndex,
								inputHashTable);
						String rowNo = WebHelperUtil.getCellData("Row", layoutSheet, templateIndex, inputHashTable);
						String colNo = WebHelperUtil.getCellData("Column", layoutSheet, templateIndex, inputHashTable);
						String controlName = WebHelperUtil.getCellData("ControlName", layoutSheet, templateIndex,
								inputHashTable);
						String controlType = WebHelperUtil.getCellData("ControlType", layoutSheet, templateIndex,
								inputHashTable);
						String controlID = WebHelperUtil.getCellData("ControlID", layoutSheet, templateIndex,
								inputHashTable);
						report.setTestcaseId(testcaseID.toString());
						report.setTrasactionType(vTransaction.toString());
						report.setTestDescription(controller.testDescription);
						report.setFromDate(Config.dtFormat.format(vdate));
						report.setIteration(Config.cycleNumber);
						report.setStatus("PASS");
						report.setMessage(" ");

						if (tableType.equalsIgnoreCase("NonUniform")) {
							columns.add("TestCaseID");
							columns.add("TransactionType");
							if (isClaimsApplication) {
								columns.add("Header");
								columns.add("CycleDate");
							}
							columns.add("CurrentDate");
							columnsData = new ArrayList<String>();
							columnsData.add(testcaseID.toString());
							columnsData.add(vTransaction.toString());
							if (isClaimsApplication) {
								columnsData.add(MainControllerClaims.mainControllerHeaderNo);
								columnsData.add(cycleDate);
							}
							columnsData.add(Config.dtFormat.format(vdate));

							int colNoInt;
							String strXPath;

							for (templateIndex = 1; templateIndex < templateRowCount; templateIndex++) {
								layoutRow = layoutSheet.getRow(templateIndex);
								columnName = WebHelperUtil.getCellData("ColumnName", layoutSheet, templateIndex,
										inputHashTable);
								columns.add(columnName.toString());

								rowNo = WebHelperUtil.getCellData("Row", layoutSheet, templateIndex, inputHashTable);
								colNo = WebHelperUtil.getCellData("Column", layoutSheet, templateIndex, inputHashTable);
								colNoInt = Integer.parseInt(colNo);
								int rowNoInt = Integer.parseInt(rowNo);
								strXPath = tableID + "/tbody/tr[" + (rowNoInt + 1) + "]/td[" + (colNoInt + 1) + "]";
								String sVal = Automation.driver.findElement(By.xpath(strXPath)).getText();
								if (sVal.equalsIgnoreCase("")) {
									sVal = Automation.driver.findElement(By.xpath(strXPath)).getAttribute("value");
								}
								columnsData.add(sVal);
							}
							rows.add(columnsData);

						} else if (tableType.equalsIgnoreCase("DB")) {
							if (tableIDType.equalsIgnoreCase("Expected")) {
								SheetName = "Expected";
								ActualPath = Config.expectedValuesValuesPath + expectedDirPath + "\\" + transactionType
										+ "_Expected" + EXCEL_FILE_EXTENSTION;
							}
							ResultSet rs = null;
							Connection conn = JDBCConnection.establishDBConn();
							Statement st = conn.createStatement();
							rs = st.executeQuery(tableID);
							isTableFound = true;
							columns.add("TestCaseID");
							columns.add("TransactionType");
							if (isClaimsApplication) {
								columns.add("Header");
								columns.add("CycleDate");
							}
							columns.add("CurrentDate");

							for (int rIndex = Integer.parseInt(startRow.toString()); rs.next(); rIndex++)// rowElements.size()
							{
								columnsData = new ArrayList<String>();
								columnsData.add(testcaseID.toString());
								columnsData.add(vTransaction.toString());
								if (isClaimsApplication) {
									columnsData.add(MainControllerClaims.mainControllerHeaderNo);
									columnsData.add(cycleDate);
								}
								columnsData.add(Config.dtFormat.format(vdate));

								for (templateIndex = 1; templateIndex < templateRowCount; templateIndex++) {
									layoutRow = layoutSheet.getRow(templateIndex);
									columnName = WebHelperUtil.getCellData("ColumnName", layoutSheet, templateIndex,
											inputHashTable);
									colNo = WebHelperUtil.getCellData("Column", layoutSheet, templateIndex,
											inputHashTable);

									if (rIndex == Integer.parseInt(startRow.toString())) {
										columns.add(columnName.toString());
									}
									columnsData.add(rs.getString(Integer.parseInt(colNo)));
								}
								rows.add(columnsData);
							}
							rs.close();
							st.close();
							JDBCConnection.closeConnection(conn);

						} else if (tableType.equalsIgnoreCase("Uniform")) {
							int rCount = 0;
							columns.add("TestCaseID");
							columns.add("TransactionType");

							if (isClaimsApplication) {
								columns.add("Header");
								columns.add("CycleDate");
							}
							if (isBillingApplication) {
								columns.add("CycleDate");

							}
							if (!(isBillingApplication)) {
								columns.add("CurrentDate");
							}

							int colNoInt;
							String strXPath;

							tableElement = getElementByType(tableIDType, tableID, "");
							rowElements = tableElement.findElements(By.xpath(tableID + "/tbody/tr"));

							if (endRow.equalsIgnoreCase("0") || endRow.equalsIgnoreCase("")) {
								rCount = rowElements.size();
							} else {
								rCount = Integer.parseInt(endRow);
								if (rCount > rowElements.size()) {
									rCount = rowElements.size();
								}
							}

							for (int rIndex = Integer.parseInt(startRow.toString()); rIndex < rCount; rIndex++) {
								columnsData = new ArrayList<String>();
								columnsData.add(testcaseID.toString());
								columnsData.add(vTransaction.toString());

								if (isClaimsApplication) {
									columnsData.add(MainControllerClaims.mainControllerHeaderNo);
									columnsData.add(cycleDate);
								}

								if (!(isBillingApplication)) {
									columnsData.add(Config.dtFormat.format(vdate));
								}
								if (isBillingApplication) {
									columnsData.add(cycleDate);
								}

								for (templateIndex = 1; templateIndex < templateRowCount; templateIndex++) {
									layoutRow = layoutSheet.getRow(templateIndex);
									columnName = WebHelperUtil.getCellData("ColumnName", layoutSheet, templateIndex,
											inputHashTable);
									colNo = WebHelperUtil.getCellData("Column", layoutSheet, templateIndex,
											inputHashTable);
									controlType = WebHelperUtil.getCellData("ControlType", layoutSheet, templateIndex,
											inputHashTable);
									colNoInt = Integer.parseInt(colNo);

									if (rIndex == Integer.parseInt(startRow.toString())) {
										columns.add(columnName.toString());
									}

									if (controlType == "" || controlType.equalsIgnoreCase(null)) {
										if (columnName.contains("Policy #")) {

											try {

												strXPath = tableID + "/tbody/tr[" + (rIndex + 1) + "]/td["
														+ (colNoInt + 1) + "]" + "/div/span";
												WebElement newelement = Automation.driver
														.findElement(By.xpath(strXPath));
												((JavascriptExecutor) Automation.driver)
														.executeScript("arguments[0].click();", newelement);

											} catch (Exception e) {
												// log.error(e.getMessage(), e);
												strXPath = null;
											}

										}

										strXPath = tableID + "/tbody/tr[" + (rIndex + 1) + "]/td[" + (colNoInt + 1)
												+ "]";

										String sVal = Automation.driver.findElement(By.xpath(strXPath)).getText();
										columnsData.add(sVal);

									} else {
										strXPath = tableID + "/tbody/tr[" + (rIndex + 1) + "]/td[" + (colNoInt + 1)
												+ "]/" + controlName;// TM-10/02/2015:
										String sVal = Automation.driver.findElement(By.xpath(strXPath))
												.getAttribute("value");
										columnsData.add(sVal);
									}
								}
								rows.add(columnsData);
							}
							break;
						} else if (tableType.equalsIgnoreCase("UniformDynamic")) {
							columns.add("TestCaseID");
							columns.add("TransactionType");
							if (isClaimsApplication) {
								columnsData.add(MainControllerClaims.mainControllerHeaderNo);
								columns.add("Header");
								columns.add("CycleDate");
							}
							columns.add("CurrentDate");
							int colNoInt;
							String strXPath;

							if (!Config.seleniumExecution.equalsIgnoreCase("RC")) {
								tableElement = getElementByType(tableIDType, tableID, "");
								rowElements = tableElement.findElements(By.tagName("tr"));
							}

							for (int rIndex = Integer.parseInt(startRow.toString()); rIndex < rowElements
									.size(); rIndex++) {
								columnsData = new ArrayList<String>();
								columnsData.add(testcaseID.toString());
								columnsData.add(vTransaction.toString());
								if (isClaimsApplication) {
									columnsData.add(cycleDate);
								}
								columnsData.add(Config.dtFormat.format(vdate));
								for (templateIndex = 1; templateIndex < templateRowCount; templateIndex++) {
									layoutRow = layoutSheet.getRow(templateIndex);
									columnName = WebHelperUtil.getCellData("ColumnName", layoutSheet, templateIndex,
											inputHashTable);
									colNo = WebHelperUtil.getCellData("Column", layoutSheet, templateIndex,
											inputHashTable);
									colNoInt = Integer.parseInt(colNo);

									if (rIndex == Integer.parseInt(startRow.toString())) {
										columns.add(columnName.toString());
									}

									strXPath = tableID + "/tbody/tr[" + (rIndex + 1) + "]/td[" + (colNoInt + 1) + "]";
									String sVal = Automation.driver.findElement(By.xpath(strXPath)).getText();
									columnsData.add(sVal);

								}

								rows.add(columnsData);
							}
							break;
						} else if (tableType.toString().equalsIgnoreCase("ControlNames")) {
							log.info("Inside Control Names block");
							isFromVerification = true;
							columns.add("TestCaseID");
							columns.add("TransactionType");
							columns.add("CycleDate");

							if (isClaimsApplication) {
								columns.add("Header");
							}
							if (!(isBillingApplication)) {
								columns.add("CurrentDate");
							}
							columnsData = new ArrayList<String>();
							columnsData.add(testcaseID.toString());
							columnsData.add(vTransaction.toString());

							if (isClaimsApplication) {
								columnsData.add(MainControllerClaims.mainControllerHeaderNo);
							}

							if (!(isBillingApplication)) {
								columnsData.add(Config.dtFormat.format(vdate));
							}

							if (isBillingApplication) {
								columnsData.add(cycleDate);
							}

							for (templateIndex = 1; templateIndex < templateRowCount; templateIndex++) {
								layoutRow = layoutSheet.getRow(templateIndex);

								columnName = WebHelperUtil.getCellData("ColumnName", layoutSheet, templateIndex,
										inputHashTable);
								controlType = WebHelperUtil.getCellData("ControlType", layoutSheet, templateIndex,
										inputHashTable);
								controlName = WebHelperUtil.getCellData("ControlName", layoutSheet, templateIndex,
										inputHashTable);
								controlID = WebHelperUtil.getCellData("ControlID", layoutSheet, templateIndex,
										inputHashTable);
								tableElement = getElementByType(controlID.toString(), controlName.toString(), "");
								columns.add(columnName.toString());
								String sVal = null;
								if (isClaimsApplication) {
									sVal = WebHelperClaims.doAction("", vRow, "", "", controlType.toString(),
											controlID.toString(), controlName.toString(), "Verification", "", "", "",
											"", "V", tableElement, false, null, null, 0, 0, "", "", operationtype, "",
											"");
								} else if (ITAFWebDriver.isBillingApplication() || (ITAFWebDriver.isSuiteApplication()
										&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Billing"))) {
									sVal = WebHelperBilling.doAction("", vRow, "", "", controlType.toString(),
											controlID.toString(), controlName.toString(), "Verification", "", "", "",
											"", "V", tableElement, false, null, null, 0, 0, "", "", operationtype, "",
											"");
								}

								columnsData.add(sVal);

							}
							rows.add(columnsData);

						}
					}

					actualSheet = WebVerification.createActualSheet(vTransaction.toString(), columns, rows, ActualPath,
							SheetName, cycleDate, testcaseID.toString(), duplicateActualPath);

					File expectedFile = new File(expectedSheetPath);
					if (expectedFile.exists()) {
						Sheet expectedSht = WebHelperUtil.getSheet(expectedSheetPath, "Expected");
						log.info(actualSheet + "," + expectedSht + "," + columns + "," + columnsData + ","
								+ testcaseID.toString() + "," + vTransaction.toString() + "," + operationtype + ","
								+ cycleDate);

						if (isClaimsApplication) {// Need to check with PAS
							report = ExcelUtility.CompareExcel(actualSheet, expectedSht, columns, columnsData,
									testcaseID.toString(), vTransaction.toString(), operationtype, cycleDate);

						} else if (ITAFWebDriver.isBillingApplication() || (ITAFWebDriver.isSuiteApplication()
								&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Billing"))) {
							report = ExcelUtility.CompareExcelBilling(actualSheet, expectedSht, columns, columnsData,
									testcaseID.toString(), vTransaction.toString(), cycleDate);
						}
						report.setReport(report);
					} else {
						report.setStatus("PASS");
						report.setMessage("Expected Sheet not found| Actual Sheet created");
					}

				} else if (!vTransaction.equalsIgnoreCase(transactionType) && rowIndex == rowCount - 1) {
					controller.pauseFun("Transaction " + transactionType + " not Found");
				}

			} catch (Exception e) {
				log.error(e.getMessage(), e);
				report.setStatus("FAIL");
				if (isClaimsApplication) {
					report.setCycleDate(controller.cycleDateCellValue);
				} else if (ITAFWebDriver.isBillingApplication() || (ITAFWebDriver.isSuiteApplication()
						&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Billing"))) {
					dt = new Date();
					report.setToDate(Config.dtFormat.format(dt));
					if (ITAFWebDriver.isClaimsApplication() || (ITAFWebDriver.isSuiteApplication()
							&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims")))
						ExcelUtility.writeReportPAS(report);
					else
						ExcelUtility.writeReport(report);
					controller.pauseFun(e.getMessage());
				}
			} finally {
				columns.clear();
				columnsData.clear();
				rows.clear();
			}
		}
		dt = new Date();

		report.getReport();
		report.setTestDescription(controller.testDescription);
		report.setToDate(Config.dtFormat.format(dt));

		// Below given change done for Billing UI verification
		if ((ITAFWebDriver.isBillingApplication() || (ITAFWebDriver.isSuiteApplication()
				&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Billing")))
				&& (operationtype.equalsIgnoreCase("Capture") || operationtype.equalsIgnoreCase("InputandVerify"))) {
			report.setTrasactionType(controller.controllerTransactionType.toString());
			report.setVerificationType(transactionType);
			report.setCycleDate(cycleDate);
			report.setStatus(report.getStatus());
			report.setMessage(report.getMessage());
			// report.setTrasactionType(transactionType);

		}

		if (isClaimsApplication) {
			if (report.getStatus().equalsIgnoreCase("FAIL")) {
				report.setMessage(failedMsg);
			}
			report.setCycleDate(controller.cycleDateCellValue);
		}
		if (StringUtils.isBlank(report.getMessage()))
			report.setMessage("See Detailed Results");
		//
		isTableFound = false;
		if (ITAFWebDriver.isClaimsApplication() || (ITAFWebDriver.isSuiteApplication()
				&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims"))) {
			ExcelUtility.writeReportPAS(report);
		} else {
			ExcelUtility.writeReport(report);
		}

	}

	public static void performVerification(String transactionType, String testcaseID) throws IOException, Exception {

		if (ITAFWebDriver.isSuiteApplication() && MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims"))
			isClaimsApplication = true;
		if (ITAFWebDriver.isSuiteApplication() && MainControllerSuite.transactionSilo.equalsIgnoreCase("Billing"))
			isBillingApplication = true;

		Sheet vTableSheet = ExcelUtility.GetSheet(Config.verificationTableISTPath, "VerificationTables");
		int rowCount = vTableSheet.getLastRowNum() + 1;
		vTableListMap = WebHelperUtil.getValueFromHashMap(vTableSheet);
		Reporter report = new Reporter();
		Sheet actualSheet = null;
		WebElement tableElement = null;
		List<WebElement> rowElements = null;
		List<WebElement> columnElements = null;
		String ActualPath = null;
		String expectedSheetPath = null;
		String SheetName = "ActualValues";
		vdate = new Date();
		WebHelper.frmDate = new Date();
		report.setFromDate(Config.dtFormat.format(vdate));
		boolean reportWrittenAfterException = false;
		for (int rowIndex = 1; rowIndex < rowCount && !controller.pauseExecution; rowIndex++) {
			try {
				reportWrittenAfterException = false;
				Row vRow = vTableSheet.getRow(rowIndex);
				String executeFlag = WebHelperUtil.getCellData("Verify", vTableSheet, rowIndex, inputHashTable);
				String vTransaction = WebHelperUtil.getCellData("TransactionType", vTableSheet, rowIndex,
						inputHashTable);
				if (isTableFound == true && !vTransaction.equalsIgnoreCase(transactionType)) {
					break;
				}

				if (executeFlag.toString().equalsIgnoreCase("Y")
						&& vTransaction.toString().equalsIgnoreCase(transactionType.toString())) {
					isTableFound = true;
					String functionalFlag = WebHelperUtil.getCellData("Functional", vTableSheet, rowIndex,
							inputHashTable);
					String templateDir = WebHelperUtil.getCellData("TemplateAdditionalPath", vTableSheet, rowIndex,
							inputHashTable);
					String templateSheet = WebHelperUtil.getCellData("TemplateSheet", vTableSheet, rowIndex,
							inputHashTable);
					String expectedDirPath = WebHelperUtil.getCellData("ExpectedDataAdditionalPath", vTableSheet,
							rowIndex, inputHashTable);
					String expectedSheet = WebHelperUtil.getCellData("Expected", vTableSheet, rowIndex, inputHashTable);
					String templatePath = Config.verificationTemplatePath + templateDir.toString() + "\\"
							+ templateSheet.toString();
					log.info(templatePath);
					ActualPath = Config.expectedValuesValuesPath + expectedDirPath + "\\" + transactionType
							+ "_Actual.xls";
					// TM:16-01-2015
					expectedSheetPath = Config.expectedValuesValuesPath + expectedDirPath.toString() + "\\"
							+ expectedSheet.toString();
					Sheet layoutSheet = ExcelUtility.GetSheet(templatePath, "Layout");
					int templateRowCount = layoutSheet.getLastRowNum() + 1;
					templateMap = WebHelperUtil.getValueFromHashMap(layoutSheet);

					for (int templateIndex = 1; templateIndex < templateRowCount
							&& !controller.pauseExecution; templateIndex++) {
						Row layoutRow = layoutSheet.getRow(templateIndex);
						String tableID = WebHelperUtil.getCellData("TableID", layoutSheet, templateIndex,
								inputHashTable);
						String tableType = WebHelperUtil.getCellData("TableType", layoutSheet, templateIndex,
								inputHashTable);
						String tableIDType = WebHelperUtil.getCellData("TableIDType", layoutSheet, templateIndex,
								inputHashTable);
						String startRow = WebHelperUtil.getCellData("StartRow", layoutSheet, templateIndex,
								inputHashTable);
						String endRow = WebHelperUtil.getCellData("EndRow", layoutSheet, templateIndex, inputHashTable);
						String columnName = WebHelperUtil.getCellData("ColumnName", layoutSheet, templateIndex,
								inputHashTable);
						String rowNo = WebHelperUtil.getCellData("Row", layoutSheet, templateIndex, inputHashTable);
						String colNo = WebHelperUtil.getCellData("Column", layoutSheet, templateIndex, inputHashTable);
						String controlName = WebHelperUtil.getCellData("ControlName", layoutSheet, templateIndex,
								inputHashTable);
						String controlType = WebHelperUtil.getCellData("ControlType", layoutSheet, templateIndex,
								inputHashTable);
						String controlID = WebHelperUtil.getCellData("ControlID", layoutSheet, templateIndex,
								inputHashTable);
						report.setTestcaseId(testcaseID.toString());
						report.setTrasactionType(vTransaction.toString());
						report.setTestDescription(controller.testDescription);
						report.setFromDate(Config.dtFormat.format(vdate));
						report.setIteration(Config.cycleNumber);
						report.setGroupName(controller.controllerGroupName.toString());
						report.setStatus("PASS");
						report.setMessage(" ");

						if (tableType.equalsIgnoreCase("NonUniform")) {
							columns.add("TestCaseID");
							columns.add("TransactionType");
							columns.add("CurrentDate");

							columnsData = new ArrayList<String>();
							columnsData.add(testcaseID.toString());
							columnsData.add(vTransaction.toString());
							columnsData.add(Config.dtFormat.format(vdate));
							int rowNoInt;
							int colNoInt;
							String strXPath;

							for (templateIndex = 1; templateIndex < templateRowCount; templateIndex++) {
								layoutRow = layoutSheet.getRow(templateIndex);
								columnName = WebHelperUtil.getCellData("ColumnName", layoutSheet, templateIndex,
										inputHashTable);
								columns.add(columnName.toString());

								controlName = WebHelperUtil.getCellData("ControlName", layoutSheet, templateIndex,
										inputHashTable);
								tableID = WebHelperUtil.getCellData("TableID", layoutSheet, templateIndex,
										inputHashTable);
								tableType = WebHelperUtil.getCellData("TableType", layoutSheet, templateIndex,
										inputHashTable);

								if (tableType.equalsIgnoreCase("NonTbody")) {

									strXPath = tableID + controlName;

									try {
										if (Automation.driver.findElement(By.xpath(strXPath)).isDisplayed()) {

											String sVal = Automation.driver.findElement(By.xpath(strXPath)).getText();

											columnsData.add(sVal);

										}

									} catch (Exception e) {
										log.error(e.getMessage(), e);
										String sVal = null;

										columnsData.add(sVal);

									}
								} else {

									rowNo = WebHelperUtil.getCellData("Row", layoutSheet, templateIndex,
											inputHashTable);
									colNo = WebHelperUtil.getCellData("Column", layoutSheet, templateIndex,
											inputHashTable);
									rowNoInt = Integer.parseInt(rowNo);
									colNoInt = Integer.parseInt(colNo);
									strXPath = tableID + "/tbody/tr[" + (rowNoInt + 1) + "]/td[" + (colNoInt + 1) + "]"
											+ controlName;
									try {
										if (Automation.driver.findElement(By.xpath(strXPath)).isDisplayed()) {

											String sVal = Automation.driver.findElement(By.xpath(strXPath)).getText();

											columnsData.add(sVal);

										}

									} catch (Exception e) {
										log.error(e.getMessage(), e);
										String sVal = null;

										columnsData.add(sVal);

									}
								}
							}
							rows.add(columnsData);

						}

						if (tableType.equalsIgnoreCase("RateFactor")) {
							int i;
							columns.add("TestCaseID");
							columns.add("TransactionType");
							columns.add("CurrentDate");
							columnsData = new ArrayList<String>();
							columnsData.add(testcaseID.toString());
							columnsData.add(vTransaction.toString());
							columnsData.add(Config.dtFormat.format(vdate));
							String strXPath;
							String sVal;

							for (templateIndex = 1; templateIndex < templateRowCount; templateIndex++) {
								layoutRow = layoutSheet.getRow(templateIndex);
								columnName = WebHelperUtil.getCellData("ColumnName", layoutSheet, templateIndex,
										inputHashTable);
								columns.add(columnName.toString());

								controlType = WebHelperUtil.getCellData("ControlType", layoutSheet, templateIndex,
										inputHashTable);
								tableID = WebHelperUtil.getCellData("TableID", layoutSheet, templateIndex,
										inputHashTable);

								strXPath = tableID + controlName;
								try {
									if (Automation.driver.findElement(By.xpath(strXPath)).isDisplayed()) {

										if (controlType.equalsIgnoreCase("WebEdit")) {
											sVal = Automation.driver.findElement(By.xpath(strXPath))
													.getAttribute("value");
										} else {
											sVal = Automation.driver.findElement(By.xpath(strXPath)).getText();
										}
										columnsData.add(sVal);

									}

								} catch (Exception e) {
									log.error(e.getMessage(), e);
									sVal = null;

									columnsData.add(sVal);

								}

							}
							rows.add(columnsData);
							// }
						}

						if (tableType.equalsIgnoreCase("NonTable")) {
							columns.add("TestCaseID");
							columns.add("TransactionType");
							columns.add("CurrentDate");
							columnsData = new ArrayList<String>();
							columnsData.add(testcaseID.toString());
							columnsData.add(vTransaction.toString());
							columnsData.add(Config.dtFormat.format(vdate));
							String strXPath;
							String sVal;

							for (templateIndex = 1; templateIndex < templateRowCount; templateIndex++) {
								layoutRow = layoutSheet.getRow(templateIndex);
								columnName = WebHelperUtil.getCellData("ColumnName", layoutSheet, templateIndex,
										inputHashTable);
								columns.add(columnName.toString());

								controlType = WebHelperUtil.getCellData("ControlType", layoutSheet, templateIndex,
										inputHashTable);
								tableID = WebHelperUtil.getCellData("TableID", layoutSheet, templateIndex,
										inputHashTable);

								strXPath = tableID + controlName;
								try {
									if (Automation.driver.findElement(By.xpath(strXPath)).isDisplayed()) {

										if (controlType.equalsIgnoreCase("WebEdit")) {
											sVal = Automation.driver.findElement(By.xpath(strXPath))
													.getAttribute("value");
										} else {
											sVal = Automation.driver.findElement(By.xpath(strXPath)).getText();
										}
										columnsData.add(sVal);

									} else {

										sVal = null;

										columnsData.add(sVal);

									}

								} catch (Exception e) {
									log.error(e.getMessage(), e);
									sVal = null;

									columnsData.add(sVal);

								}

							}
							rows.add(columnsData);

						} else if (tableType.equalsIgnoreCase("DB")) {
							if (tableIDType.equalsIgnoreCase("Expected")) {
								SheetName = "Expected";
								ActualPath = Config.expectedValuesValuesPath + expectedDirPath + "\\" + transactionType
										+ "_Expected.xls";
							}
							Connection conn = null;
							ResultSet rs = null;// =
							// JDBCConnection.establishDBConn(tableID);
							Statement st = null;
							if ((ITAFWebDriver.isDMApplication() || (ITAFWebDriver.isSuiteApplication()
									&& MainControllerSuite.transactionSilo.equalsIgnoreCase("DM")))
									|| tableIDType.equalsIgnoreCase("Dynamic")) {

								tableID = DMProduct.getDynamicSQLQuery(tableID, controlName);
								log.info("tableID : " + tableID);
								conn = JDBCConnection.establishPASDBConn();
								st = conn.createStatement();
								rs = st.executeQuery(tableID);
							} else {
								conn = JDBCConnection.establishDBConn();
								st = conn.createStatement();
								rs = st.executeQuery(tableID);
							}

							isTableFound = true;
							columns.add("TestCaseID");
							columns.add("TransactionType");
							columns.add("CurrentDate");

							for (int rIndex = Integer.parseInt(startRow.toString()); rs.next(); rIndex++)// rowElements.size()
							{
								columnsData = new ArrayList<String>();
								columnsData.add(testcaseID.toString());
								columnsData.add(vTransaction.toString());
								columnsData.add(Config.dtFormat.format(vdate));

								for (templateIndex = 1; templateIndex < templateRowCount; templateIndex++) {
									layoutRow = layoutSheet.getRow(templateIndex);
									columnName = WebHelperUtil.getCellData("ColumnName", layoutSheet, templateIndex,
											inputHashTable);
									colNo = WebHelperUtil.getCellData("Column", layoutSheet, templateIndex,
											inputHashTable);

									if (rIndex == Integer.parseInt(startRow.toString())) {
										columns.add(columnName.toString());
									}
									columnsData.add(rs.getString(Integer.parseInt(colNo)));
								}
								rows.add(columnsData);
							}
							rs.close();
							st.close();
							JDBCConnection.closeConnection(conn);
						} else if (tableType.equalsIgnoreCase("UniformDynamic"))

						{
							int rCount = 0;
							columns.add("TestCaseID");
							columns.add("TransactionType");
							columns.add("CurrentDate");
							columnsData = new ArrayList<String>();
							columnsData.add(testcaseID.toString());
							columnsData.add(vTransaction.toString());
							columnsData.add(Config.dtFormat.format(vdate));

							int colNoInt;
							String strXPath;

							for (templateIndex = 1; templateIndex < templateRowCount; templateIndex++)// column-wise
							{
								rowNo = WebHelperUtil.getCellData("Row", layoutSheet, templateIndex, inputHashTable);

								int rIndex = Integer.parseInt(rowNo.toString());
								layoutRow = layoutSheet.getRow(templateIndex);
								columnName = WebHelperUtil.getCellData("ColumnName", layoutSheet, templateIndex,
										inputHashTable);
								colNo = WebHelperUtil.getCellData("Column", layoutSheet, templateIndex, inputHashTable);
								controlType = WebHelperUtil.getCellData("ControlType", layoutSheet, templateIndex,
										inputHashTable);
								colNoInt = Integer.parseInt(colNo);
								controlName = WebHelperUtil.getCellData("ControlName", layoutSheet, templateIndex,
										inputHashTable);
								tableID = WebHelperUtil.getCellData("TableID", layoutSheet, templateIndex,
										inputHashTable);
								columns.add(columnName.toString());

								if (controlType == "" || controlType.equalsIgnoreCase(null)) {
									strXPath = tableID + "/tbody/tr[" + (rIndex + 1) + "]/td[" + (colNoInt + 1) + "]";
									String sVal = Automation.driver.findElement(By.xpath(strXPath)).getText();
									columnsData.add(sVal);

								} else {
									strXPath = tableID + "/tbody/tr[" + (rIndex + 1) + "]/td[" + (colNoInt + 1) + "]"
											+ controlName;
									try {
										if (Automation.driver.findElement(By.xpath(strXPath)).isDisplayed()) {
											String sVal = Automation.driver.findElement(By.xpath(strXPath)).getText();
											columnsData.add(sVal);

										} else {
											String sVal = null;
											columnsData.add(sVal);
										}
									} catch (Exception e) {
										log.error(e.getMessage(), e);
										String sVal = null;
										columnsData.add(sVal);
									}
								}
								rows.add(columnsData);
							}
							break;
						} else if (tableType.equalsIgnoreCase("Uniform")) {
							columns.add("TestCaseID");
							columns.add("TransactionType");
							columns.add("CurrentDate");
							int colNoInt;
							String strXPath;

							if (ITAFWebDriver.isPASApplication() || (ITAFWebDriver.isSuiteApplication()
									&& MainControllerSuite.transactionSilo.equalsIgnoreCase("PAS"))) {
								tableElement = getElementByType(tableIDType, tableID, "verification");
								rowElements = tableElement.findElements(By.xpath(tableID + "/tbody/tr"));

								int rCount = 0;

								if (endRow.equalsIgnoreCase("0") || endRow.equalsIgnoreCase("")) {
									rCount = rowElements.size();
								}

								else {
									rCount = Integer.parseInt(endRow);
									if (rCount > rowElements.size()) {
										rCount = rowElements.size();
									}
								}

							}

							else {
								if (!Config.seleniumExecution.equalsIgnoreCase("RC")) {
									tableElement = getElementByType(tableIDType, tableID, "verification");
									rowElements = tableElement.findElements(By.tagName("tr"));
								}
							}
							for (int rIndex = Integer.parseInt(startRow.toString()); rIndex < rowElements
									.size(); rIndex++) {
								columnsData = new ArrayList<String>();
								columnsData.add(testcaseID.toString());
								columnsData.add(vTransaction.toString());
								columnsData.add(Config.dtFormat.format(vdate));
								for (templateIndex = 1; templateIndex < templateRowCount; templateIndex++) {
									layoutRow = layoutSheet.getRow(templateIndex);
									columnName = WebHelperUtil.getCellData("ColumnName", layoutSheet, templateIndex,
											inputHashTable);
									colNo = WebHelperUtil.getCellData("Column", layoutSheet, templateIndex,
											inputHashTable);
									colNoInt = Integer.parseInt(colNo);

									if (rIndex == Integer.parseInt(startRow.toString())) {
										columns.add(columnName.toString());
									}
									if (ITAFWebDriver.isDMApplication() || (ITAFWebDriver.isSuiteApplication()
											&& MainControllerSuite.transactionSilo.equalsIgnoreCase("DM"))) {
										strXPath = tableID + "/tr[" + (rIndex + 1) + "]/td[" + (colNoInt + 1) + "]";
										String sVal = Automation.driver.findElement(By.xpath(strXPath)).getText();
										columnsData.add(sVal);
									} else {
										try {
											if (controlType == "" || controlType.equalsIgnoreCase(null)) {
												strXPath = tableID + "/tbody/tr[" + (rIndex + 1) + "]/td["
														+ (colNoInt + 1) + "]";
												String sVal = Automation.driver.findElement(By.xpath(strXPath))
														.getText();
												columnsData.add(sVal);
											} else {
												strXPath = tableID + "/tbody/tr[" + (rIndex + 1) + "]/td["
														+ (colNoInt + 1) + "]/" + controlName;
												String sVal = Automation.driver.findElement(By.xpath(strXPath))
														.getAttribute("value");
												columnsData.add(sVal);
											}

										} catch (Exception e) {

											log.error(e.getMessage() + "WebElement not present");
											Thread.sleep(500);
										}

									}
								}

								rows.add(columnsData);
							}
							break;
						} else if (tableType.toString().equalsIgnoreCase("ControlNames")) {
							isFromVerification = true;
							columns.add("TestCaseID");
							columns.add("TransactionType");
							columns.add("CurrentDate");
							columnsData = new ArrayList<String>();
							columnsData.add(testcaseID.toString());
							columnsData.add(vTransaction.toString());
							columnsData.add(Config.dtFormat.format(vdate));

							for (templateIndex = 1; templateIndex < templateRowCount; templateIndex++) {
								layoutRow = layoutSheet.getRow(templateIndex);

								columnName = WebHelperUtil.getCellData("ColumnName", layoutSheet, templateIndex,
										inputHashTable);
								controlType = WebHelperUtil.getCellData("ControlType", layoutSheet, templateIndex,
										inputHashTable);
								controlName = WebHelperUtil.getCellData("ControlName", layoutSheet, templateIndex,
										inputHashTable);
								controlID = WebHelperUtil.getCellData("ControlID", layoutSheet, templateIndex,
										inputHashTable);
								tableElement = getElementByType(controlID.toString(), controlName.toString(), "");
								columns.add(columnName.toString());
								String sVal = WebHelperPAS.doAction("", controlType.toString(), controlID.toString(),
										controlName.toString(), "Verification", "", "V", tableElement, false, null,
										null, 0, 0, "", "");
								columnsData.add(sVal);

							}
							rows.add(columnsData);

						} else if (tableType.toString().equalsIgnoreCase("ControlNames_DM")) {
							isFromVerification = true;
							columns.add("TestCaseID");
							columns.add("TransactionType");
							columns.add("CurrentDate");

							int colNoInt;
							String strXPath;

							if (!Config.seleniumExecution.equalsIgnoreCase("RC")) {
								tableElement = getElementByType(tableIDType, tableID, "verification");
								rowElements = tableElement.findElements(By.tagName("tr"));
							}

							for (int rIndex = Integer.parseInt(startRow.toString()); rIndex < rowElements
									.size(); rIndex++) {

								columnsData = new ArrayList<String>();
								columnsData.add(testcaseID.toString());
								columnsData.add(vTransaction.toString());
								columnsData.add(Config.dtFormat.format(vdate));

								for (templateIndex = 1; templateIndex < templateRowCount; templateIndex++) {
									layoutRow = layoutSheet.getRow(templateIndex);

									columnName = WebHelperUtil.getCellData("ColumnName", layoutSheet, templateIndex,
											inputHashTable);
									controlType = WebHelperUtil.getCellData("ControlType", layoutSheet, templateIndex,
											inputHashTable);
									controlName = WebHelperUtil.getCellData("ControlName", layoutSheet, templateIndex,
											inputHashTable);
									controlID = WebHelperUtil.getCellData("ControlID", layoutSheet, templateIndex,
											inputHashTable);
									String[] arrTempcontrolName = controlName.split("\\$");
									controlName = arrTempcontrolName[0].toString() + (rIndex + 1)
											+ arrTempcontrolName[1].toString();

									tableElement = getElementByType(controlID.toString(), controlName.toString(),
											"verification");

									if (rIndex == Integer.parseInt(startRow.toString())) {
										columns.add(columnName.toString());
									}

									String sVal = "";
									if (tableElement == null) {
										sVal = "Element not found with " + controlID.toString() + " as '"
												+ controlName.toString() + "'";
									} else {
										sVal = WebHelperPAS.doAction("", controlType.toString(), controlID.toString(),
												controlName.toString(), "Verification", "", "V", tableElement, false,
												null, null, 0, 0, "", "");
									}
									columnsData.add(sVal);
								}
								rows.add(columnsData);

							}
						}

					}

					actualSheet = createActualSheet(vTransaction.toString(), columns, rows, ActualPath, SheetName);
					File expectedFile = new File(expectedSheetPath);
					if (expectedFile.exists()) {
						Sheet expectedSht = ExcelUtility.GetSheet(expectedSheetPath, "Expected");

						/*
						 * report = ExcelUtility.compareExcel(actualSheet, expectedSht, columns,
						 * columnsData, testcaseID.toString(), vTransaction.toString());
						 */
						report = ExcelUtility.compareExcel(actualSheet, expectedSht, columns, columnsData,
								testcaseID.toString(), vTransaction.toString());

						report.setReport(report);

					} else {
						report.setStatus("FAIL");
						report.setMessage("Expected Sheet not found| Actual Sheet created");
					}
					// }

				} else if (!vTransaction.equalsIgnoreCase(transactionType) && rowIndex == rowCount - 1) {
					controller.pauseFun("Transaction " + transactionType + " not Found");
				}

			} catch (Exception e) {
				dt = new Date();
				report.setToDate(Config.dtFormat.format(dt));
				report.setStatus("FAIL");
				// log.error(e.getMessage(), e);
				String errorMessage = e.getMessage();
				String firstLine = errorMessage.split("\\R")[0];
				log.error("Error :" + firstLine);
				// ExcelUtility.writeReportPAS(report);
				controller.pauseFun("Issue in WebVerification");
				reportWrittenAfterException = true;
			} finally {
				columns.clear();
				columnsData.clear();
				rows.clear();
			}
		}

		if (!reportWrittenAfterException) {
			dt = new Date();

			report.getReport();
			report.setGroupName(controller.controllerGroupName.toString());
			report.setTestDescription(controller.testDescription.toString());
			report.setTestcaseId(controller.controllerTestCaseID.toString());
			report.setTrasactionType(controller.controllerTransactionType.toString());
			report.setFromDate(Config.dtFormat.format(vdate));
			report.setToDate(Config.dtFormat.format(dt));

			// DW: 13-Nov-2019 changes made to add the following message only if Blank and
			if (StringUtils.isBlank(report.getMessage()) && (report.getStatus()).equalsIgnoreCase("FAIL")) {
				report.setMessage("Check Detailed Results");
			}
			isTableFound = false;
			ExcelUtility.writeReportPAS(report);
		}
	}

	public static Row createHeader(Sheet actualSheet, List<String> columns) {
		Row actualRow = actualSheet.getRow(0);
		Cell testCaseID = actualRow.createCell(0);
		Cell transactionType = actualRow.createCell(1);
		Cell Date = actualRow.createCell(2);
		Iterator<String> iterator = columns.iterator();
		int count = 0;
		if (iterator.hasNext() == true) {
			count = count + 1;
			Cell dynamicColumns = actualRow.createCell(count);
		}
		return actualRow;
	}
//DM
	/*
	 * public static Sheet createActualSheet(String transactionType, List<String>
	 * columns, List<List<String>> columnData, String actualPath, String sheetName)
	 * throws IOException { FileOutputStream out = null; FileInputStream in = null;
	 * Workbook workBook = null; Sheet workSheet = null;
	 * 
	 * try { File file = new File(actualPath); boolean isExpectedSheet =
	 * sheetName.equalsIgnoreCase("Expected");
	 * 
	 * if (file.exists() && isExpectedSheet) { file.delete();
	 * log.info("Deleted Existing Expected Sheet"); }
	 * 
	 * in = new FileInputStream(file); POIFSFileSystem lPOIfs = new
	 * POIFSFileSystem(in); workBook = new HSSFWorkbook(lPOIfs);
	 * 
	 * // Check if the sheet already exists boolean sheetExists = false; for (int i
	 * = 0; i < workBook.getNumberOfSheets(); i++) { if
	 * (workBook.getSheetName(i).equalsIgnoreCase(sheetName)) { sheetExists = true;
	 * workSheet = workBook.getSheetAt(i); break; } }
	 * 
	 * if (!sheetExists) { workSheet = workBook.createSheet(sheetName);
	 * 
	 * // Write column headers only if the sheet is newly created Row headerRow =
	 * workSheet.createRow(0); for (int i = 0; i < columns.size(); i++) { Cell cell
	 * = headerRow.createCell(i); cell.setCellValue(columns.get(i)); } }
	 * 
	 * // Find the index of the last row with data int lastRowNum =
	 * workSheet.getLastRowNum();
	 * 
	 * // Set rowIndex to the next available row int rowIndex = lastRowNum + 1;
	 * 
	 * // If the sheet is empty, set rowIndex to 0 if (lastRowNum == -1) { rowIndex
	 * = 0; }
	 * 
	 * // Write data for (List<String> rowData : columnData) { Row dataRow =
	 * workSheet.createRow(rowIndex); for (int i = 0; i < rowData.size(); i++) {
	 * Cell cell = dataRow.createCell(i); cell.setCellValue(rowData.get(i)); } //
	 * Increment rowIndex for the next row rowIndex++; }
	 * 
	 * out = new FileOutputStream(actualPath); workBook.write(out); } finally { if
	 * (out != null) { out.close(); } if (in != null) { in.close(); } if (workBook
	 * != null) { workBook.close(); } }
	 * 
	 * return workSheet; }
	 */
	
	public static Sheet createActualSheet(String transactionType, List<String> columns, List<List<String>> columnData,
			String actualPath, String SheetName) throws IOException, Exception {
		FileOutputStream out = null;
		FileOutputStream out1 = null;
		POIFSFileSystem lPOIfs = null;
		InputStream in = null;
		Workbook workBook = null;
		Sheet workSheet = null;
		File lFile = new File(actualPath);
		if (lFile.exists() && SheetName.equalsIgnoreCase("Expected")) {
			lFile.delete();
			System.out.println("Deleted Existing Expected Sheet");
		}
		int columnSize = columns.size();
		int columnIndex = 0;
		Cell actualCell = null;
		Cell valuesCell = null;
		if (!lFile.exists()) {
			try {
				workBook = new HSSFWorkbook();
				workSheet = workBook.createSheet(SheetName);
				Row actualRowHeader = workSheet.createRow(0);
				int rowNum = 0;
				Row valuesRow = workSheet.createRow(rowNum + 1);
				currentRowIndex = rowNum;
				// Iterator<String> iterator = columns.iterator();
				int rowCount = rows.size();
				for (int rIndex = 0; rIndex < rowCount; rIndex++) // changed by
				{
					if (rIndex > 0) {
						valuesRow = workSheet.createRow(rowNum + 1 + rIndex);
					}
					columnSize = rows.get(rIndex).size();
					for (columnIndex = columnSize; columnIndex >= 1; columnIndex--) {
						actualCell = actualRowHeader.createCell(columnSize - columnIndex);
						valuesCell = valuesRow.createCell(columnSize - columnIndex);
						actualCell.setCellValue(columns.get(columnSize - columnIndex));
						valuesCell.setCellValue(rows.get(rIndex).get(columnSize - columnIndex));
						System.out.print(valuesCell.toString());
					}
				}
				out = new FileOutputStream(actualPath);
				workBook.write(out);
			} catch (IOException ioe) {
				ioe.getLocalizedMessage();
			} finally {
				out.flush();
				out.close();
				workBook.close();
			}
		} else {
			try {
				in = new FileInputStream(actualPath);
				lPOIfs = new POIFSFileSystem(in);
				workBook = new HSSFWorkbook(lPOIfs);
				workSheet = workBook.getSheet(SheetName);
				int lastRow = workSheet.getLastRowNum();
				currentRowIndex = lastRow;
				Row row = workSheet.getRow(lastRow + 1);
				if (row == null) {
					row = workSheet.createRow(lastRow + 1);
				}
				int rowCount = rows.size();
				for (int rIndex = 0; rIndex < rowCount; rIndex++) {
					if (rIndex > 0) {
						row = workSheet.createRow(lastRow + 1 + rIndex);
					}
					int cSize = rows.get(rIndex).size();

					/*
					 * for (columnIndex = cSize; columnIndex >= 1; columnIndex--) { valuesCell =
					 * row.getCell(cSize - columnIndex); if (valuesCell == null) { valuesCell =
					 * row.createCell(cSize - columnIndex);
					 * valuesCell.setCellType(Cell.CELL_TYPE_STRING);
					 * valuesCell.setCellValue(rows.get(rIndex).get(cSize - columnIndex)); } }
					 */
					for (columnIndex = cSize; columnIndex >= 1; columnIndex--) {
					    valuesCell = row.getCell(cSize - columnIndex);
					    if (valuesCell == null) {
					        valuesCell = row.createCell(cSize - columnIndex);
					    }
					    valuesCell.setCellValue(rows.get(rIndex).get(cSize - columnIndex));
					}
					
			
				}
				out1 = new FileOutputStream(actualPath);
				workBook.write(out1);
			} catch (Exception ioe) {
				controller.pauseFun(ioe.getLocalizedMessage() + " from CreateActualSheet Function ");
			} finally {
				out1.flush();
				out1.close();
				// System.out.println("HELLO WORLD");
			}
		}
		return workSheet;
	}



	public static Sheet createActualSheet(String transactionType, List<String> columns, List<List<String>> columnData,
			String actualPath, String sheetName, String cycleDate, String testcaseID, String duplicateActualPath)
			throws IOException {
		FileOutputStream out = null;
		FileInputStream in = null;
		Workbook workBook = null;
		Sheet workSheet = null;

		try {
			File file = new File(actualPath);
			boolean isActualSheet = sheetName.equalsIgnoreCase("ActualValues");

			if (file.exists() && isActualSheet) {
				file.delete();
				log.info("Deleted Existing Actual Sheet");
			}

			in = new FileInputStream(file);
			workBook = new XSSFWorkbook(in);
			workSheet = workBook.createSheet(sheetName);

			int rowIndex = 0;
			Row headerRow = workSheet.createRow(rowIndex++);

			// Write column headers
			for (int i = 0; i < columns.size(); i++) {
				Cell cell = headerRow.createCell(i);
				cell.setCellValue(columns.get(i));
			}

			// Write data
			for (List<String> rowData : columnData) {
				Row dataRow = workSheet.createRow(rowIndex++);
				for (int i = 0; i < rowData.size(); i++) {
					Cell cell = dataRow.createCell(i);
					cell.setCellValue(rowData.get(i));
				}
			}

			out = new FileOutputStream(actualPath);
			workBook.write(out);
		} finally {
			if (out != null) {
				out.close();
			}
			if (in != null) {
				in.close();
			}
			if (workBook != null) {
				workBook.close();
			}
		}

		return workSheet;
	}

	private static Workbook createWorkbook() throws IOException {
		return createWorkbook(null);
	}

	private static Workbook createWorkbook(POIFSFileSystem lPOIfs1) throws IOException {
		if (ITAFWebDriver.isBillingApplication() || (ITAFWebDriver.isSuiteApplication()
				&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Billing"))) {
			return new XSSFWorkbook();
		} else if (ITAFWebDriver.isClaimsApplication() || (ITAFWebDriver.isSuiteApplication()
				&& MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims"))) {
			if (lPOIfs1 == null) {
				return new HSSFWorkbook();
			} else {
				return new HSSFWorkbook(lPOIfs1);
			}
		}
		return null;
	}

	public static WebElement getElementByType(String controlID, String controlName, String purpose) throws IOException {
		WebElement element = null;
		try {
			Constants.ControlIdEnum controlId = Constants.ControlIdEnum.valueOf(controlID);
			switch (controlId) {
			case ClassName:
				element = Automation.driver.findElement(By.className(controlName));
				break;
			case Id:
			case HTMLID:
				element = Automation.driver.findElement(By.id(controlName));
				break;
			case Name:
				element = Automation.driver.findElement(By.name(controlName));
				break;
			case TagName:
				element = Automation.driver.findElement(By.tagName(controlName));
				break;
			case XPath:
				element = Automation.driver.findElement(By.xpath(controlName));
				break;
			default:
				break;
			}
		} catch (Exception e) {
			if (StringUtils.equalsIgnoreCase(purpose, "verification")) {
				return element;
			}
			failedMsg = e.getMessage();
			controller.pauseFun(e.getMessage() + " from getElementByType Function");
		}
		return element;
	}

	public static String getText(ITAFWebDriverPAS driver, WebElement element) {
		return (String) ((JavascriptExecutor) driver).executeScript("return jQuery(arguments[0]).text();", element);
	}

}
