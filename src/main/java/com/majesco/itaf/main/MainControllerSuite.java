package com.majesco.itaf.main;

import java.sql.ResultSet;
import java.sql.Statement;
//import java.sql.Timestamp;
import java.time.Duration;
import java.io.IOException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Date;
import java.util.HashMap;
import javax.swing.JOptionPane;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.vo.Reporter;

public class MainControllerSuite extends MainController {
	private final static Logger log = LogManager.getLogger(MainControllerSuite.class.getName());
	public static String claimsloader = "//div[@class='overlay']/div[@class='logo-wrapper']/div";
	private static HashMap<String, Integer> sheetValues = new HashMap<String, Integer>();
	private static HashMap<String, Integer> controlsheet = new HashMap<String, Integer>();
	private static HashMap<String, Integer> startpointerValues = new HashMap<String, Integer>();
	private static int startCol = 0;
	private static int startRow = 0;
	private static Row controllerRow = null;
	public static boolean recoverydone = false;
	public static Sheet MainControlSheet = null;
	public static Row startpointerRow = null;
	public static int START_ROW_POS = 0;
	public static int START_COLUMN_POS = 0;
	protected static ResultSet result = null;
	public static String scrsubject = "";
	public static String scrshotbody = "";
	public static String strScrCompleteData = "";
	public static String scrEmail = "";
	private static String scruniquenumbersheet = "";
	public static String scrreportsheet = "";
	public static String scrdetailreportsheet = "";
	protected static Statement st = null;
	public static String cbd_requestxml = null;
	public static String cbd_wsdl = null;
	public static Sheet startpointerSheet = null;
	public static Sheet MainControllerSheet = null;
	public static Row updatedstartpointerRow = null;
	private static HashMap<String, Integer> updatedstartpointerValues = new HashMap<String, Integer>();
	public static int UPDATED_START_ROW = 0;
	public static int UPDATED_START_COLUMN = 0;
	public static int ustartCol = 0;
	public static int ustartRow = 0;
	public static boolean columnChanged = false;
	public static String jobName = null, PolicyNo = null, brokerNumber = null, AccountNumber = null,
			brokerSystemCode = null, accountSystemCode = null, job_SEQ = null;
	public static boolean rowChanged = false;
	public static String mainControllerHeaderNo = null;
	//Mandar
	public static int DResult = 1, DVResult = 1;
	public static String transactionSilo = "", transactionVerificationType = "";
	public static Date frmDate = null;
	public static boolean reportWritten = false;
	//Mandar
	
	
	private static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();

	/**
	 * Finds the Start Pointer in the MainController Sheet and executes the
	 * Transaction
	 **/

	public Reporter ControllerData(String FilePath) throws NullPointerException, Exception {
		Reporter report = new Reporter();
		Sheet reqSheet = null;
		int execFlag = 0, rowCount = 0, colCount = 0;

		// Read MainControlSheet Tab
		try {
			reqSheet = ExcelUtility.GetSheet(Config.controllerFilePath, "MainControlSheet");
			sheetValues = WebHelperUtil.getValueFromHashMap(reqSheet);
			execFlag = Integer.parseInt(sheetValues.get("ExecuteFlag").toString());
			rowCount = reqSheet.getLastRowNum() + 1;
		} catch (Exception e) {
			log.error("Failed to Read Tab 'MainControlSheet' from MainController file <-|-> LocalizeMessage "
					+ e.getLocalizedMessage() + " <-|-> Message " + e.getMessage() + " <-|-> Cause " + e.getCause());
			throw new Exception("Failed to Read Tab 'MainControlSheet' from MainController file : " + e.getMessage());
		}

		// Read StartPointer Tab
		try {
			startpointerSheet = ExcelUtility.GetSheet(Config.controllerFilePath, "StartPointer");
			startpointerValues = WebHelperUtil.getValueFromHashMap(startpointerSheet);
			startpointerRow = startpointerSheet.getRow(1);
			START_ROW_POS = Integer.parseInt(startpointerValues.get("StartRow").toString());
			START_COLUMN_POS = Integer.parseInt(startpointerValues.get("StartCol").toString());
			startRow = (int) startpointerRow.getCell(START_ROW_POS).getNumericCellValue();
			startCol = (int) startpointerRow.getCell(START_COLUMN_POS).getNumericCellValue();
		} catch (Exception e) {
			log.error("Failed to Read Tab 'StartPointer' from MainController file <-|-> LocalizeMessage "
					+ e.getLocalizedMessage() + " <-|-> Message " + e.getMessage() + " <-|-> Cause " + e.getCause());
			throw new Exception("Failed to Read Tab 'StartPointer' from MainController file " + e.getMessage());
		}

		startRow = startRow - 1;
		startCol = startCol - 1;
		ustartCol = startCol;
		ustartRow = startRow;
		for (int rowIndex = startRow; rowIndex < rowCount; rowIndex++) {

			if (stopExecution == true) {
				return report;
			}

			pauseExecution = false;
			ustartRow = rowIndex + 1;

			reqSheet = ExcelUtility.GetSheet(Config.controllerFilePath, "MainControlSheet");

			log.info("Get sheet values");
			sheetValues = WebHelperUtil.getValueFromHashMap(reqSheet);

			try {
				controllerRow = reqSheet.getRow(rowIndex);
			} catch (Exception e) {
				controllerRow = null;
			}
			if (controllerRow == null) {
				log.info("No TestCaseID or Data present in Maincontroller at Row : " + (rowIndex + 1));
				continue;
			}

			colCount = controllerRow.getLastCellNum();
			if (rowIndex > startRow) {
				rowChanged = true;
				startCol = 6;
			}
			testDescription = WebHelperUtil.getCellData("Test_Description", reqSheet, rowIndex, sheetValues);
			Cell executeFlag = controllerRow.getCell(execFlag);
			controllerTestCaseID = controllerRow.getCell(Integer.parseInt(sheetValues.get("TestCaseID").toString()));
			controllerGroupName = controllerRow.getCell(Integer.parseInt(sheetValues.get("GroupName").toString()));
			controllerSC_No = controllerRow.getCell(Integer.valueOf((int) sheetValues.get("SC_NO").floatValue()));// 19/09/2022

			if (controllerTestCaseID == null
					|| StringUtils.equalsIgnoreCase(controllerTestCaseID.getStringCellValue(), "")) {
				log.info("No TestCaseID present in Maincontroller at Row : " + (rowIndex + 1));
				continue;
			}
			if (executeFlag != null) {
				if (executeFlag.toString().equalsIgnoreCase("Y")) {
					for (int columnIndex = startCol; columnIndex < colCount && !pauseExecution; columnIndex++) {
						try {
							controllerTransactionType = controllerRow.getCell(columnIndex);
						} catch (Exception e) {
							controllerTransactionType = null;
						}
						mainControllerHeaderNo = reqSheet.getRow(0).getCell(columnIndex).toString();
						if (controllerTransactionType != null
								&& StringUtils.isNotBlank(controllerTransactionType.getStringCellValue())) {

							// log.info("Header : " + mainControllerHeaderNo);//Sel4
							log.info("GroupName : " + controllerGroupName);
							log.info("TestCaseID : " + controllerTestCaseID);
							log.info("SC_NO : " + controllerSC_No);
							log.info("Transaction : " + controllerTransactionType);

							if (controllerTransactionType.toString().equalsIgnoreCase("PAUSE")) {
								pauseFun("Do You Wish To Continue");
							}

							else if (controllerTransactionType.toString().equalsIgnoreCase("END")
									|| controllerTransactionType.toString().equalsIgnoreCase("END1")) {
								try {
									FileInputStream in = new FileInputStream(Config.controllerFilePath);
									Workbook LinearupdatedmainWB = null;
									if (Config.controllerFilePath.endsWith(".xls")) {
										LinearupdatedmainWB = new HSSFWorkbook(in);
									} else if (Config.controllerFilePath.endsWith(".xlsx")
											| Config.controllerFilePath.endsWith(".xlsm")) {
										LinearupdatedmainWB = new XSSFWorkbook(in);
									}
									MainControllerSheet = LinearupdatedmainWB.getSheet("StartPointer");
									updatedstartpointerRow = MainControllerSheet.getRow(1);
									updatedstartpointerValues = WebHelperUtil.getValueFromHashMap(MainControllerSheet);
									UPDATED_START_ROW = updatedstartpointerValues.get("StartRow");
									UPDATED_START_COLUMN = updatedstartpointerValues.get("StartCol");
									System.out.println("START_ROW : " + ustartRow + ", START_Col : " + ustartCol);
									updatedstartpointerRow.getCell(UPDATED_START_ROW).setCellValue(ustartRow);
									ustartCol = ustartCol + 2;
									updatedstartpointerRow.getCell(UPDATED_START_COLUMN).setCellValue(ustartCol);
									FileOutputStream out = new FileOutputStream(Config.controllerFilePath);
									log.info("START_ROW : " + ustartRow + ", START_Col : " + ustartCol);
									LinearupdatedmainWB.write(out);
									out.flush();
									out.close();
									in.close();

									// Set values in the report (outside of the try block)
									webDriver.getReport().setTrasactionType(controllerTransactionType.toString());
									webDriver.getReport().setTestDescription(testDescription.toString());
									webDriver.getReport().setTestcaseId(controllerTestCaseID.toString());
									webDriver.getReport().setCycleDate(cycleDateCellValue.toString());

									// Assign the report value
									report = webDriver.getReport();
								} catch (Exception e) {
									log.error(e.getMessage(), e);
									e.printStackTrace();
								}

								// Return the report value outside the try-catch block
								return report;

							} else {
								try {
									ustartCol = columnIndex + 1;
									FileInputStream in = new FileInputStream(Config.controllerFilePath);
									Workbook LinearupdatedmainWB = null;
									if (Config.controllerFilePath.endsWith(".xls")) {
										LinearupdatedmainWB = new HSSFWorkbook(in);
									} else if (Config.controllerFilePath.endsWith(".xlsx")
											| Config.controllerFilePath.endsWith(".xlsm")) {
										LinearupdatedmainWB = new XSSFWorkbook(in);
									}
									MainControllerSheet = LinearupdatedmainWB.getSheet("StartPointer");
									updatedstartpointerRow = MainControllerSheet.getRow(1);
									updatedstartpointerValues = WebHelperUtil.getValueFromHashMap(MainControllerSheet);
									UPDATED_START_ROW = Integer
											.parseInt(updatedstartpointerValues.get("StartRow").toString());
									UPDATED_START_COLUMN = Integer
											.parseInt(updatedstartpointerValues.get("StartCol").toString());
									updatedstartpointerRow.getCell(UPDATED_START_ROW).setCellValue(ustartRow);
									updatedstartpointerRow.getCell(UPDATED_START_COLUMN).setCellValue(ustartCol);
									FileOutputStream out = new FileOutputStream(Config.controllerFilePath);
									LinearupdatedmainWB.write(out);
									out.flush();
									out.close();
									in.close();

									log.info("Updated StartPointer with START_ROW as " + ustartRow + ", START_Col as "
											+ ustartCol);

									webDriver.report.setTrasactionType(controllerTransactionType.toString());
									webDriver.report.setTestDescription(testDescription.toString());
									webDriver.report.setTestcaseId(controllerTestCaseID.toString());
									webDriver.report.setGroupName(controllerGroupName.toString());
									WebHelper.frmDate = new Date();
									webDriver.report.setFromDate(Config.dtFormat.format(WebHelper.frmDate));
									webDriver.report.setStatus("");
									webDriver.report.setMessage("");
									webDriver.report.setScreenShot("");
									
									if (controllerTransactionType.toString().equalsIgnoreCase("CloseBrowser")) {
										// WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, 90);
										WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver,
												Duration.ofSeconds(90));// Selenium4
										WaitForPageLoad.until(ExpectedConditions
												.invisibilityOfElementLocated(By.xpath(claimsloader)));
										Automation.driver.close();
										Automation.driver.quit();
										webDriver.report.setStatus("PASS");
										ExcelUtility.writeReportPAS(webDriver.getReport());
									} else if (controllerTransactionType.toString().equalsIgnoreCase("CloseAndLaunchNewBrowser")||
									controllerTransactionType.toString().equalsIgnoreCase("CloseAndLaunchBrowser")){
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
										webDriver.report.setStatus("PASS");
										ExcelUtility.writeReportPAS(webDriver.getReport());
										
									} else if (controllerTransactionType.toString()
												.equalsIgnoreCase("ClearCache")) {

									WebHelper.clearCache(WebHelper.currentdriver);
									webDriver.report.setStatus("PASS");
									webDriver.report.setMessage("Cookies deleted sucessfully");
									ExcelUtility.writeReportPAS(webDriver.getReport());
									
									}else if (controllerTransactionType.toString().toLowerCase().contains("claimlogin".toLowerCase())
											&& ((Config.environment != null))){
											
										if ( Config.environment.toString().toLowerCase().contains("Cloud".toLowerCase())) {
											
										
										// Your logic here
										log.info("Transaction is Login - Closing the existing window and reopening a new one...");
										Automation.setUp();
										//TransactionMapping.TransactionInputData("Login2");
										//Thread.sleep(3000);
										//WebHelperUtil.saveScreenShot();
										//columnChanged = false;
										//webDriver.getReport()
											//	.setTrasactionType(controllerTransactionType.toString());

												report = TransactionMapping.TransactionInputData(
														controllerTestCaseID.toString(), controllerTransactionType.toString(),
														Config.transactionInputFilePath);		
											}else {
												log.info("Non Cloud env...");
												report = TransactionMapping.TransactionInputData(
														controllerTestCaseID.toString(), controllerTransactionType.toString(),
														Config.transactionInputFilePath);
											}
												}
									else {
										report = TransactionMapping.TransactionInputData(
												controllerTestCaseID.toString(), controllerTransactionType.toString(),
												Config.transactionInputFilePath);
									}
								} catch (Exception e) {
									log.error("Issue while executing " + controllerTransactionType.toString()
											+ " at Row: " + (rowIndex + 1) + ", Column : " + (columnIndex + 1));
								}
							}
						} else {
							log.info("No Transaction Found in the Maincontroller at Row : " + (rowIndex + 1)
									+ ", Column : " + (columnIndex + 1));
						}
					}
				} else {
					log.info("Execute Flag is NOT 'Y' at Row : " + (rowIndex + 1));
				}
			} else {
				log.info("Execute Flag is not set at Row : " + (rowIndex + 1));
			}
		}
		startCol = execFlag + 1;
		return report;
	}

	public void recoveryhandler() {
		try {
			
			if (Config.recovery_scenario.equalsIgnoreCase("TRUE")) {

			log.info("Recovery Triggered to mark testcases N...");
			pauseExecution = true;
			log.info("Updating MainController ExecutionFlag");
			FileInputStream in;
			Workbook mainWB = null;
			in = new FileInputStream(Config.controllerFilePath);
			if (Config.controllerFilePath.endsWith(".xls")) {
				mainWB = new HSSFWorkbook(in);
			} else if (Config.controllerFilePath.endsWith(".xlsx") | Config.controllerFilePath.endsWith(".xlsm")) {
				mainWB = new XSSFWorkbook(in);
			}
			MainControlSheet = mainWB.getSheet("MainControlSheet");
			controlsheet = WebHelperUtil.getValueFromHashMap(MainControlSheet);
			int maincontrollersheetrowCount = MainControlSheet.getLastRowNum() + 1;
			Cell SC_NO_Cell, TC_ID_Cell;
			// String TXN_NAME = null;
			int SC_Header_Num = Integer.parseInt(controlsheet.get("SC_NO").toString());
			int TC_Header_Num = Integer.parseInt(controlsheet.get("TestCaseID").toString());
			// controllerTransactionType.toString();
			int ExecuteFlag_Header_Num = Integer.parseInt(controlsheet.get("ExecuteFlag").toString());
			TC_ID_Cell = controllerRow.getCell(TC_Header_Num);
			SC_NO_Cell = controllerRow.getCell(SC_Header_Num);
			// TXN_NAME = controllerTransactionType.toString();
			String updateFlag = "";
			int allN = 0, selectiveN = 0;
			// If SC_NO is not provided, No changes would be made to MainController
			// Execution Flag

			// Assuming SC_NO_Cell is an instance of org.apache.poi.ss.usermodel.Cell
			// if (SC_NO_Cell == null || SC_NO_Cell.getCellType() == CellType.BLANK ||
			// SC_NO_Cell.getStringCellValue().isEmpty()) {
			if ((SC_NO_Cell == null) || SC_NO_Cell.toString().trim().isEmpty()) {

				log.info("MainControlSheet: SC_NO is blank in row " + (controllerRow.getRowNum() + 1)
						+ ". Enter appropriate SC_NO against all TCs.");
				pauseExecution = false;// Sel4
				return;
			}

			// Decide which rows to mark N based upon SC_NO
			updateFlag = "MarkAllAsN";

			// Get SC_NO_Cell value
			String targetSC_NO = SC_NO_Cell.toString().trim();

			// Loop from (current failed row + 1) till last row of MainController to mark
			// Applicable row as N
			for (int rowindx = controllerRow.getRowNum(); rowindx < maincontrollersheetrowCount; rowindx++) {

				Row tempRow = null;
				try {
					tempRow = MainControlSheet.getRow(rowindx);
				} catch (Exception e) {
					tempRow = null;
				}
				if (tempRow == null)
					continue;

				if (updateFlag.equals("MarkAllAsN")) {
					Cell tempSC_NO_Cell = tempRow.getCell(SC_Header_Num);
					if (tempSC_NO_Cell != null && tempSC_NO_Cell.toString().trim().equalsIgnoreCase(targetSC_NO)) {
						// Mark the ExecuteFlag as "N" for this row
						tempRow.getCell(ExecuteFlag_Header_Num).setCellValue("N");
						selectiveN++;
					}
				}
			}

			if (allN > 0)
				log.info("As Pre SC_NO is failed, all remaining SC_NO against TC " + TC_ID_Cell.toString()
						+ " are marked as N");
			// Update log messages
			if (selectiveN > 0) {
				log.info("Execution flag set to 'N' for all rows with SC_NO " + targetSC_NO);
				// pauseExecution = true;//Sel4
			}

			FileOutputStream out;
			out = new FileOutputStream(Config.controllerFilePath);
			mainWB.write(out);
			out.flush();
			out.close();
			in.close();
			log.info("MainController ExecutionFlag Updation Completed");
			
			}else {
				log.info("Recovery scenario is set to FALSE in the config sheet");	
				
			}

		} catch (IOException e) {
			// log.error(e.getLocalizedMessage());12/05/23 - Not required
			e.printStackTrace();
		} catch (Exception e) {
			// log.error(e.getLocalizedMessage());12/05/23 - Not required
		}
	}

	/** Pauses the Execution **/
	public boolean pauseFun(String message) {
		String userInteraction = "TRUE";
		log.info("Entering PauseFun method...");

		webDriver.report.setGroupName(controllerGroupName.toString());
		webDriver.report.setTestcaseId(controllerTestCaseID.toString());
		webDriver.report.setTestDescription(testDescription);
		webDriver.report.setTrasactionType(controllerTransactionType.toString());

		// To avoid lengthy ExecutionDescription in SummaryResult which spreads across
		// multiple rows
		int temp = message.indexOf("(Session info:");
		if (temp != -1) {
			message = message.substring(0, temp).trim();
		}
		temp = message.indexOf("Build info:");
		if (temp != -1) {
			message = message.substring(0, temp).trim();
		}

		webDriver.report.setMessage(message);
		WebHelperClaims.toDate = new Date();
		webDriver.report.setToDate(Config.dtFormat.format(WebHelperClaims.toDate));
		webDriver.report.setFromDate(Config.dtFormat.format(WebHelper.frmDate));

		if (message == null) {
			message = "TestCase: " + controllerTestCaseID + " Tranasction: " + controllerTransactionType
					+ " Error: Unknown...";
			webDriver.report.setMessage(message);
		}

		if ((!message.equalsIgnoreCase("Check Detailed Results"))
				&& (!webDriver.report.getTrasactionType().startsWith("OASService"))
				&& (!webDriver.report.getTrasactionType().startsWith("DB"))) {
			WebHelperUtil.saveScreenShot("");
		}
		if (Config.getConfgiMapSize() != 0) {
			try {
				if (Config.userInteraction == null) {
					throw new Exception("Null Value Found for UserInteractioin Parameter");
				} else {
					userInteraction = Config.userInteraction;
				}
			} catch (Exception e) {
				JOptionPane.showConfirmDialog(webDriver.frame, "Null Value Found for UserInteractioin Parameter");
			}
		}

		/** Don't mark status as FAIL if transaction name is PAUSE **/
		if (!controllerTransactionType.toString().equalsIgnoreCase("PAUSE")) {

			webDriver.report.setStatus("FAIL");
			if (StringUtils.equalsIgnoreCase(Config.emailOnFailure, "true")) {

				String tempMsg = webDriver.report.getMessage();
				if (StringUtils.contains(tempMsg, "\"")) {
					tempMsg = tempMsg.replaceAll("\"", "'");
				}
				if (StringUtils.contains(tempMsg, ",")) {
					tempMsg = "\"\"" + tempMsg + "\"\"";
				}

				if (controllerTestCaseID.toString().equalsIgnoreCase(controllerGroupName.toString()))
					scrsubject = Config.projectName + " - Failure In - " + controllerTestCaseID.toString() + " -> "
							+ controllerTransactionType.toString();
				else
					scrsubject = Config.projectName + " - Failure In - " + controllerGroupName.toString() + " -> "
							+ controllerTestCaseID.toString() + " -> " + controllerTransactionType.toString();
				scrshotbody = "PFA the Screenshot for failed Transaction.$$Error message:$" + tempMsg;
				scrEmail = Config.emailId;
				scrreportsheet = Config.resultOutput;

				scrdetailreportsheet = Config.verificationResultPath;
				File tempfile = new File(scrdetailreportsheet);
				boolean exists = tempfile.exists();
				if (exists != true) {
					scrdetailreportsheet = "";
				}

				if (StringUtils.equalsIgnoreCase(Config.emailTransactionInfo, "true")) {
					scruniquenumbersheet = Config.transactionInfo;
				}

				strScrCompleteData = scrEmail + "##" + scrsubject + "##" + scrshotbody + "##" + scrreportsheet + "##"
						+ scrdetailreportsheet + "##" + FailScreen + "##" + scruniquenumbersheet;
				log.info("\nData sent to SendMail_v2.vbs : " + strScrCompleteData + "\n");
				try {
					Runtime.getRuntime().exec("wscript SendMail_v2.vbs " + (char) 34 + strScrCompleteData + (char) 34);
				} catch (IOException e) {
					log.error("Unable to send an Email");
					log.error(e.getMessage(), e);
					// throw new RuntimeException(e.getMessage(), e);
				}

			}

			/*
			 * scrshotbody = "PFA the Screenshot for failed Transaction"; scrEmail =
			 * Config.emailId; if (!scrEmail.toString().equalsIgnoreCase("NA") &&
			 * !scrEmail.equalsIgnoreCase("") && !scrEmail.equalsIgnoreCase(null)) {
			 * strScrCompleteData = scrEmail + "#" + scrshotbody + "#" + scrreportsheet +
			 * "#" + scrdetailreportsheet + "#" + FailScreen + "#" + ""; try {
			 * Runtime.getRuntime().exec("wscript SendMail.vbs " + (char) 34 +
			 * strScrCompleteData + (char) 34); } catch (IOException e) {
			 * e.printStackTrace(); } }
			 */
		}
		try {
			ExcelUtility.writeReportPAS(webDriver.getReport());
		} catch (IOException e) {
			log.error("Failed while executing ExcelUtility.writeReportPAS <-|-> LocalizeMessage "
					+ e.getLocalizedMessage() + " <-|-> Message " + e.getMessage() + " <-|-> Cause " + e.getCause(), e);
			throw new RuntimeException(e.getMessage(), e);
		}

		if (!userInteraction.equalsIgnoreCase("FALSE")) {
			webDriver.frame.setVisible(true);
			webDriver.frame.setAlwaysOnTop(true);
			webDriver.frame.setLocationRelativeTo(null);

			JOptionPane.setRootFrame(webDriver.frame);
			int response = JOptionPane.showConfirmDialog(webDriver.frame, message, "iTAF - Do you want to STOP...",
					JOptionPane.YES_NO_OPTION);
			if (response == JOptionPane.YES_OPTION) {
				stopExecution = true;
				pauseExecution = true;
			} else if (response == 1) {
				pauseExecution = false;
			} else {
				log.info("You have pressed cancel" + response);
				stopExecution = true;
				pauseExecution = true;
			}
			webDriver.frame.dispose();
		} else {
			recoveryhandler();
		}
		return pauseExecution;

	}

	public void batchRecoveryScenario(String batchNo) {
	}
	public static void setSiloParameters(String silo) {
		silo = silo.toUpperCase();
		switch (silo) {
		case "BILLING":
			Config.applicationdatabaseusername = Config.applicationdatabaseusername_billing;
			Config.applicationdatabasepassword = Config.applicationdatabasepassword_billing;
			Config.recovery_scenario = Config.recovery_scenario_billing;
			Config.applyStaticWait = Config.applyStaticWait_billing;
			break;
		case "PAS":
			Config.applicationdatabaseusername = Config.applicationdatabaseusername_pas;
			Config.applicationdatabasepassword = Config.applicationdatabasepassword_pas;
			Config.recovery_scenario = null;
			Config.applyStaticWait = Config.applyStaticWait_pas;
			break;
		case "CLAIMS":
			Config.applicationdatabaseusername = Config.applicationdatabaseusername_claims;
			Config.applicationdatabasepassword = Config.applicationdatabasepassword_claims;
			Config.recovery_scenario = Config.recovery_scenario_claims;
			Config.applyStaticWait = Config.applyStaticWait_claims;
			break;
		case "DIGITAL1ST":
			Config.applicationPlatform = "digital1st";
			Config.applyStaticWait = Config.applyStaticWait_pas;
			break;

		}
	}


}
