package com.majesco.itaf.main;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import javax.swing.JOptionPane;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.vo.Reporter;

public class MainControllerLNA extends MainController {
	private final static Logger log = LogManager.getLogger(MainControllerLNA.class.getName());
	private static HashMap<String, Integer> sheetValues = new HashMap<String, Integer>();
	private static int startCol = 0;
	private static int startRow = 0;
	private static Row controllerRow = null;
	private static String scrshotbody = "";
	private static String strScrCompleteData = "";
	private static String scrEmail = "";
	private static String scrreportsheet = "";
	private static String scrdetailreportsheet = "";
	private static String scruniquenumbersheet = "";
	private static String scrsubject = "";
	private static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();

	@Override
	public Reporter ControllerData(String FilePath) throws NullPointerException, Exception {
		System.out.println("In MainController value of pauseExecution1:" + pauseExecution);
		Reporter report = new Reporter();
		Sheet reqSheet = ExcelUtility.GetSheet(Config.controllerFilePath, "MainControlSheet");
		sheetValues = WebHelperUtil.getValueFromHashMap(reqSheet);
		int execFlag = sheetValues.get("ExecuteFlag");
		int rowCount = reqSheet.getLastRowNum() + 1;
		int colCount = 0;
		boolean isStartFound = false;
		for (int rowindex = 0; rowindex < rowCount && !isStartFound; rowindex++) {
			controllerRow = reqSheet.getRow(rowindex);
			if (controllerRow.getCell(execFlag) != null) {
				if (controllerRow.getCell(execFlag).toString().equals("Y")) {
					colCount = controllerRow.getLastCellNum() + 1;
					for (int colIndex = execFlag + 1; colIndex < colCount; colIndex++) {
						Cell cellVal = controllerRow.getCell(colIndex);
						if (cellVal != null) {
							if (cellVal.toString().equalsIgnoreCase("START")) {
								startCol = colIndex;
								startRow = rowindex;
								isStartFound = true;
								break;

							}
						} else {
							System.out.println("START not Found");
						}

					}
				} else {
					System.out.println("Execute Flag is N");
				}
			}

		}
		log.info("In MainController value of pauseExecution2:" + pauseExecution);
		for (int rowIndex = startRow; rowIndex < rowCount; rowIndex++) {
			pauseExecution = false;
			controllerRow = reqSheet.getRow(rowIndex);
			colCount = controllerRow.getLastCellNum() + 1;
			testDescription = WebHelperUtil.getCellData("Test_Description", reqSheet, rowIndex, sheetValues);
			Cell executeFlag = controllerRow.getCell(execFlag);
			controllerTestCaseID = controllerRow.getCell(sheetValues.get("TestCaseID"));
			controllerGroupName = controllerRow.getCell(sheetValues.get("GroupName"));

			// added for PAS reporting for QuoteID
			if (sheetValues.get("QuoteId") != null) {
				controllerQuoteId = controllerRow.getCell(sheetValues.get("QuoteId")).toString();
			}

			if (controllerTestCaseID.getStringCellValue().equalsIgnoreCase("") || controllerTestCaseID.equals(null)) {
				System.out.println("No KeyWord Found");
				continue;
			}

			if (executeFlag != null) {
				if (executeFlag.toString().equalsIgnoreCase("Y")) {
					for (int columnIndex = startCol + 1; columnIndex < colCount && !pauseExecution; columnIndex++) {
						controllerTransactionType = controllerRow.getCell(columnIndex);
						log.info("Value of controllerTransactionType: " + controllerTransactionType);
						log.info("Value of controllerTestCaseID: " + controllerTestCaseID);
						if (controllerTransactionType != null
								&& StringUtils.isNotBlank(controllerTransactionType.getStringCellValue())) {
							if (controllerTransactionType.toString().equalsIgnoreCase("PAUSE")) {
								pauseFun("Do You Wish To Continue");
							} else {
								report = TransactionMapping.TransactionInputData(controllerTestCaseID.toString(),
										controllerTransactionType.toString(), Config.transactionInputFilePath);
							}
						} else {
							log.info("No Transaction Found in the Maincontroller at Cell : " + columnIndex);
						}

					}
				}
			} else {
				log.info("Execute Flag is not Set");
			}

		}
		log.info("In MainController value of pauseExecution3:" + pauseExecution);
		startCol = execFlag + 1;
		return report;
	}

	@Override
	public boolean pauseFun(String message) {

		/**
		 * DS:18-07-2014:Replacing timeout in msg String tempMsg = "Timed out after
		 * CONFIGTIMEOUT seconds waiting for presence of element located by" ; tempMsg =
		 * tempMsg.replace("CONFIGTIMEOUT",
		 * Automation.configHashMap.get("TIMEOUT").toString()); if(message!= null)
		 * message = message.replace(tempMsg, "Element not found");
		 **/

		String userInteraction = "TRUE";
		try {

			webDriver.getReport().setGroupName(controllerGroupName.toString());
			webDriver.getReport().setTestcaseId(controllerTestCaseID.toString());
			webDriver.getReport().setTestDescription(testDescription);
			webDriver.getReport().setTrasactionType(controllerTransactionType.toString());
			WebHelperPAS.toDate = new Date();
			webDriver.getReport().setMessage(message);
			webDriver.getReport().setToDate(Config.dtFormat.format(WebHelperPAS.toDate));
			// Below line added for PAS QuoteID reporting
			webDriver.getReport().setStrQuoteId(controllerQuoteId);
			WebHelperLNA.saveScreenShot();
			if (message == null) {
				message = "TestCase: " + controllerTestCaseID + " Tranasction: " + controllerTransactionType
						+ " Error: Unknown...";
				webDriver.getReport().setMessage(message);
			}
			if (Config.getConfgiMapSize() != 0) {
				try {
					if (Config.userInteraction == null) {

						throw new Exception("Null Value Found for UserInteractioin Parameter");

					} else {
						userInteraction = Config.userInteraction;
					}
				} catch (Exception e) {
					log.error(e.getMessage(), e);
					JOptionPane.showConfirmDialog(webDriver.getFrame(),
							"Null Value Found for UserInteractioin Parameter");
				}
			}

			/** Don't mark status as FAIL if transaction name is PAUSE **/
			if (!controllerTransactionType.toString().equalsIgnoreCase("PAUSE")) {
				webDriver.getReport().setStatus("FAIL");

				if (StringUtils.equalsIgnoreCase(Config.emailOnFailure, "true")) {
					scrsubject = Config.projectName + " - Automated Test Run Results";
					scrshotbody = "PFA the Screenshot for failed Transaction";
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

					strScrCompleteData = scrEmail + "#" + scrsubject + "#" + scrshotbody + "#" + scrreportsheet + "#"
							+ scrdetailreportsheet + "#" + FailScreen + "#" + scruniquenumbersheet;
					try {
						Runtime.getRuntime().exec("wscript SendMail.vbs " + (char) 34 + strScrCompleteData + (char) 34);
					} catch (IOException e) {
						log.error(e.getMessage(), e);
						throw new RuntimeException(e.getMessage(), e);
					}

				}
			}

			if (!userInteraction.equalsIgnoreCase("FALSE")) {
				webDriver.getFrame().setVisible(true);
				webDriver.getFrame().setAlwaysOnTop(true);
				webDriver.getFrame().setLocationRelativeTo(null);
				int response;

				JOptionPane.setRootFrame(webDriver.getFrame());
				// if (Config.projectName.equals("DistributionManagement"))
				response = JOptionPane.showConfirmDialog(webDriver.getFrame(), message,
						"iTAF - Do you wish to Continue...", JOptionPane.YES_NO_OPTION);// Minaakshi : 03-10-2018 :
				log.info(response);
				if (response == JOptionPane.YES_OPTION) {
					pauseExecution = true;
				} else if (response == 1) {
					/** Call error reporting and stop execution **/
					try {
						ExcelUtility.writeReportPAS(webDriver.getReport());
						if (Config.projectName.equals("DistributionManagement")) {
							System.gc();
							// Minaakshi 03-10-2018
							Automation.driver.quit();
							System.exit(0);
						}
					} catch (IOException e) {
						log.error(e.getMessage(), e);
						throw new RuntimeException(e.getMessage(), e);
					}
				} else {
					log.info("You have pressed cancel" + response);
					pauseExecution = true;
				}
			} else {
				webDriver.getReport().setMessage(message);
				pauseExecution = true;
			}
		} finally {
			webDriver.getFrame().dispose();
		}
		return pauseExecution;

	}

	@Override
	public void recoveryhandler() {

	}

	@Override
	public void batchRecoveryScenario(String batchNo) {

	}

}
