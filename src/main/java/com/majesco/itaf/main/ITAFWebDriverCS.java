package com.majesco.itaf.main;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import javax.swing.JOptionPane;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.majesco.itaf.report.utils.VisualReport;
import com.majesco.itaf.util.BackupUtility;
import com.majesco.itaf.util.CalendarSnippet;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.util.Jacob;
import com.majesco.itaf.verification.WebVerification;
import org.apache.commons.lang.StringUtils;

public class ITAFWebDriverCS extends ITAFWebDriver {

	private final static Logger log = LogManager.getLogger(ITAFWebDriverCS.class.getName());
	private MainController controller;

	ITAFWebDriverCS() {
	}

	protected void init() throws IOException {

		/*
		 * String reportsheet = ""; String detailreportsheet = ""; String
		 * uniquenumbersheet = "";// Mrinmayee String nodetailreportsheet = ""; String
		 * FailedScreen = ""; String body = ""; String strCompleteData = ""; String
		 * subject = "";
		 */
		try {
			// Unlockscreen("wscript D:/unlockScreen.vbs");
			try {
				Automation.clearTempDirectory();
			} catch (Exception e) {
				log.info("Unable to clear temp directory: " + e.getMessage());
			}
			if (!StringUtils.equalsIgnoreCase(Config.executionScope, "only_api")) {
				Automation.setUp();
			}

			controller = ObjectFactory.getMainController();
			
			/*
			 * subject = Config.projectName + " - Automated Test Run Results"; email =
			 * Config.emailId; reportsheet = Config.resultOutput; detailreportsheet =
			 * Config.verificationResultPath;
			 * 
			 * if (StringUtils.equalsIgnoreCase(Config.emailTransactionInfo, "true"))
			 * uniquenumbersheet = Config.transactionInfo; else uniquenumbersheet = "";
			 */
			
			log.info("formula or string:" + Config.controllerFilePath);
			controller.ControllerData(Config.controllerFilePath);

		} catch (Exception e) {
			log.error(e.getMessage(), e);
			report.setStatus("FAIL");
			report.setTestcaseId(controller.controllerTestCaseID.toString());
			report.setTrasactionType(controller.controllerTransactionType.toString());
			try {
				controller.pauseFun("d: " + controller.controllerTestCaseID + ", Tranasction: "
						+ controller.controllerTransactionType + ", Error: " + e.getMessage());
			} catch (Exception e1) {
				log.error(e1.getMessage(), e1);
				controller.pauseFun("File Not Found");
			}
		} finally {

			try {
				// 11/Nov:- Commenting below line as its failing during API run.
				// API TC's doesn't invoke any Browser & hence no need to close if it's not
				// running.
				// Automation.driver.quit();
				if (com.majesco.itaf.util.CalendarSnippet.isProcessRunning("IEDriverServer.exe"))
					com.majesco.itaf.util.CalendarSnippet.killProcess("IEDriverServer.exe");
				if (com.majesco.itaf.util.CalendarSnippet.isProcessRunning("chromedriver.exe"))
					com.majesco.itaf.util.CalendarSnippet.killProcess("chromedriver.exe");
				if (com.majesco.itaf.util.CalendarSnippet.isProcessRunning("msedgedriver.exe"))
					com.majesco.itaf.util.CalendarSnippet.killProcess("msedgedriver.exe");
				if (com.majesco.itaf.util.CalendarSnippet.isProcessRunning("firefox.exe"))
					com.majesco.itaf.util.CalendarSnippet.killProcess("firefox.exe");

			} catch (Exception e) {
				log.info(e.getMessage(), e);
			}

			// TM: 16-01-2015
			if (Config.userInteraction.equalsIgnoreCase("true")) {
				frame.setVisible(true);
				frame.setAlwaysOnTop(true);
				frame.setLocationRelativeTo(null);
				JOptionPane.setRootFrame(frame);
				JOptionPane.showMessageDialog(frame, "Execution Completed");
				frame.dispose();
			}

			
			//Below code not required as we are not sending an individual failure emails
			/*
			 * File tempfile = new File(detailreportsheet); boolean exists =
			 * tempfile.exists();
			 * 
			 * if (!email.isEmpty()) { if (exists == true) { // body = "PFA the Reports";
			 * int TotalVerificationPoints = WebHelperPAS.TotalpassCount +
			 * WebHelperPAS.TotalfailCount;
			 * 
			 * body = "Total Verification Points :" + TotalVerificationPoints + "$" +
			 * "Total Passed:" + WebHelperPAS.TotalpassCount + "$" + "Total Failed:" +
			 * WebHelperPAS.TotalfailCount + "$" +
			 * "Please see attached reports for details.";
			 * 
			 * System.out.println(body); strCompleteData = email + "#" + subject + "#" +
			 * body + "#" + reportsheet + "#" + detailreportsheet + "#" + FailedScreen + "#"
			 * + uniquenumbersheet;
			 * 
			 * Runtime.getRuntime().exec("wscript SendMail.vbs " + (char) 34 +
			 * strCompleteData + (char) 34); } else { body = "PFA the Report";
			 * strCompleteData = email + "#" + subject + "#" + body + "#" + reportsheet +
			 * "#" + FailedScreen + "#" + nodetailreportsheet + "#" + uniquenumbersheet;
			 * Runtime.getRuntime().exec("wscript SendMail.vbs " + (char) 34 +
			 * strCompleteData + (char) 34); } }
			 */
			
			Date finalfrmDate = new Date();
			report.setIteration(Config.cycleNumber);
			report.setTestcaseId("XML Comparison Macro");
			report.setGroupName("XML Comparison Macro");
			// report.setCycleDate("Common");
			report.setTrasactionType("XML Comparison Macro");
			report.setTestDescription("XML Comparison Macro");
			report.setFromDate(Config.dtFormat.format(finalfrmDate));

			if ("True".equalsIgnoreCase(Config.openAPIserviceComparison)) {
				String openAPIserviceUtility = Config.openAPIServiceComparisonUtilityPath;
				log.info("openAPIserviceUtility : " + openAPIserviceUtility);
				threadSleep(5000);

				File utilityFile = new File(openAPIserviceUtility);
				if (utilityFile.exists()) {
					System.out.println(openAPIserviceUtility + " macro is running...");
					log.info(openAPIserviceUtility + " macro is running...");
					Jacob.main(openAPIserviceUtility, "openAPIserviceVerification_Linear");
				}

				File f1 = new File(Config.resultFilePath + "OpenAPIServiceLinear_Comparison_Results.xls");
				if (f1.exists()) {
					log.info("OpenAPIService_Comparison_Results file exists.");
					report.setStatus("PASS");
					report.setMessage(
							"API : XML comparison completed successfully.\nResults placed at: " + f1.getAbsolutePath());

				} else {
					log.info("OpenAPIService_Comparison_Results file does not exist ");
					report.setStatus("FAIL");
					report.setMessage("API : Issue in XML comparison");
				}
				report.setToDate(Config.dtFormat.format(new Date()));
				ExcelUtility.writeReportPAS(report);
				// killExcels();
			}
			try {
				String statusUtility = Config.executionStatusReportUtility;
				log.info("Status Utility : " + statusUtility);

				if (statusUtility == null) {
					log.info("Execution Status Tracker utility path not found in Config");
				} else {
					File execStatusFile = new File(statusUtility);
					if (execStatusFile.exists()) {
						log.info(statusUtility + " macro is runnnig...");
						Jacob.main(statusUtility, "report");
						log.info("Status Report sent");
					} else {
						System.out.println("Execution Status Tracker utility not found.");
					}
				}
			} catch (Exception e) {
				log.error(e.getMessage(), e);
				e.printStackTrace(); // TODO Auto-generated catch block
			}

			if ("True".equalsIgnoreCase(Config.openAPIserviceComparison)) {
				killExcels();
			}

			try {
				VisualReport.generateVisualReport();
				log.info("Please refer following sheet for visual report: " + Config.resultFilePath
						+ "VisualReport.xlsx");
			} catch (Exception e) {
				log.info("[OPTIONAL]Could not create visual report, because of following reason: " + e.getMessage());
				e.printStackTrace();
			}
			// killExcels();

			// Initializing Backup Process
			if (Config.enableBackup != null && Config.enableBackup.equalsIgnoreCase("TRUE")
					&& Config.backupFilepath != null && Config.backupFileList != null) {
				log.info("INITIALIZING BACKUP PROCESS......");
				try {
					String backupFilePath = BackupUtility.initBackup(Config.backupFileList, Config.backupFilepath);
					log.info("Backup is available at following path: " + backupFilePath);
				} catch (Exception e) {
					log.info("Exception occured while taking backup: " + e.getMessage());
					// e.printStackTrace();
				}
			}
			// Initializing Backup Process END

			threadSleep(200);
			log.info("\n\n==================================iTAF Execution Finish!!=================================");
			System.exit(0);
		} // finally ends
	}

	public void DataInput(String filePath, String testcaseID, String transactionType, String transactionCode,
			String operationType) throws Exception {
		System.out.println("filePath is:" + filePath);
		if (transactionCode == null) {
			transactionCode = transactionType;
		}
		System.out.println(transactionCode);

		if (operationType.equalsIgnoreCase("InputandVerify") && !operationType.isEmpty()) {
			ExcelUtility.GetDataFromValues(filePath, testcaseID.toString(), transactionType.toString());
			WebVerification.performVerification(transactionType, testcaseID);
		} else if (!operationType.equalsIgnoreCase("Verify") && !operationType.isEmpty()) {
			ExcelUtility.GetDataFromValues(filePath, testcaseID.toString(), transactionType.toString());
		} else if (!operationType.equalsIgnoreCase("Input") && !operationType.isEmpty()) {
			System.out.println("------------------------------------------------------");
			System.out.println("-------------------------INPUT-----------------------------");
			System.out.println("------------------------------------------------------");
			WebVerification.performVerification(transactionType, testcaseID);
		}
	}

	public static void Unlockscreen(String filepath) {

		try {
			Runtime.getRuntime().exec(filepath);
		} catch (IOException e) {
			log.error(e.getMessage(), e);
			System.out.println("Unlock file is not available");

		}

	}

	private void killExcels() {
		try {
			CalendarSnippet.killProcess("EXCEL.EXE");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void threadSleep(int i) {
		try {
			Thread.sleep(i);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}

}
