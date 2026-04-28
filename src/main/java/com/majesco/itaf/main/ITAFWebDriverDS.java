package com.majesco.itaf.main;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import javax.swing.JOptionPane;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.majesco.itaf.report.utils.VisualReport;
import com.majesco.itaf.util.BackupUtility;
import com.majesco.itaf.util.CalendarSnippet;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.util.Jacob;
import com.majesco.itaf.verification.WebVerification;

public class ITAFWebDriverDS extends ITAFWebDriver {

	private final static Logger log = LogManager.getLogger(ITAFWebDriverDS.class.getName());
	private static MainController controller;

	ITAFWebDriverDS() {
	}

	public void init() throws IOException {
		/*
		 * String reportsheet = ""; String detailreportsheet = ""; String
		 * nodetailreportsheet = ""; String FailedScreen = ""; String body = "", subject
		 * = "", uniquenumbersheet = ""; String strCompleteData = "";
		 */
		try {
			CalendarSnippet.killProcess("EXCEL.EXE");
			try {
				Automation.clearTempDirectory();
			} catch (Exception e) {
				log.info("Unable to clear temp directory: " + e.getMessage());
			}
			if (Config.resetStartPointer != null && !Config.resetStartPointer.equalsIgnoreCase("NA")) {
				Jacob.main(Config.resetStartPointer, "Macro1");
			}
			Automation.setUp();
			controller = ObjectFactory.getMainController();
			
			/*
			 * email = Config.emailId; reportsheet = Config.resultOutput; detailreportsheet
			 * = Config.verificationResultPath; subject = Config.projectName +
			 * " - Automated Test Run Results"; if
			 * (StringUtils.equalsIgnoreCase(Config.emailTransactionInfo, "true"))
			 * uniquenumbersheet = Config.transactionInfo; else uniquenumbersheet = "";
			 */
			// END of set mail parameters

			log.info("MainController File Path is :" + Config.controllerFilePath);

			controller.ControllerData(Config.controllerFilePath); // Start Transaction

		} catch (Exception e) {
			log.error(e.getMessage(), e);
			log.error(e.getLocalizedMessage());
			report.setStatus("FAIL");
			report.setTestcaseId(controller.controllerTestCaseID.toString());
			report.setTrasactionType(controller.controllerTransactionType.toString());
			report.setGroupName(controller.controllerGroupName.toString());
			try {
				controller.pauseFun("TestCaseID: " + controller.controllerTestCaseID + ", Tranasction: "
						+ controller.controllerTransactionType + ", Error: " + e.getMessage());
			} catch (Exception e1) {
				log.error(e1.getMessage(), e1);
				controller.pauseFun("File Not Found");
			}
		} finally {

			Date finalfrmDate = new Date();
			report.setIteration(Config.cycleNumber);
			report.setTestcaseId("Common");
			report.setGroupName("Common");
			report.setTrasactionType("XML Comparison Macro");
			report.setTestDescription("XML Comparison Macro");
			report.setFromDate(Config.dtFormat.format(finalfrmDate));

			if ("True".equalsIgnoreCase(Config.openAPIserviceComparison)) {
				String openAPIserviceUtility = Config.openAPIServiceComparisonUtilityPath;
				log.info("openAPIserviceUtility : " + openAPIserviceUtility);
				threadSleep(5000);

				File utilityFile = new File(openAPIserviceUtility);
				if (utilityFile.exists()) {
					log.info(openAPIserviceUtility + " macro is running...");
					Jacob.main(openAPIserviceUtility, "openAPIserviceVerification_Linear");
				}

				File f1 = new File(Config.resultFilePath + "OpenAPIServiceLinear_Comparison_Results.xls");
				if (f1.exists()) {
					log.info("OpenAPIService_Comparison_Results file exist ");
					report.setStatus("PASS");
					report.setMessage("OpenAPIService Verification Executed Successfully");
					ExcelUtility.writeReportPAS(report);
					killExcels();
				} else {
					log.info("OpenAPIService_Comparison_Results file does not exist ");
				}
			}

			if (controller.controllerTransactionType.toString() != null)
				if (controller.controllerTransactionType.toString().equalsIgnoreCase("END")
						|| controller.controllerTransactionType.toString().equalsIgnoreCase("END1")) {
					if (controller.controllerTransactionType.toString().equalsIgnoreCase("END")) {
						ExcelUtility.writeReportPAS(report);
						System.exit(0);
					} else {
						ExcelUtility.writeReportPAS(report);
						Jacob.main("D:\\iTAFSeleniumWeb\\UpdateStartPointer.xlsm", "Macro1");
						System.exit(0);
					}
				}

			//Below code not required as we are not sending an individual failure emails
			/*
			 * if (!email.toString().equalsIgnoreCase("NA") && !email.equalsIgnoreCase("")
			 * && !email.equalsIgnoreCase(null)) { File tempfile = new
			 * File(detailreportsheet); boolean exists = tempfile.exists();
			 * 
			 * body = "PFA the Report"; if (exists == true) { strCompleteData = email + "##"
			 * + subject + "##" + body + "##" + reportsheet + "##" + detailreportsheet +
			 * "##" + FailedScreen + "##" + uniquenumbersheet; } else { strCompleteData =
			 * email + "##" + subject + "##" + body + "##" + reportsheet + "##" +
			 * FailedScreen + "##" + nodetailreportsheet + "##" + uniquenumbersheet; }
			 * Runtime.getRuntime().exec("wscript SendMail_v2.vbs " + (char) 34 +
			 * strCompleteData + (char) 34);
			 * 
			 * }
			 */

			// mail sent macro call
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
				e.printStackTrace();
			}

			try {
				VisualReport.generateVisualReport();
				log.info("Please refer following sheet for visual report: " + Config.resultFilePath
						+ "VisualReport.xlsx");
			} catch (Exception e) {
				log.info("[OPTIONAL]Could not create visual report, because of following reason: " + e.getMessage());
				e.printStackTrace();
			}

			killExcels();
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
			if (Config.userInteraction.equalsIgnoreCase("true")) {
				frame.setVisible(true);
				frame.setAlwaysOnTop(true);
				frame.setLocationRelativeTo(null);
				JOptionPane.setRootFrame(frame);
				JOptionPane.showMessageDialog(frame, "Execution Completed");
				frame.dispose();
			}
			System.exit(0);
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

	@SuppressWarnings("unused")
	public void DataInput(String filePath, String testcaseID, String transactionType, String transactionCode,
			String operationType) throws Exception {

		if (transactionCode == null) {
			transactionCode = transactionType;
		}
		log.info(transactionCode);

		if (operationType.equalsIgnoreCase("InputandVerify") && !StringUtils.isEmpty(operationType.toString())) {
			ExcelUtility.GetDataFromValues(filePath, testcaseID.toString(), transactionType.toString());
			WebVerification.performVerification(transactionType, testcaseID);// Mandar

		} else if (operationType.equalsIgnoreCase("Capture") && !StringUtils.isEmpty(operationType)) {
			log.info("Capture");
			String OperationType = "Capture";
			WebVerification.performVerification(transactionType, testcaseID);
		} else if (!operationType.equalsIgnoreCase("Verify") && !StringUtils.isEmpty(operationType.toString())) {
			ExcelUtility.GetDataFromValues(filePath, testcaseID.toString(), transactionType.toString());
		} else if (!operationType.equalsIgnoreCase("Input") && !StringUtils.isEmpty(operationType.toString())) {
			log.info("INPUT");
			WebVerification.performVerification(transactionType, testcaseID);// Mandar
		}
	}
}
