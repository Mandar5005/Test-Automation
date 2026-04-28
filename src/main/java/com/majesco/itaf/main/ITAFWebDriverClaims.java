package com.majesco.itaf.main;

import java.io.File;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
import javax.swing.JOptionPane;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.majesco.itaf.outlookmail.ReportScreenshot;
import com.majesco.itaf.outlookmail.SendMail;
import com.majesco.itaf.util.CalendarSnippet;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.util.Jacob;
import com.majesco.itaf.verification.WebVerification;

public class ITAFWebDriverClaims extends ITAFWebDriver {

	private final static Logger log = LogManager.getLogger(ITAFWebDriverClaims.class.getName());
	private static MainController controller;

	ITAFWebDriverClaims() {
	}

	public void init() throws IOException {
		/*
		 * String reportsheet = ""; String detailreportsheet = ""; String
		 * nodetailreportsheet = ""; String FailedScreen = ""; String body = "", subject
		 * = "", uniquenumbersheet = ""; String strCompleteData = "";
		 */
		String summaryReportName = null;

		try {
			log.info("Closing all open excel sheets...");
			CalendarSnippet.killProcess("EXCEL.EXE");
			log.info("Completed.");
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

			// set mail parameters
			
			/*
			 * email = Config.emailId; reportsheet = Config.resultOutput; detailreportsheet
			 * = Config.verificationResultPath; subject = Config.projectName +
			 * " - Automated Test Run Results"; if
			 * (StringUtils.equalsIgnoreCase(Config.emailTransactionInfo, "true"))
			 * uniquenumbersheet = Config.transactionInfo; else uniquenumbersheet = "";
			 */
			// END of set mail parameters

			log.info("MainController File Path is :" + Config.controllerFilePath);
			try {
				controller.ControllerData(Config.controllerFilePath); // Start Transaction
			} catch (Exception e) {
				// Log the specific exception details
				log.error("An error occurred while initializing: " + e.getMessage(), e);

				// Additional logging to identify where the failure occurred
				log.error("Failed at controller.ControllerData(Config.controllerFilePath)");

				// Rethrow the exception to handle it further if needed
				throw e;
			}

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

			if (controller.controllerTransactionType.toString() != null)// Sel4 major issue 19/03/24
				// if (controller.controllerTransactionType != null &&
				// controller.controllerTransactionType.toString() != null)
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

			// BACKUP
			String target_path = Config.backupFilepath;

			if (!(target_path == null || target_path.isEmpty())) {
				log.info("Backup Started...");

				File targetPathBCK = new File(target_path);

				if (!targetPathBCK.exists()) {
					targetPathBCK.mkdir();
				}

				String input_folder = Config.inputDataFilePath;
				String result_folder = Config.resultFilePath;
				String mainC_file = Config.controllerFilePath;
				String uniqueNo_file = Config.transactionInfo; // Added for Claims 24/05/2023
				Calendar cd = Calendar.getInstance();

				int currMonth = cd.get((Calendar.MONTH)) + 1;
				String s_prefix = cd.get(Calendar.DAY_OF_MONTH) + "_" + currMonth + "_" + cd.get(Calendar.YEAR) + "_"
						+ cd.get(Calendar.HOUR) + "_" + cd.get(Calendar.MINUTE) + "_" + cd.get(Calendar.SECOND);

				target_path = target_path + "/" + "Bck_" + s_prefix;
				File targetPath = new File(target_path);

				if (!targetPath.exists()) {
					targetPath.mkdir();
				}
				log.info("Copying Input Folder...");
				backup(new File(input_folder), new File(target_path + "/Input"));
				log.info("Copying Result Folder...");

				File[] resultFolderContents = new File(result_folder).listFiles();
				if (resultFolderContents != null) {
					for (File file : resultFolderContents) {
						// Check if the item is a directory and its name is not in the exclusion list
						if (file.isDirectory() && !(file.getName().equals("AutomatedBackup")
								|| file.getName().equals("ExecutionStatusReportLogs")
								|| file.getName().equals("Outlook"))) {
							// Copy the entire folder and its contents recursively
							backup(file, new File(target_path + "/Results/" + file.getName()));
						} else if (!(file.getName().equals("AutomatedBackup")
								|| file.getName().equals("ExecutionStatusReportLogs")
								|| file.getName().equals("Outlook"))) {
							// If it's a file, or a directory to be excluded, copy it directly
							backup(file, new File(target_path + "/Results/"));
						}
					}
				}

				// backup(new File(result_folder), new File(target_path + "/Results"));
				log.info("Copying Maincontoller...");
				backup(new File(mainC_file), new File(target_path));
				log.info("Copying Unique No...");// Added for Base Claims 24/05/2023
				backup(new File(uniqueNo_file), new File(target_path));

				report.setTrasactionType("Backup");
				report.setStatus("Backup done");
				report.setMessage(target_path);
				ExcelUtility.writeReport(report);

				log.info("Backup Completed");
			}

			// BACKUP --End

			// Execution Summary Report & Auto report emailing

			if (Config.webOutlook_Email != null && Config.webOutlook_Email.equalsIgnoreCase("TRUE")) {
				// if (Config.webOutlook_Email.equalsIgnoreCase("TRUE")) {

				log.info("Initiating Execution Summary Report Utility - WebOutlook Version");
				// execution status utility macro call
				runExecutionSummaryReport();

				log.info("Taking a snapshot of the Execution Summary Report");
				try {
					summaryReportName = ReportScreenshot.printexcel();
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				log.info("Initiating WebOutlook Email...");
				try {
					SendMail.SendMailWebOutlook(target_path, summaryReportName);
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				log.info("WebOutlook Email Completed");
			} else if (Config.webOutlook_Email == null || Config.webOutlook_Email.equalsIgnoreCase("FALSE")) {
				log.info(
						"WebOutlook_Email parameter missing from config sheet.Execution Summary email initiated via Desktop Outlook App");
				// execution status utility macro call
				runExecutionSummaryReport();
				log.info("Execution Status Report sent via Desktop Outllook App");
			}

			killExcels();

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

	// Backup
	public static void backup(File sourceLocation, File targetLocation) throws IOException {
		if (sourceLocation.isDirectory()) {
			FileUtils.copyDirectory(sourceLocation, targetLocation);
		} else {
			FileUtils.copyFileToDirectory(sourceLocation, targetLocation);
		}
	}

	// Backup

	@SuppressWarnings("unused")
	public void DataInput(String filePath, String testcaseID, String transactionType, String transactionCode,
			String operationType) throws Exception {

		// log.info("filePath is:" + filePath);
		if ((transactionCode == null) || (transactionCode.isEmpty())) {
			transactionCode = transactionType;
		}

		log.info("TransactionType is :" + transactionCode);// 24/05/2023

		// if(operationType.equalsIgnoreCase("InputandVerify")&&!operationType.isEmpty())
		if (operationType.equalsIgnoreCase("InputandVerify") && !StringUtils.isEmpty(operationType.toString())) {
			// ExcelUtility.GetDataFromValues(null, filePath, testcaseID.toString(),
			// transactionType.toString(), cycleDate, operationType);
			// WebVerification.performVerification(transactionType, testcaseID, "",
			// cycleDate);
			ExcelUtility.GetDataFromValues(filePath, testcaseID.toString(), transactionType.toString());
			// WebVerification.performVerification(transactionType, testcaseID, "", "NA");
			WebVerification.performVerification(transactionType, testcaseID);

		} else if (operationType.equalsIgnoreCase("Capture") && !StringUtils.isEmpty(operationType)) {
			log.info("Capture");
			String OperationType = "Capture";
			WebVerification.performVerification(transactionType, testcaseID);
		}

		else if (!operationType.equalsIgnoreCase("Verify") && !StringUtils.isEmpty(operationType.toString())) {
			// ExcelUtility.GetDataFromValues(null, filePath, testcaseID.toString(),
			// transactionType.toString(), cycleDate, operationType);
			ExcelUtility.GetDataFromValues(filePath, testcaseID.toString(), transactionType.toString());
		}

		else if (!operationType.equalsIgnoreCase("Input") && !StringUtils.isEmpty(operationType.toString())) {
			log.info("INPUT");
			WebVerification.performVerification(transactionType, testcaseID);
		}
	}

	private static void runExecutionSummaryReport() {
		try {
			String statusUtility = Config.executionStatusReportUtility;
			// log.info("Status Utility : " + statusUtility);

			if (statusUtility == null) {
				log.info("Execution Status Report utility path not found in Config");
			} else {

				File execStatusFile = new File(statusUtility);
				if (execStatusFile.exists()) {
					log.info(
							"Execution Status Report macro is running, this migh take 10 -15 minutes depending upon the data ...");
					Jacob.main(statusUtility, "report");
					log.info("Completed");
					// log.info("Status Report sent");
				}

				else {
					log.info("Execution Status Report utility not found.");
				}
			}
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
		}
	}

}
