package com.majesco.itaf.main;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.List;
import javax.swing.JOptionPane;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.majesco.itaf.report.utils.VisualReport;
import com.majesco.itaf.util.BackupUtility;
import com.majesco.itaf.util.CalendarSnippet;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.util.Jacob;
import com.majesco.itaf.verification.WebVerification;

public class ITAFWebDriverSuite extends ITAFWebDriver {
	private final static Logger log = LogManager.getLogger(ITAFWebDriverSuite.class.getName());
	private MainController controller;
	public static boolean failure = false;

	ITAFWebDriverSuite() {
	}

	protected void init() throws IOException {

		String reportsheet = "";
		String detailreportsheet = "";
		String uniquenumbersheet = "";
		//String nodetailreportsheet = "";Sel4
		String FailedScreen = "";
		String body = "";
		String strCompleteData = "";
		String subject = "";
		try {
			MainControllerSuite.frmDate = new Date();
			subject = Config.projectName + " - Automated Test Run Results";
			email = Config.emailId;
			reportsheet = Config.resultOutput;
			detailreportsheet = Config.verificationResultPath;
			uniquenumbersheet = Config.transactionInfo;

			//cleanUp();//Mandar
			

			try {
				Automation.clearTempDirectory();
				log.info("Temp files deleted");
			} catch (Exception e) {
				log.info("Unable to clear temp directory: " + e.getMessage());
			}

			Automation.setUp();
			if (!ITAFWebDriverSuite.failure) {
				controller = ObjectFactory.getMainController();
				log.info("MainController File Path is :" + Config.controllerFilePath);
				controller.ControllerData(Config.controllerFilePath);
			}
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			try {
				// controller.pauseFun("Failure outside transaction : " + e.getMessage());
				controller.pauseFun(e.getMessage());
			} catch (Exception e1) {
				log.error(e1.getMessage(), e1);
				log.error("Issue while hanlding error. Check logs.");

			}
		} finally {

			log.info("\n==================================Post Execution Activities================================");

			// Call to various comparison utilities in Billing
			if (Config.webserviceComparison.equalsIgnoreCase("True")
					|| "True".equalsIgnoreCase(Config.invoiceComparision)) {

				report.setSilo("Billing");
				report.setIteration(Config.cycleNumber);
				report.setTestcaseId("Common");
				report.setCycleDate("Common");
				report.setTransactionVerificationType("NA");

				// call to xml\webservice comparison utility
				if (Config.webserviceComparison.equalsIgnoreCase("True")) {
					Date finalfrmDate = new Date();
					report.setTrasactionType("XML Comparison Macro");
					report.setFromDate(Config.dtFormat.format(finalfrmDate));

					String webserviceUtility = Config.webserviceComparisonUtilityPath;
					log.info("Webservice Comparison Utility Path: " + webserviceUtility);
					File webServiceUtilityFile = new File(webserviceUtility);
					if (webServiceUtilityFile.exists()) {
						log.info(webserviceUtility + " macro is running...");
						Jacob.main(webserviceUtility, "WebserviceVerification");
						threadSleep(5000);
					}

					File f1 = new File(Config.resultFilePath + Config.webServiceComparisonResultPath);
					if (f1.exists()) {
						log.info("WebService comparison results file created successfuly");
						report.setStatus("PASS");
						report.setMessage("WebService Verification Executed Successfully");
						ExcelUtility.writeReportSuite(report);
						threadSleep(2000);
						killExcels();
					} else {
						log.info("WebService Comparison Results file does not exist");
					}
				}

				// call to invoice comparison utility
				if ("True".equalsIgnoreCase(Config.invoiceComparision)) {

					Date finalfrmDate = new Date();
					report.setTrasactionType("Invoice Comparison Macro");
					report.setFromDate(Config.dtFormat.format(finalfrmDate));
					String invoiceXMLUtility = Config.invoiceXmlComparisionUtilityPath;
					log.info("Invoice XML comparision Utility : " + invoiceXMLUtility);

					File invoiceXMLUtilityFile = new File(invoiceXMLUtility);
					if (invoiceXMLUtilityFile.exists()) {
						log.info(invoiceXMLUtility + " macro is running...");
						Jacob.main(invoiceXMLUtility, "BillXMLVerification");
						threadSleep(5000);

					}
					// checkAgain
					File f1 = new File(Config.resultFilePath + "Invoice_Comparison_Results.xls");
					if (f1.exists()) {
						log.info("Invoice Comparison Results file exist ");
						report.setStatus("PASS");
						report.setMessage("Invoice Comparison Verification Executed Successfully");
						ExcelUtility.writeReportSuite(report);
						killExcels();
					} else {
						log.info("Invoice_Comparison_Results file does not exist ");
					}
				}
			}
			// Raw iTAF mail with attachments
			if (!email.toString().equalsIgnoreCase("NA") && !email.equalsIgnoreCase("") && email != null) {
				String tempReport = "", tempDetailedReport = "", tempUniqueNum = "", tempFailedScreen = "";

				File tempFile = new File(reportsheet);
				if (!tempFile.exists())
					tempReport = "";
				else
					tempReport = reportsheet;

				tempFile = new File(detailreportsheet);
				if (!tempFile.exists())
					tempDetailedReport = "";
				else
					tempDetailedReport = detailreportsheet;

				tempFile = new File(uniquenumbersheet);
				if (!tempFile.exists())
					tempUniqueNum = "";
				else
					tempUniqueNum = uniquenumbersheet;

				tempFile = new File(FailedScreen);
				if (!tempFile.exists())
					tempFailedScreen = "";
				else
					tempFailedScreen = FailedScreen;

				body = "Please find attached Report";
				strCompleteData = email + "##" + subject + "##" + body + "##" + tempReport + "##" + tempDetailedReport
						+ "##" + tempFailedScreen + "##" + tempUniqueNum;
				Runtime.getRuntime().exec("wscript SendMail_v2.vbs " + (char) 34 + strCompleteData + (char) 34);
				threadSleep(5000);
			}

			// Status report macro call
			try {
				String statusUtility = Config.executionStatusReportUtility;

				if (statusUtility == null || StringUtils.equals(statusUtility, "")) {
					log.info("Execution Status Tracker utility path not found in Config file");
				} else {

					File execStatusFile = new File(statusUtility);
					if (execStatusFile.exists()) {
						killExcels();
						threadSleep(3000);
						log.info("Execution Status Tracker utility Path: " + statusUtility);
						log.info("Execution Status Tracker utility is runnnig...");
						Jacob.main(statusUtility, "report");
						threadSleep(3000);
						log.info("Execution Status Report sent");
					} else {
						log.info("Execution Status Tracker utility not found : " + statusUtility);
					}
				}
			} catch (Exception e) {
				log.error(e.getMessage(), e);
			}

			// Visual report generation
			try {
				VisualReport.generateVisualReport();
				log.info("Please refer following sheet for visual report: " + Config.resultFilePath
						+ "VisualReport.xlsx");
			} catch (Exception e) {
				log.info("[OPTIONAL]Could not create visual report, because of following reason: " + e.getMessage());
				e.printStackTrace();
			}

			// Close open excels and related threads
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
				}
			}
			// Initializing Backup Process END

			threadSleep(200);
			log.info("\n==================================iTAF Execution Finished!=================================");

			// Completion pop-up in case of UserInteraction as True
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
	//Mandar

	/*
	 * public void DataInput(String structurePath, String filePath, String
	 * testcaseID, String transactionType, String transactionCode, String
	 * operationType, String cycleDate) throws Exception { if (transactionCode ==
	 * null) { transactionCode = transactionType; } log.info(transactionType);
	 * 
	 * if (MainControllerSuite.transactionSilo.equalsIgnoreCase("Billing")) { if
	 * (StringUtils.equalsIgnoreCase(operationType, "InputandVerify")) {
	 * ExcelUtility.GetDataFromValues(structurePath, filePath,
	 * testcaseID.toString(), transactionType.toString(), cycleDate, operationType);
	 * WebVerification.performVerificationBilling(transactionType, testcaseID,
	 * operationType, cycleDate); } else if
	 * (StringUtils.equalsIgnoreCase(operationType, "Capture")) {
	 * WebVerification.performVerificationBilling(transactionType, testcaseID,
	 * operationType, cycleDate); } else if
	 * (!StringUtils.equalsIgnoreCase(operationType, "Verify")) {
	 * ExcelUtility.GetDataFromValues(structurePath, filePath,
	 * testcaseID.toString(), transactionType.toString(), cycleDate, operationType);
	 * } else if (!StringUtils.equalsIgnoreCase(operationType, "Input")) {
	 * WebVerification.performVerificationBilling(transactionType, testcaseID, "",
	 * cycleDate); } } else if
	 * (MainControllerSuite.transactionSilo.equalsIgnoreCase("Claims")) { if
	 * (StringUtils.equalsIgnoreCase(operationType, "InputandVerify")) {
	 * ExcelUtility.GetDataFromValues(filePath, testcaseID.toString(),
	 * transactionType.toString());
	 * WebVerification.performVerification(transactionType, testcaseID); } else if
	 * (StringUtils.equalsIgnoreCase(operationType, "Capture")) {
	 * WebVerification.performVerification(transactionType, testcaseID); } else if
	 * (!StringUtils.equalsIgnoreCase(operationType, "Verify")) {
	 * ExcelUtility.GetDataFromValues(filePath, testcaseID.toString(),
	 * transactionType.toString()); } else if
	 * (!StringUtils.equalsIgnoreCase(operationType, "Input")) {
	 * WebVerification.performVerification(transactionType, testcaseID); } } else if
	 * (MainControllerSuite.transactionSilo.equalsIgnoreCase("DM")) { if
	 * (StringUtils.equalsIgnoreCase(operationType, "InputandVerify")) {
	 * ExcelUtility.GetDataFromValues(filePath, testcaseID.toString(),
	 * transactionType.toString());
	 * WebVerification.performVerification(transactionType, testcaseID); } else if
	 * (!StringUtils.equalsIgnoreCase(operationType, "Verify")) {
	 * ExcelUtility.GetDataFromValues(filePath, testcaseID.toString(),
	 * transactionType.toString()); } else if
	 * (!StringUtils.equalsIgnoreCase(operationType, "Input")) {
	 * WebVerification.performVerification(transactionType, testcaseID); } } else if
	 * (MainControllerSuite.transactionSilo.equalsIgnoreCase("PAS") ||
	 * MainControllerSuite.transactionSilo.equalsIgnoreCase("Digital1st")) { if
	 * (StringUtils.equalsIgnoreCase(operationType, "InputandVerify")) {
	 * ExcelUtility.GetDataFromValues(filePath, testcaseID.toString(),
	 * transactionType.toString());
	 * WebVerification.performVerification(transactionType, testcaseID); } else if
	 * (!StringUtils.equalsIgnoreCase(operationType, "Verify")) {
	 * ExcelUtility.GetDataFromValues(filePath, testcaseID.toString(),
	 * transactionType.toString()); } else if
	 * (!StringUtils.equalsIgnoreCase(operationType, "Input")) {
	 * WebVerification.performVerification(transactionType, testcaseID); } } }
	 */
	
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

	private void killExcels() {
		try {
			CalendarSnippet.killProcess("EXCEL.EXE");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void cleanUp() {
		// delete open excel threads and temp directory content
		try {
			CalendarSnippet.killProcess("EXCEL.EXE");
			Automation.clearTempDirectory();
		} catch (Exception e) {
			log.info("Unable to clear temp directory: " + e.getMessage());
		}

		// delete webservice comparison result file
		if (Config.webServiceComparisonResultPath != null) {
			File webServiceResultFile = new File(Config.webServiceComparisonResultPath);
			if (webServiceResultFile.exists())
				webServiceResultFile.delete();
		}

		if (Config.inputMasterFilePath != null && Config.inputDataFilePath != null) {
			try {
				File source = new File(Config.inputMasterFilePath);
				File dest = new File(Config.inputDataFilePath + "Billing");
				FileUtils.copyDirectory(source, dest);
				log.info("Files copied from InputMaster to 'Resources\\Input\\Billing' folder successfully");
			} catch (Exception e) {
				log.error("Issue while copying 'InputMaster' folder content to 'Resources\\Input\\Billing' folder. "
						+ e.getStackTrace());
				ITAFWebDriverSuite.failure = true;
			}
		}
		if (Config.inputMasterFilePath != null && Config.controllerFilePath != null) {
			try {
				String MainControllerFileName = Config.controllerFilePath
						.substring(Config.controllerFilePath.lastIndexOf("\\") + 1, Config.controllerFilePath.length());
				File source = new File(Config.inputMasterFilePath + MainControllerFileName);
				File dest = new File(
						Config.controllerFilePath.substring(0, Config.controllerFilePath.lastIndexOf("\\") + 1)
								+ MainControllerFileName);
				FileUtils.copyFile(source, dest);
				log.info("MainController copied from InputMaster to 'CommonResources' folder successfully");
			} catch (Exception e) {
				log.error("Issue while copying 'MainController' file from InputMaster to 'CommonResources' folder. "
						+ e.getStackTrace());
				ITAFWebDriverSuite.failure = true;
			}
		}
	}

	public List<File> listFiles(String directoryName, List<File> files) {
		File directory = new File(directoryName);

		// Get all files from a directory.
		File[] fList = directory.listFiles();
		if (fList != null)
			for (File file : fList) {
				if (file.isFile()) {
					files.add(file);
				} else if (file.isDirectory()) {
					listFiles(file.getAbsolutePath(), files);
				}
			}
		return files;
	}

	private void threadSleep(int i) {
		try {
			Thread.sleep(i);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}
}
