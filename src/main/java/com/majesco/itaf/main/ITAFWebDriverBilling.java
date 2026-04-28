package com.majesco.itaf.main;

import java.io.File;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;
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

public class ITAFWebDriverBilling extends ITAFWebDriver {

	private final static Logger log = LogManager.getLogger(ITAFWebDriverBilling.class.getName());
	private MainController controller;

	ITAFWebDriverBilling() {
	}

	protected void init() throws IOException {
		String webserviceUtility = "";
		String restfulserviceUtility = "";
		String openAPIserviceUtility = "";
		String summaryReportName = null;

		try {

			try {
				Automation.clearTempDirectory();
				log.info("Temp files deleted");
			} catch (Exception e) {
				log.info("Unable to clear temp directory: " + e.getMessage());
			}

			Automation.setUp();

			controller = ObjectFactory.getMainController();

			email = Config.emailId;
			// reportsheet = Config.resultOutput;//Sel4
			// detailreportsheet = Config.verificationResultPath;//Sel4

			log.info("MainController File Path is :" + Config.controllerFilePath);
			MainControllerBilling.frmDateNew = new Date();

			TransactionMapping.TransactionInputData("Login"); // Default Login
			// START
			WebHelperUtil.saveScreenShot();

			// Meghna- For Suite Testing//
			if (Config.suiteTesting.equalsIgnoreCase("TRUE")) {
				SuiteDriver.start();
			}

			controller.ControllerData(Config.controllerFilePath); // Start
			// Transaction

		} catch (Exception e) {
			log.error(e.getMessage(), e);
			log.error(e.getLocalizedMessage());
			report.setStatus("FAIL");
			report.setTestcaseId(controller.controllerTestCaseID.toString());
			report.setTrasactionType(controller.controllerTransactionType.toString());
			try {
				controller.pauseFun("d: " + controller.controllerTestCaseID + ", Tranasction: "
						+ controller.controllerTransactionType + ", Error: " + e.getMessage());
			} catch (Exception e1) {
				log.error(e.getMessage(), e);
				controller.pauseFun("File Not Found");
			}
		} finally {
			// TM: 16-01-2015
			try {
				if (com.majesco.itaf.util.CalendarSnippet.isProcessRunning("IEDriverServer.exe"))
					com.majesco.itaf.util.CalendarSnippet.killProcess("IEDriverServer.exe");
			} catch (Exception e) {
				log.error(e.getMessage(), e);
			}
			// Added keyword END for SUITE Integration//
			if (controller.controllerTransactionType != null) {
				if (controller.controllerTransactionType.toString().equalsIgnoreCase("END")) {
					ExcelUtility.writeReport(report);
					System.exit(0);
				}
			}
			// END

			Date finalfrmDate = new Date();
			report.setIteration(Config.cycleNumber);
			report.setTestcaseId("Common");
			report.setCycleDate("Common");
			report.setTrasactionType("XML Comparison Macro");
			report.setFromDate(Config.dtFormat.format(finalfrmDate));
			// report.setStrGroupName(MainController.controllerGroupName.toString());
			// ***

			if (Config.webserviceComparison.equalsIgnoreCase("True")) {
				webserviceUtility = Config.webserviceComparisonUtilityPath;
				log.info("webserviceUtility Location : " + webserviceUtility);
				// threadSleep(5000);

				File webServiceUtilityFile = new File(webserviceUtility);
				if (webServiceUtilityFile.exists()) {
					log.info(webserviceUtility + " Macro is running...");
					Jacob.main(webserviceUtility, "WebserviceVerification");
					log.info("Completed");
				}

				File f1 = new File(Config.resultFilePath + "WebService_Comparison_Results.xls");
				if (f1.exists()) {
					log.info("WebService_Comparison_Results file exist");
					report.setStatus("PASS");
					report.setMessage("WebService Verification Executed Successfully");
					ExcelUtility.writeReport(report);
					killExcels();
				} else {
					log.info("WebService_Comparison_Results file does not exist");
				}
			}

			if ("True".equalsIgnoreCase(Config.restfulserviceComparison)) {
				restfulserviceUtility = Config.restfulServiceComparisonUtilityPath;
				log.info("restfulserviceUtility : " + restfulserviceUtility);
				threadSleep(5000);

				File restUtilityFile = new File(restfulserviceUtility);
				if (restUtilityFile.exists()) {
					log.info(restfulserviceUtility
							+ " Macro is running, this migh take 10-15 minutes depending upon the data ...");
					Jacob.main(restfulserviceUtility, "restfulserviceVerification");
					log.info("Completed");
				}

				File f1 = new File(Config.resultFilePath + "RestfulService_Comparison_Results.xls");
				if (f1.exists()) {
					log.info("RestfulService_Comparison_Results file exist");
					report.setStatus("PASS");
					report.setMessage("RestfulService Verification Executed Successfully");
					ExcelUtility.writeReport(report);
					killExcels();
				} else {
					log.info("RestfulService_Comparison_Results file does not exist");
				}
			}
			if ("True".equalsIgnoreCase(Config.openAPIserviceComparison)) {
				openAPIserviceUtility = Config.openAPIServiceComparisonUtilityPath;
				log.info("openAPIserviceUtility : " + openAPIserviceUtility);
				threadSleep(5000);

				File openAPIUtilityFile = new File(openAPIserviceUtility);
				if (openAPIUtilityFile.exists()) {
					log.info(openAPIserviceUtility
							+ " Macro is running, this migh take 10-15 minutes depending upon the data ...");
					Jacob.main(openAPIserviceUtility, "openAPIserviceVerification");
					log.info("Completed");
				}

				File f1 = new File(Config.resultFilePath + "OpenAPIService_Comparison_Results.xls");
				if (f1.exists()) {
					log.info("OpenAPIService_Comparison_Results file exist ");
					report.setStatus("PASS");
					report.setMessage("OpenAPIService Verification Executed Successfully");
					ExcelUtility.writeReport(report);
					killExcels();
				} else {
					log.info("OpenAPIService_Comparison_Results file does not exist ");
				}
			}

			// added to run macro invoice comparsion utility
			if ("True".equalsIgnoreCase(Config.invoiceComparision)) {
				String invoiceXMLUtility = Config.invoiceXmlComparisionUtilityPath;
				log.info("Invoice XML comparision Utility : " + invoiceXMLUtility);
				threadSleep(5000);

				File invoiceXMLUtilityFile = new File(invoiceXMLUtility);
				if (invoiceXMLUtilityFile.exists()) {
					log.info(invoiceXMLUtility + " macro is running...");
					Jacob.main(invoiceXMLUtility, "BillXMLVerification");
				}

				File f1 = new File(Config.resultFilePath + "Invoice_Comparison_Results.xls");
				if (f1.exists()) {
					log.info("Invoice Comparison Results file exist ");
					report.setStatus("PASS");
					report.setMessage("Invoice Comparison Verification Executed Successfully");
					ExcelUtility.writeReport(report);
					killExcels();
				} else {
					log.info("Invoice_Comparison_Results file does not exist ");
				}
			}
			// BACKUP --
			log.info("Backup Started...");
			String target_path = Config.backupFilepath;
			File targetPathBCK = new File(target_path);

			if (!targetPathBCK.exists()) {
				targetPathBCK.mkdir();
			}

			if (!(target_path.equals(""))) {
				String input_folder = Config.inputDataFilePath;
				String result_folder = Config.resultFilePath;
				String mainC_file = Config.controllerFilePath;
				String uniqueNo_file = Config.transactionInfo; // Added for Base Billing 06/04/2023

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
				// Sel4
				// backup(new File(result_folder), new File(target_path + "/Results"));

				File[] resultFolderContents = new File(result_folder).listFiles();
				if (resultFolderContents != null) {
					for (File file : resultFolderContents) {
						// Check if the item is a directory and its name is not in the exclusion list
						if (file.isDirectory() && !(file.getName().equals("AutomatedBackup")
								|| file.getName().equals("ExecutionStatusReportLogs")
								|| file.getName().equals("Outlook")
								|| file.getName().equals("ChangeBusinessdate_Log")
								|| file.getName().equals("XMLOutput"))) {
							// Copy the entire folder and its contents recursively
							backup(file, new File(target_path + "/Results/" + file.getName()));
						} else if (!(file.getName().equals("AutomatedBackup")
								|| file.getName().equals("ExecutionStatusReportLogs")
								|| file.getName().equals("Outlook")
								|| file.getName().equals("ChangeBusinessdate_Log")
								|| file.getName().equals("XMLOutput"))) {
							// If it's a file, or a directory to be excluded, copy it directly
							backup(file, new File(target_path + "/Results/"));
						}
					}
				}

				log.info("Copying Maincontoller...");
				backup(new File(mainC_file), new File(target_path));
				log.info("Copying Unique No...");// Added for Base Billing 06/04/2023
				backup(new File(uniqueNo_file), new File(target_path));

				report.setTrasactionType("Backup");
				report.setStatus("Backup done");
				report.setMessage(target_path);
				ExcelUtility.writeReport(report);

				log.info("Backup Completed");
			}

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

	public static void backup(File sourceLocation, File targetLocation) throws IOException {
		if (sourceLocation.isDirectory()) {
			FileUtils.copyDirectory(sourceLocation, targetLocation);
		} else {
			FileUtils.copyFileToDirectory(sourceLocation, targetLocation);
		}
	}

	@SuppressWarnings("unused")
	public void DataInput(String structurePath, String filePath, String testcaseID, String transactionType,
			String transactionCode, String operationType, String cycleDate) throws Exception {
		log.info("filePath is:" + filePath);
		if (transactionCode == null) {
			transactionCode = transactionType;
		}
		log.info(transactionCode);
		if (transactionType.equalsIgnoreCase("CommissionExtraction") && cycleDate.equalsIgnoreCase("02/06/2012")) {
			System.out.println("HI");
		}

		// if(operationType.equalsIgnoreCase("InputandVerify")&&!operationType.isEmpty())
		if (operationType.equalsIgnoreCase("InputandVerify") && !StringUtils.isEmpty(operationType.toString())) {
			ExcelUtility.GetDataFromValues(structurePath, filePath, testcaseID.toString(), transactionType.toString(),
					cycleDate, operationType);

			if (ITAFWebDriver.isBillingApplication()) {
				WebVerification.performVerificationBilling(transactionType, testcaseID, operationType, cycleDate);
			} else {
				WebVerification.performVerification(transactionType, testcaseID);
			}

		} else if (operationType.equalsIgnoreCase("Capture") && !StringUtils.isEmpty(operationType)) {
			log.info("Capture");
			String OperationType = "Capture";

			if (ITAFWebDriver.isBillingApplication()) {
				WebVerification.performVerificationBilling(transactionType, testcaseID, operationType, cycleDate);
			} else {
				WebVerification.performVerification(transactionType, testcaseID);
			}
		} else if (!operationType.equalsIgnoreCase("Verify") && !StringUtils.isEmpty(operationType.toString())) {
			ExcelUtility.GetDataFromValues(structurePath, filePath, testcaseID.toString(), transactionType.toString(),
					cycleDate, operationType);
		} else if (!operationType.equalsIgnoreCase("Input") && !StringUtils.isEmpty(operationType.toString())) {
			log.info("INPUT");

			if (ITAFWebDriver.isBillingApplication()) {
				WebVerification.performVerificationBilling(transactionType, testcaseID, "", cycleDate);
			} else {
				WebVerification.performVerification(transactionType, testcaseID);
			}
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
							" Execution Status Report macro is running, this migh take 10 -15 minutes depending upon the data ...");
					Jacob.main(statusUtility, "report");
					log.info("Execution Status Report generated");
					// log.info("Status Report sent");
				}

				else {
					System.out.println("Execution Status Report utility not found.");
				}
			}
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
		}
	}

}
