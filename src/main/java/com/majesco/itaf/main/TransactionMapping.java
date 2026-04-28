package com.majesco.itaf.main;

import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Sheet;
import com.majesco.itaf.batch.RunBatch;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.util.PDFComparisonUtil;
import com.majesco.itaf.vo.Reporter;

public class TransactionMapping {

	private final static Logger log = LogManager.getLogger(TransactionMapping.class.getName());
	public static String operationType = "";
	public static String directoryPathFileUpload = "";
	private static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	private static MainController controller = ObjectFactory.getMainController();
	public static Boolean isTransactionStructureCommon = false;
	
	//For Billing
	public static Reporter TransactionInputData(String cycleDate, String controllerTestCaseID,
			String controllerTransactionType, String filePath) throws IOException, Exception {
		Reporter report = new Reporter();
		Sheet workSheet = null;
		HashMap<String, Integer> inputHashTable = new HashMap<>();
		try {
			if (ITAFWebDriver.isClaimsApplication()) {
				workSheet = ExcelUtility.GetSheet(Config.transactionInputFilePath, "Web_Transaction_Input_Files");
			} else {
				workSheet = ExcelUtility.GetSheet(Config.transactionInputFilePath, "Web_Transaction_Input_Files");
				log.info("TRANSACTION INPUT FILE PATH : " + Config.transactionInputFilePath);
			}
		} catch (IOException ioe) {
			log.error("Failed to access TRANSACTION_INPUT_FILEPATH <-|-> LocalizeMessage " + ioe.getLocalizedMessage()
					+ " <-|-> Message " + ioe.getMessage() + " <-|-> Cause " + ioe.getCause(), ioe);
			throw new Exception(
					"Failed to access TRANSACTION_INPUT_FILEPATH   <-|-> LocalizeMessage " + ioe.getLocalizedMessage()
							+ " <-|-> Message" + ioe.getMessage() + " <-|-> Cause " + ioe.getCause());
		}

		int rowCount = workSheet.getLastRowNum() + 1;
		for (int rowIndex = 1; rowIndex < rowCount && !controller.pauseExecution; rowIndex++) {
			String transactionCode = WebHelperUtil.getCellData("TransactionCode", workSheet, rowIndex, inputHashTable);
			String transactionType1 = WebHelperUtil.getCellData("TransactionType", workSheet, rowIndex, inputHashTable);// dev
			String directoryPath = WebHelperUtil.getCellData("DirPath", workSheet, rowIndex, inputHashTable);
			String inputExcel = WebHelperUtil.getCellData("InputSheet", workSheet, rowIndex, inputHashTable);

			boolean rowFound = false;
			String[] transactionTypeArr = transactionType1.split(",");
			int transactionTypeArrSize = transactionTypeArr.length;

			for (int Introwindex = 0; Introwindex <= transactionTypeArrSize - 1; Introwindex++) {
				String transactionType = transactionTypeArr[Introwindex];
				if (transactionTypeArrSize > 1) {
				}

				if (transactionType.equalsIgnoreCase(controllerTransactionType)) {
					if (transactionCode != null && directoryPath == null
							&& controllerTransactionType.equalsIgnoreCase(transactionType)) {
						report.setInputPath("");
						report.setOperationType("");
						report.setTransactioncode(transactionCode);

						webDriver.getReport().setTransactioncode(transactionCode);
						webDriver.getReport().setTestcaseId(controllerTestCaseID);
						webDriver.DataInput("", "", controllerTestCaseID, transactionType, transactionCode, "",
								cycleDate);// Meghna:R10.10-For Common Structure Sheet
						rowFound = true;
						break;
					}

					if (!transactionType.startsWith("Verify")) {
						operationType = "Input";
					}
					if (transactionType.startsWith("Verify") && (!StringUtils.isEmpty(directoryPath))
							&& (!StringUtils.isEmpty(inputExcel)) && (StringUtils.isEmpty(transactionCode))) {
						operationType = "InputandVerify";
						// log.info("InputandVerify");

					}
					if (transactionType.startsWith("Verify") && (transactionCode.startsWith("Verify"))
							&& (!StringUtils.isEmpty(directoryPath)) && (!StringUtils.isEmpty(inputExcel))) {
						operationType = "InputandCapture";

					} else if (transactionType.startsWith("Verify") && (StringUtils.isEmpty(directoryPath))
							&& (StringUtils.isEmpty(inputExcel))) {
						operationType = "Verify";
					}

					log.info("operationType is:" + operationType);
					if (controllerTransactionType.equalsIgnoreCase(transactionType)) {
						if ((directoryPath == null || inputExcel == null) && operationType != "Verify") {
							controller.pauseFun("Please Enter the directory or excelsheet name");
						} else {
							String inputFilePath = null;
							String structurePath = null;
							if (operationType != "Verify") {
								inputFilePath = Config.inputDataFilePath + directoryPath + "\\" + inputExcel;
								directoryPathFileUpload = directoryPath; 
								if (ITAFWebDriver.isBillingApplication()) {
									structurePath = Config.structureSheetFilePath + directoryPath + "\\" + inputExcel;
								}
							}
							log.info(inputFilePath);
							report.setInputPath(inputFilePath);
							report.setOperationType(operationType);
							report.setTransactioncode(transactionCode);

							webDriver.getReport().setTransactioncode(transactionCode);
							webDriver.getReport().setTestcaseId(controllerTestCaseID);

							if (ITAFWebDriver.isClaimsApplication()) {
								webDriver.DataInput(inputFilePath, controllerTestCaseID, transactionType,
										transactionCode, operationType, cycleDate);
							} else if (ITAFWebDriver.isBillingApplication()) {
								webDriver.DataInput(structurePath, inputFilePath, controllerTestCaseID, transactionType,
										transactionCode, operationType, cycleDate);
							}
							rowFound = true;
							break;
						}
					}
				} else if (!transactionType.equalsIgnoreCase(controllerTransactionType) && rowIndex == rowCount - 1) {
					controller.pauseFun("Transaction " + controller.controllerTransactionType + " Not Found");
					ExcelUtility.writeReport(webDriver.getReport());
				}
			} 
				
			if (rowFound == true)
				break;
			
		}
		return null;
	}

	//For Billing
	public static void TransactionInputData(String controllerTransactionType)
			throws IOException, NullPointerException, Exception {
		try {
			HashMap<String, Integer> inputHashTable1 = new HashMap<>();
			Sheet workSheet;

			/*
			 * String normalizedPath =
			 * Config.transactionInputFilePath.replace("\\\\", "\\");
			 * log.info("Normalized File Path: " + normalizedPath);
			 */

			if (ITAFWebDriver.isClaimsApplication()) {
				workSheet = ExcelUtility.GetSheet(Config.transactionInputFilePath, "Web_Transaction_Input_Files");
			} else {
				workSheet = ExcelUtility.GetSheet(Config.transactionInputFilePath, "Web_Transaction_Input_Files");
			}

			int rowCount1 = workSheet.getLastRowNum() + 1;
			log.info("Transaction Name in MainController : " + controllerTransactionType);

			webDriver.getReport().setTestcaseId("Common");

			if (controller.cycleDateCellValue == ("") || controller.cycleDateCellValue == null) {
				webDriver.getReport().setCycleDate("Common");
			} else {
				webDriver.getReport().setCycleDate(controller.cycleDateCellValue);
			}

			webDriver.getReport().setTrasactionType(controllerTransactionType);
			Date toDate = new Date();
			webDriver.getReport().setToDate(Config.dtFormat.format(toDate));

			for (int rowIndex = 1; rowIndex < rowCount1; rowIndex++) {
				
				webDriver.getReport().setMessage("");
				webDriver.getReport().setScreenShot("");

				String transactionCode1 = WebHelperUtil.getCellData("TransactionCode", workSheet, rowIndex,
						inputHashTable1);
				String transactionType1 = WebHelperUtil.getCellData("TransactionType", workSheet, rowIndex,
						inputHashTable1);
				String directoryPath1 = WebHelperUtil.getCellData("DirPath", workSheet, rowIndex, inputHashTable1);
				String inputExcel1 = WebHelperUtil.getCellData("InputSheet", workSheet, rowIndex, inputHashTable1);
				// log.info("Transaction Name in transaction mapping : "+transactionType1);

				if (transactionType1.equalsIgnoreCase(controllerTransactionType)) {
					if (inputExcel1 == null || StringUtils.isEmpty(inputExcel1)) {
						controller.pauseFun("Please Enter the Input excel sheet name");
					}
					if ((directoryPath1 != null && inputExcel1 != null)
							|| (!StringUtils.isEmpty(directoryPath1) && !StringUtils.isEmpty(inputExcel1))) {
						operationType = "Input";
						String inputFilePath1 = null;
						inputFilePath1 = Config.inputDataFilePath + directoryPath1 + "\\" + inputExcel1;

						if (ITAFWebDriver.isClaimsApplication()) {
							webDriver.DataInput(inputFilePath1, "", transactionType1, transactionCode1, operationType,
									"");
						} else if (ITAFWebDriver.isBillingApplication()) {
							// Meghna:R10.10-For Common Structure Sheet
							String inputFilePath_structure = null;
							inputFilePath_structure = Config.structureSheetFilePath + directoryPath1 + "\\"
									+ inputExcel1;
							webDriver.DataInput(inputFilePath_structure, inputFilePath1, "", transactionType1,
									transactionCode1, operationType, "");
						}
						break;
					}
					if ((directoryPath1 == null && inputExcel1 != null)
							|| (StringUtils.isEmpty(directoryPath1) && !StringUtils.isEmpty(inputExcel1))) {
						operationType = "Input";
						String inputFilePath2 = null;
						inputFilePath2 = Config.inputDataFilePath + inputExcel1;

						if (ITAFWebDriver.isClaimsApplication()) {
							webDriver.DataInput(inputFilePath2, "", transactionType1, transactionCode1, operationType,
									"");
						} else if (ITAFWebDriver.isBillingApplication()) {
							
							String inputFilePath_structure = null;
							inputFilePath_structure = Config.structureSheetFilePath + inputExcel1;
							webDriver.DataInput(inputFilePath_structure, inputFilePath2, "", transactionType1,
									transactionCode1, operationType, "");
						}
						break;
					}

				} else if (!transactionType1.equalsIgnoreCase(controllerTransactionType) && rowIndex == rowCount1 - 1) {
					controller.pauseFun("Transaction " + transactionType1 + " Not Found");
				}
			}
		} catch (IOException ne) {
			log.error("Failed to access TRANSACTION_INPUT_FILEPATH <-|-> LocalizeMessage " + ne.getLocalizedMessage()
					+ " <-|-> Message " + ne.getMessage() + " <-|-> Cause " + ne.getCause(), ne);
			throw new Exception("Failed to access TRANSACTION_INPUT_FILEPATH   <-|-> LocalizeMessage "
					+ ne.getLocalizedMessage() + " <-|-> Message" + ne.getMessage() + " <-|-> Cause " + ne.getCause());
		} catch (NullPointerException ne) {
			log.error("Failed to access TRANSACTION_INPUT_FILEPATH <-|-> LocalizeMessage " + ne.getLocalizedMessage()
					+ " <-|-> Message " + ne.getMessage() + " <-|-> Cause " + ne.getCause(), ne);
			throw new Exception("Failed to access TRANSACTION_INPUT_FILEPATH   <-|-> LocalizeMessage "
					+ ne.getLocalizedMessage() + " <-|-> Message" + ne.getMessage() + " <-|-> Cause " + ne.getCause());
		} catch (Exception e) {
			log.error("Failed while access TRANSACTION_INPUT_FILEPATH  <-|-> LocalizeMessage " + e.getLocalizedMessage()
					+ " <-|-> Message " + e.getMessage() + " <-|-> Cause " + e.getCause(), e);
			throw new Exception("Failed while access TRANSACTION_INPUT_FILEPATH  <-|-> LocalizeMessage "
					+ e.getLocalizedMessage() + " <-|-> Message" + e.getMessage() + " <-|-> Cause " + e.getCause());
		}
	}

	public static Reporter TransactionCaptureData(String cycleDate, String controllerTestCaseID,
			String controllerTransactionType, String filePath) throws Exception {
		Reporter report = new Reporter();
		HashMap<String, Integer> inputHashTable = new HashMap<>();
		Sheet workSheet;

		if (ITAFWebDriver.isClaimsApplication()) {
			workSheet = ExcelUtility.GetSheet(Config.transactionInputFilePath, "Web_Transaction_Input_Files");
		} else {
			workSheet = ExcelUtility.GetSheet(Config.transactionInputFilePath, "Web_Transaction_Input_Files");
		}

		int rowCount = workSheet.getLastRowNum() + 1;
		for (int rowIndex = 1; rowIndex < rowCount && !controller.pauseExecution; rowIndex++) {
			String transactionCode = WebHelperUtil.getCellData("TransactionCode", workSheet, rowIndex, inputHashTable);
			String transactionType = WebHelperUtil.getCellData("TransactionType", workSheet, rowIndex, inputHashTable);
			String directoryPath = WebHelperUtil.getCellData("DirPath", workSheet, rowIndex, inputHashTable);
			String inputExcel = WebHelperUtil.getCellData("InputSheet", workSheet, rowIndex, inputHashTable);
			if (transactionType.equalsIgnoreCase(controllerTransactionType)) {
				if (transactionCode != null && directoryPath == null
						&& controllerTransactionType.equalsIgnoreCase(transactionType)) {
					report.setInputPath("");
					report.setOperationType("");
					report.setTransactioncode(transactionCode);
					if (ITAFWebDriver.isClaimsApplication()) {
						webDriver.DataInput("", controllerTestCaseID, transactionType, transactionCode, "", cycleDate);
					} else if (ITAFWebDriver.isBillingApplication()) {
						webDriver.DataInput("", "", controllerTestCaseID, transactionType, transactionCode, "",
								cycleDate);
					}

					break;
				}

				if (!transactionType.startsWith("Verify")) {
					operationType = "Input";
				}
				if (transactionType.startsWith("Verify") && (!StringUtils.isEmpty(directoryPath))
						&& (!StringUtils.isEmpty(inputExcel)) && (StringUtils.isEmpty(transactionCode))) {
					operationType = "InputandVerify";
					log.info("InputandVerify");
				} else if (transactionType.startsWith("Verify") && (StringUtils.isEmpty(directoryPath))
						&& (StringUtils.isEmpty(inputExcel))) {
					operationType = "Capture";
				}

				log.info("operationType is:" + operationType);
				if (controllerTransactionType.equalsIgnoreCase(transactionType)) {
					if ((directoryPath == null || inputExcel == null) && operationType != "Capture") {
						controller.pauseFun("Please Enter the directory or excelsheet name");
					} else {
						String inputFilePath = null;

						// if(operationType != "Capture")
						// {
						inputFilePath = Config.inputDataFilePath + directoryPath + "\\" + inputExcel;
						// }
						log.info(inputFilePath);
						report.setInputPath(inputFilePath);
						report.setOperationType(operationType);
						report.setTransactioncode(transactionCode);

						if (ITAFWebDriver.isClaimsApplication()) {
							webDriver.DataInput(inputFilePath, controllerTestCaseID, transactionType, transactionCode,
									operationType, cycleDate);

						} else if (ITAFWebDriver.isBillingApplication()) {
							String structurePath = null;
							structurePath = Config.structureSheetFilePath + directoryPath + "\\" + inputExcel;
							webDriver.DataInput(structurePath, inputFilePath, controllerTestCaseID, transactionType,
									transactionCode, operationType, cycleDate);
						}

						break;
					}
				}
			} else if (!transactionType.equalsIgnoreCase(controllerTransactionType) && rowIndex == rowCount - 1) {
				controller.pauseFun("Transaction " + controller.controllerTransactionType + " Not Found");
				ExcelUtility.writeReport(webDriver.getReport());
			}
		}
		return null;
	}
	
	//For others - PAS...
	public static Reporter TransactionInputData(String controllerTestCaseID, String controllerTransactionType,
			String filePath) throws Exception {
		Reporter report = new Reporter();
		HashMap<String, Integer> inputHashTable = new HashMap<String, Integer>();
		Sheet workSheet = ExcelUtility.GetSheet(Config.transactionInputFilePath, "Web_Transaction_Input_Files");
		int rowCount = workSheet.getLastRowNum() + 1;
		isTransactionStructureCommon = false;
		for (int rowIndex = 1; rowIndex < rowCount && !controller.pauseExecution; rowIndex++) {
			// log.info("getting again in to WHutil to check trx code....");
			String transactionCode = WebHelperUtil.getCellData("TransactionCode", workSheet, rowIndex, inputHashTable);
			String transactionType = WebHelperUtil.getCellData("TransactionType", workSheet, rowIndex, inputHashTable);
			// log.info("transactionType:" + transactionType);
			String directoryPath = WebHelperUtil.getCellData("DirPath", workSheet, rowIndex, inputHashTable);
			// log.info("directoryPath:" + directoryPath);
			String inputExcel = WebHelperUtil.getCellData("InputSheet", workSheet, rowIndex, inputHashTable);
			// log.info("inputExcel:" + inputExcel);
			String transactionStructureType = WebHelperUtil.getCellData("TransactionStructureType", workSheet, rowIndex,
					inputHashTable);
			// log.info("transactionStructureType:" + transactionStructureType);

			// PAS+DM Suite implementation
			String silo = WebHelperUtil.getCellData("Silo", workSheet, rowIndex, inputHashTable);
			if (StringUtils.equalsIgnoreCase(Config.runMode, "suite")) {
				if (!silo.trim().equalsIgnoreCase("")) {
					ITAFWebDriver.setApplicationName(silo);
				} else {
					ITAFWebDriver.setApplicationName(ITAFWebDriver.getSuitePrimaryApplicationName());
				}
				ObjectFactory.setWebDriver();
				
				webDriver = ITAFWebDriver.getInstance();
				MainControllerDM.webDriver = ITAFWebDriver.getInstance();
				MainControllerPAS.webDriver = ITAFWebDriver.getInstance();
				WebHelperDM.webDriver = ITAFWebDriver.getInstance();
				WebHelperPAS.webDriver = ITAFWebDriver.getInstance();
				WebHelper.webDriver = ITAFWebDriver.getInstance();
				PDFComparisonUtil.webDriverPDF = ITAFWebDriver.getInstance();
				ExcelUtility.webDriver = ITAFWebDriver.getInstance();
				RunBatch.webDriver = ITAFWebDriver.getInstance();
			}
			// End of PAS+DM Suite Suite implementation

			if (transactionStructureType.equalsIgnoreCase("CommonValues")) {
				isTransactionStructureCommon = true;
			} else {
				isTransactionStructureCommon = false;
			}
			if (transactionType.equalsIgnoreCase(controllerTransactionType)) {
				if (transactionCode != null && directoryPath == null
						&& controllerTransactionType.equalsIgnoreCase(transactionType)) {
					report.setInputPath("");
					report.setOperationType("");
					report.setTransactioncode(transactionCode);
					webDriver.DataInput("", controllerTestCaseID, transactionType, transactionCode, "");
					break;
				}

				if (!transactionType.startsWith("Verify")) {
					operationType = "Input";
				}

				if (transactionType.startsWith("Verify") && (!directoryPath.isEmpty()) && (!inputExcel.isEmpty())) {
					operationType = "InputandVerify";
					System.out.println("--------------------------InputandVerify--------------------------");
					System.out.println("--------------------------InputandVerify--------------------------");
				} else if (transactionType.startsWith("Verify") && (directoryPath.isEmpty())
						&& (inputExcel.isEmpty())) {
					operationType = "Verify";
				}

				log.info("operationType is:" + operationType);
				if (controllerTransactionType.equalsIgnoreCase(transactionType)) {
					if ((directoryPath == null || inputExcel == null) && operationType != "Verify") {
						controller.pauseFun("Please Enter the directory or excelsheet name");
					} else {
						String inputFilePath = null;
						if (operationType != "Verify") {
							inputFilePath = Config.inputDataFilePath + directoryPath + "\\" + inputExcel;
							directoryPathFileUpload = directoryPath; // Added
							// For
							// FileUpload
						}
						if (!ITAFWebDriver.isClaimsApplication()|| (!ITAFWebDriver.isSuiteApplication()))//Selenium4 - Suite changes - 30/12/2024
							
							//System.out.println(inputFilePath);
						report.setInputPath(inputFilePath);
						report.setOperationType(operationType);
						report.setTransactioncode(transactionCode);
						webDriver.DataInput(inputFilePath, controllerTestCaseID, transactionType, transactionCode,
								operationType);
						break;
					}
				}
			} else if (!transactionType.equalsIgnoreCase(controllerTransactionType) && rowIndex == rowCount - 1) {
				controller.pauseFun("Transaction " + controller.controllerTransactionType + " Not Found");
				ExcelUtility.writeReportPAS(webDriver.getReport());
			}
		}
		return null;
	}

	public static Reporter TransactionInputDataSuite(String cycleDate, String controllerTestCaseID,
			String controllerTransactionType) throws Exception {
		Reporter report = new Reporter();
		HashMap<String, Integer> inputHashTable = new HashMap<String, Integer>();
		Sheet workSheet = ExcelUtility.GetSheet(Config.transactionInputFilePath, "Web_Transaction_Input_Files");
		int rowCount = workSheet.getLastRowNum() + 1;
		isTransactionStructureCommon = false;
		for (int rowIndex = 1; rowIndex < rowCount && !controller.pauseExecution; rowIndex++) {
			String transactionCode = WebHelperUtil.getCellData("TransactionCode", workSheet, rowIndex, inputHashTable);
			String transactionTypeTemp = WebHelperUtil.getCellData("TransactionType", workSheet, rowIndex,
					inputHashTable);
			String directoryPath = WebHelperUtil.getCellData("DirPath", workSheet, rowIndex, inputHashTable);
			String inputExcel = WebHelperUtil.getCellData("InputSheet", workSheet, rowIndex, inputHashTable);
			String transactionStructureType = WebHelperUtil.getCellData("TransactionStructureType", workSheet, rowIndex,
					inputHashTable);
			String silo = WebHelperUtil.getCellData("Silo", workSheet, rowIndex, inputHashTable);
			String verificationType = WebHelperUtil.getCellData("TransactionVerificationType", workSheet, rowIndex,
					inputHashTable);
			boolean rowFound = false;
			String[] transactionTypeArr = transactionTypeTemp.split(",");
			int transactionTypeArrSize = transactionTypeArr.length;

			for (int Introwindex = 0; Introwindex <= transactionTypeArrSize - 1; Introwindex++) {
				String transactionType = transactionTypeArr[Introwindex];

				if (transactionType.equalsIgnoreCase(controllerTransactionType)) {
					rowFound = true;

					if (verificationType.trim().equals(""))
						MainControllerSuite.transactionVerificationType = "NA";
					else
						MainControllerSuite.transactionVerificationType = verificationType;

					if (!silo.trim().equalsIgnoreCase("")) {
						MainControllerSuite.transactionSilo = silo;
						MainControllerSuite.setSiloParameters(silo);
					} else {
						log.error("Silo value not provided in " + Config.transactionInputFilePath);
						controller.pauseFun("Silo value not provided in " + Config.transactionInputFilePath);
						break;
					}
					if (transactionStructureType.equalsIgnoreCase("CommonValues"))
						isTransactionStructureCommon = true;
					else
						isTransactionStructureCommon = false;

					if (transactionCode != null && directoryPath == null
							&& controllerTransactionType.equalsIgnoreCase(transactionType)) {
						report.setInputPath("");
						report.setOperationType("");
						report.setTransactioncode(transactionCode);
						webDriver.DataInput("", controllerTestCaseID, transactionType, transactionCode, "");
						rowFound = true;
						break;
					}
					if (!transactionType.startsWith("Verify")) {
						operationType = "Input";
					}

					String structureFolder = null;

					if (silo.trim().equalsIgnoreCase("PAS") || silo.trim().equalsIgnoreCase("Claims")
							|| silo.trim().equalsIgnoreCase("Digital1st") || silo.trim().equalsIgnoreCase("CS")
							|| silo.trim().equalsIgnoreCase("DM")) {

						structureFolder = Config.inputDataFilePath;
						if (transactionType.startsWith("Verify") && (!directoryPath.isEmpty())
								&& (!inputExcel.isEmpty())) {
							operationType = "InputandVerify";
							log.info("--------------------------InputandVerify--------------------------");
						} else if (transactionType.startsWith("Verify") && (directoryPath.isEmpty())
								&& (inputExcel.isEmpty())) {
							operationType = "Verify";
						}
					} else if (silo.trim().equalsIgnoreCase("Billing")) {
						structureFolder = Config.structureSheetFilePath;
						if (transactionType.startsWith("Verify") && (!StringUtils.isEmpty(directoryPath))
								&& (!StringUtils.isEmpty(inputExcel)) && (StringUtils.isEmpty(transactionCode))) {
							operationType = "InputandVerify";

						}
						if (transactionType.startsWith("Verify") && (transactionCode.startsWith("Verify"))
								&& (!StringUtils.isEmpty(directoryPath)) && (!StringUtils.isEmpty(inputExcel))) {
							operationType = "InputandCapture";

						} else if (transactionType.startsWith("Verify") && (StringUtils.isEmpty(directoryPath))
								&& (StringUtils.isEmpty(inputExcel))) {
							operationType = "Verify";
						}
					}
					log.info("OperationType is:" + operationType);

					if (controllerTransactionType.equalsIgnoreCase(transactionType)) {
						if ((directoryPath == null || inputExcel == null) && operationType != "Verify") {
							controller.pauseFun("Please Enter the directory or excelsheet name");
						} else {
							String inputFilePath = null;
							String structurePath = null;
							if (operationType != "Verify") {
								inputFilePath = Config.inputDataFilePath + directoryPath + "\\" + inputExcel;
								directoryPathFileUpload = directoryPath;
								structurePath = structureFolder + directoryPath + "\\" + inputExcel;
							}
							report.setInputPath(inputFilePath);
							report.setOperationType(operationType);
							report.setTransactioncode(transactionCode);
							webDriver.getReport().setTransactioncode(transactionCode);
							webDriver.getReport().setTestcaseId(controllerTestCaseID);

							webDriver.DataInput(structurePath, inputFilePath, controllerTestCaseID, transactionType,
									transactionCode, operationType, cycleDate);

							rowFound = true;
							break;
						}
					}

				} else if (!transactionType.equalsIgnoreCase(controllerTransactionType) && rowIndex == rowCount - 1) {
					controller.pauseFun("Transaction " + controller.controllerTransactionType + " Not Found in "
							+ Config.transactionInputFilePath);
				}
			}
			if (rowFound == true)
				break;
		}
		return null;
	}
}
