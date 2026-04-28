package com.majesco.itaf.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Date;
import java.util.HashMap;
import javax.swing.JOptionPane;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import com.majesco.itaf.batch.RunBatch;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.util.InvoiceProcessing;
import com.majesco.itaf.util.JDBCConnection;
import com.majesco.itaf.util.ReconReport;
import com.majesco.itaf.vo.Reporter;
import com.majesco.itaf.webservice.WebService;
import com.majesco.itaf.util.CalendarSnippet;

public class MainControllerBilling extends MainController {

	private final static Logger log = LogManager.getLogger(MainControllerBilling.class);

	private static boolean recoverydone = false;
	private static int failedscenarionum = 0;
	private static Sheet MainControlSheet = null;
	private static Cell sc_no = null;
	public static String mainControllerHeaderNo = null;
	private static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	public static Date frmDateNew = null;

	@SuppressWarnings({ "resource", "finally" })
	public Reporter ControllerData(String FilePath) throws NullPointerException, Exception {
		// log.info("In MainController value of pauseExecution1:"+pauseExecution);
		Reporter report = new Reporter();
		Sheet reqSheet = null;
		int execFlag = 0, scenarioNO = 0, rowCount = 0;

		HashMap<String, Integer> startpointerValues = new HashMap<String, Integer>();
		int startCol = 0;
		int startRow = 0;
		Row controllerRow = null;
		HashMap<String, Integer> sheetValues = new HashMap<String, Integer>();
		Row startpointerRow = null;
		int START_ROW = 0;
		int START_COLUMN = 0;
		Row CycleDateRow = null;
		Cell cycleDate = null;
		String executionApproach = "";
		String cbd_requestxml = null;
		String cbd_wsdl = null;
		Sheet startpointerSheet = null;
		Sheet MainControllerSheet = null;
		Row updatedstartpointerRow = null;
		HashMap<String, Integer> updatedstartpointerValues = new HashMap<String, Integer>();
		int UPDATED_START_ROW = 0;
		int UPDATED_START_COLUMN = 0;
		int ustartCol = 0;
		int ustartRow = 0;
		boolean columnChanged = false;

		// boolean rowChanged = false;//Sel4
		boolean linearpExecution = false;// Tanvi :4/11/2017

		try {
			MainControllerBilling.frmDateNew = new Date();
			reqSheet = WebHelperUtil.getSheet(Config.controllerFilePath, "MainControlSheet");
			sheetValues = WebHelperUtil.getValueFromHashMap(reqSheet);
			execFlag = sheetValues.get("ExecuteFlag");
			scenarioNO = sheetValues.get("SC_NO");
			rowCount = reqSheet.getLastRowNum() + 1;
		} catch (NullPointerException ne) {
			log.error(ne.getMessage(), ne);
			log.error("Failed Read MainControlSheet in MainCOntroller  <-|-> LocalizeMessage "
					+ ne.getLocalizedMessage() + " <-|-> Message " + ne.getMessage() + " <-|-> Cause " + ne.getCause());
			throw new Exception("Failed Read MainControlSheet in MainCOntroller <-|-> LocalizeMessage "
					+ ne.getLocalizedMessage() + " <-|-> Message" + ne.getMessage() + " <-|-> Cause " + ne.getCause());
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			log.error("Failed get MainControlSheet in MainCOntroller <-|-> LocalizeMessage " + e.getLocalizedMessage()
					+ " <-|-> Message " + e.getMessage() + " <-|-> Cause " + e.getCause());
			throw new Exception("Failed get MainControlSheet in MainCOntroller  <-|-> LocalizeMessage "
					+ e.getLocalizedMessage() + " <-|-> Message" + e.getMessage() + " <-|-> Cause " + e.getCause());
		}
		// log.info("row count is : "+rowCount);
		int colCount = 0;
		// boolean isStartFound = false;
		executionApproach = Config.executionApproach;
		try {
			startpointerSheet = WebHelperUtil.getSheet(Config.controllerFilePath, "StartPointer");

			startpointerValues = WebHelperUtil.getValueFromHashMap(startpointerSheet);
			startpointerRow = startpointerSheet.getRow(1);
			START_ROW = startpointerValues.get("StartRow");
			START_COLUMN = startpointerValues.get("StartCol");
			startRow = (int) startpointerRow.getCell(START_ROW).getNumericCellValue();
			startCol = (int) startpointerRow.getCell(START_COLUMN).getNumericCellValue();
		} catch (NullPointerException ne) {
			log.error(ne.getMessage(), ne);
			log.error("Failed get Start Pointer  <-|-> LocalizeMessage " + ne.getLocalizedMessage() + " <-|-> Message "
					+ ne.getMessage() + " <-|-> Cause " + ne.getCause());
			throw new Exception("Failed get Start Pointer <-|-> LocalizeMessage " + ne.getLocalizedMessage()
					+ " <-|-> Message" + ne.getMessage() + " <-|-> Cause " + ne.getCause());
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			log.error("Failed create Chrome DriverInstance <-|-> LocalizeMessage " + e.getLocalizedMessage()
					+ " <-|-> Message " + e.getMessage() + " <-|-> Cause " + e.getCause());
			throw new Exception("Failed to read process list  <-|-> LocalizeMessage " + e.getLocalizedMessage()
					+ " <-|-> Message" + e.getMessage() + " <-|-> Cause " + e.getCause());
		}
		startRow = startRow - 1;
		startCol = startCol - 1;

		if (executionApproach.equalsIgnoreCase("Linear")) // Linear
		{
			// for(int rowIndex=startRow;rowIndex<rowCount;rowIndex++)

			ustartCol = startCol;
			ustartRow = startRow;
			for (int rowIndex = startRow; rowIndex < rowCount; rowIndex++)

			{
				if (pExecution == true) {
					return report;
				}

				linearpExecution = false;
				ustartRow = rowIndex + 1;
				pauseExecution = false;

				reqSheet = WebHelperUtil.getSheet(Config.controllerFilePath, "MainControlSheet");
				sheetValues = WebHelperUtil.getValueFromHashMap(reqSheet);
				controllerRow = reqSheet.getRow(rowIndex);
				colCount = controllerRow.getLastCellNum();

				/*
				 * if (rowIndex > startRow) // : 4/11/2017 :START { rowChanged = true; startCol
				 * = 10; }
				 */

				testDescription = WebHelperUtil.getCellData("Test_Description", reqSheet, rowIndex, sheetValues);
				Cell executeFlag = controllerRow.getCell(execFlag);
				controllerTestCaseID = controllerRow.getCell(sheetValues.get("TestCaseID"));
				controllerGroupName = controllerRow.getCell(sheetValues.get("GroupName"));
				sc_no = controllerRow.getCell(sheetValues.get("SC_NO"));

				if (controllerTestCaseID.getStringCellValue().equalsIgnoreCase("")
						|| controllerTestCaseID.equals(null)) {
					log.info("No KeyWord Found");
					continue;
				}
				if (executeFlag != null) {
					if (executeFlag.toString().equalsIgnoreCase("Y")) {
						// for(int
						// columnIndex=startCol+1;columnIndex<colCount&&!pauseExecution;columnIndex++)
						System.out.println("pExecution" + pExecution);
						for (int columnIndex = startCol; columnIndex < colCount && !linearpExecution; columnIndex++) {

							controllerTransactionType = controllerRow.getCell(columnIndex);

							mainControllerHeaderNo = reqSheet.getRow(0).getCell(columnIndex).toString();
							System.out.println(mainControllerHeaderNo);
							CycleDateRow = reqSheet.getRow(1);
							cycleDate = CycleDateRow.getCell(columnIndex);
							cycleDateCellValue = cycleDate.getStringCellValue().toString();
							cycleDateValue = null;

							if (cycleDateValue == null || cycleDateValue.toString().equalsIgnoreCase("")) {
								cycleDateValue = "NA";
							}
							if (cycleDateCellValue == null || cycleDateCellValue.toString().equalsIgnoreCase("")) {
								cycleDateCellValue = "NA";
							}

							if ((controllerTransactionType != null
									&& StringUtils.isNotBlank(controllerTransactionType.getStringCellValue()))
									&& (!cycleDateCellValue.equalsIgnoreCase("NA"))) {

								try {
									ResultSet rs = null;
									Connection conn = JDBCConnection.establishHTML5BillingDBConn();
									Statement st = conn.createStatement();
									
								//06/08/2025 - USIC Project
									if ((Config.databaseAdminaccess != null) && (Config.databaseAdminaccess.equalsIgnoreCase("FALSE"))) {
										rs = st.executeQuery("select * from BILLING_CONTENT.business_day");
									}
									else {
									rs = st.executeQuery("select * from business_day");
									}
									while (rs.next()) {
										businessDate = rs.getString("business_date");
									}
									rs.close();
									st.close();
									JDBCConnection.closeConnection(conn);
								} catch (Exception e) {
									log.error(e.getMessage(), e);
									log.error("Failed to get Current Date from Database<-|-> LocalizeMessage "
											+ e.getLocalizedMessage() + " <-|-> Message " + e.getMessage()
											+ " <-|-> Cause " + e.getCause());
								}

								// log.info("businessDate in application database is :"+businessDate);
								if (businessDate != null) {
									try {
										SimpleDateFormat simpleDateFormat = new SimpleDateFormat(
												"yyyy-MM-dd HH:mm:ss.SS");
										Date newBDate = simpleDateFormat.parse(businessDate);
										DateFormat cycleDateFormat = new SimpleDateFormat("MM/dd/yyyy");
										cycleDateValue = cycleDateFormat.format(newBDate);
										log.info("Inside business date change method : Maincontroller");

										webDriver.getReport().setTrasactionType("ChangeBusinessdate");
										String fromDate = Config.dtFormat.format(new Date());
										webDriver.getReport().setFromDate(fromDate);
										webDriver.getReport().setStatus("");
										webDriver.getReport().setMessage("");
										log.info("Maincontroller cycle date is not equal to application business date");
										businessDateValue = cycleDateCellValue.toString();
										log.info("Application business date should be : " + businessDateValue);
										// TransactionMapping.TransactionInputData("ChangeBusinessDate");

										cbd_requestxml = Config.changebusinessdaterequestxmlpath;
										cbd_wsdl = Config.changebusinessdatewsdl;

										DocumentBuilderFactory cbd_request_dbFactory = DocumentBuilderFactory
												.newInstance();
										DocumentBuilder cbd_request_dBuilder = cbd_request_dbFactory
												.newDocumentBuilder();
										Document cbd_request_doc = cbd_request_dBuilder.parse(cbd_requestxml);
										cbd_request_doc.getDocumentElement().normalize();
										SimpleDateFormat cycledateformat = new SimpleDateFormat("MM/dd/yyyy");
										Date cycledate = cycledateformat.parse(businessDateValue);
										DateFormat xmlformat = new SimpleDateFormat("yyyy-MM-dd");
										xmlbusinessdate = xmlformat.format(cycledate);
										NodeList request_list = cbd_request_doc.getElementsByTagName("*");

										for (int i = 0; i < request_list.getLength(); i++) {
											Node node = request_list.item(i);
											if (node.getNodeName().equalsIgnoreCase("Businessdate")) {
												node.setTextContent(xmlbusinessdate);
												break;
											}
										}
										TransformerFactory cbd_transformerFactory = TransformerFactory.newInstance();
										Transformer cbd_transformer = cbd_transformerFactory.newTransformer();
										DOMSource cbd_source = new DOMSource(cbd_request_doc);
										StreamResult cbd_result = new StreamResult(new File(cbd_requestxml));
										cbd_transformer.transform(cbd_source, cbd_result);

										String cbd_responsexml = WebService.callWebService("", cbd_wsdl, cbd_requestxml,
												Config.user, Config.password);

										DocumentBuilderFactory cbd_reponse_dbFactory = DocumentBuilderFactory
												.newInstance();
										DocumentBuilder cbd_response_dBuilder = cbd_reponse_dbFactory
												.newDocumentBuilder();
										Document cbd_response_doc = cbd_response_dBuilder.parse(cbd_responsexml);
										cbd_response_doc.getDocumentElement().normalize();

										NodeList response_list = cbd_response_doc.getElementsByTagName("*");
										String response_nodevalue1 = null;
										String response_nodevalue2 = null;
										for (int j = 0; j < response_list.getLength(); j++) {
											Node response_node = response_list.item(j);
											if (response_node.getNodeName().equalsIgnoreCase("ProcessStatusFlag")) {
												response_nodevalue1 = response_node.getTextContent();
											}
											if (response_node.getNodeName().equalsIgnoreCase("SuccessFlag")) {
												response_nodevalue2 = response_node.getTextContent();
												break;
											}
										}

										if (response_nodevalue1.equalsIgnoreCase("COMPLETED")
												&& response_nodevalue2.equalsIgnoreCase("SUCCESS")
												&& cbd_responsexml.contains(xmlbusinessdate)) {
											webDriver.getReport().setStatus("PASS");
										} else {
											webDriver.getReport().setStatus("FAIL");
										}

										webDriver.getReport().setMessage(response_nodevalue1 + " ; "
												+ response_nodevalue2 + " ; " + xmlbusinessdate);
									} catch (Exception e) {
										log.error(e.getMessage(), e);
										webDriver.getReport().setMessage(e.getLocalizedMessage());
										webDriver.getReport().setStatus("FAIL");
										pauseFun("Webservice issue " + e);

									} finally {
										String toDate = Config.dtFormat.format(new Date());
										webDriver.getReport().setToDate(toDate);
										ExcelUtility.writeReport(webDriver.getReport());
										log.info("Application business date is changed and rechecking in database");
										webDriver.getReport().setTrasactionType("");
									}

								}
							}

							log.info("Value of controllerTestCaseID : " + controllerTestCaseID);
							log.info("Value of controllerTransactionType : " + controllerTransactionType);
							log.info("Value of cycleDateCellValue : " + cycleDateCellValue);
							// Tanvi: 4/25/2017
							log.info("Value of cycleDateValue : " + cycleDateValue);

							// TM: wrapped the PAUSE if into another if as
							// replacement of above commented if
							if (controllerTransactionType != null
									&& StringUtils.isNotBlank(controllerTransactionType.getStringCellValue())
									&& (cycleDateCellValue.equalsIgnoreCase(cycleDateValue))) {
								if (controllerTransactionType.toString().equalsIgnoreCase("PAUSE")) {
									pauseFun("Do You Wish To Continue");
								}

								else if (controllerTransactionType.toString().equalsIgnoreCase("END")
										|| controllerTransactionType.toString().equalsIgnoreCase("END1")) {
									try {
										FileInputStream in = new FileInputStream(Config.controllerFilePath);
										Workbook LinearupdatedmainWB = new HSSFWorkbook(in);
										MainControllerSheet = LinearupdatedmainWB.getSheet("StartPointer");
										updatedstartpointerRow = MainControllerSheet.getRow(1);
										updatedstartpointerValues = WebHelperUtil
												.getValueFromHashMap(MainControllerSheet);
										UPDATED_START_ROW = updatedstartpointerValues.get("StartRow");
										UPDATED_START_COLUMN = updatedstartpointerValues.get("StartCol");
										System.out.println("ustartRow" + ustartRow);
										System.out.println("ustartCol" + ustartCol);
										updatedstartpointerRow.getCell(UPDATED_START_ROW).setCellValue(ustartRow);
										ustartCol = ustartCol + 2;
										updatedstartpointerRow.getCell(UPDATED_START_COLUMN).setCellValue(ustartCol);
										FileOutputStream out = new FileOutputStream(Config.controllerFilePath);
										System.out.println("ustartRow" + ustartRow);
										System.out.println("ustartCol" + ustartCol);
										LinearupdatedmainWB.write(out);
										out.flush();
										out.close();
										in.close();
									} catch (Exception e) {
										log.error(e.getMessage(), e);
										e.printStackTrace();
									} finally {
										webDriver.getReport().setTrasactionType(controllerTransactionType.toString());
										webDriver.getReport().setTestDescription(testDescription.toString());
										webDriver.getReport().setTestcaseId(controllerTestCaseID.toString());
										webDriver.getReport().setCycleDate(cycleDateCellValue.toString());
										return report;
									}
								} else {

									// Tanvi : 4/11/2017 :START
									ustartCol = columnIndex + 1;
									FileInputStream in = new FileInputStream(Config.controllerFilePath);
									Workbook LinearupdatedmainWB = new HSSFWorkbook(in);
									MainControllerSheet = LinearupdatedmainWB.getSheet("StartPointer");
									updatedstartpointerRow = MainControllerSheet.getRow(1);
									updatedstartpointerValues = WebHelperUtil.getValueFromHashMap(MainControllerSheet);
									UPDATED_START_ROW = updatedstartpointerValues.get("StartRow");
									UPDATED_START_COLUMN = updatedstartpointerValues.get("StartCol");
									System.out.println("ustartRow" + ustartRow);
									System.out.println("ustartCol" + ustartCol);
									updatedstartpointerRow.getCell(UPDATED_START_ROW).setCellValue(ustartRow);
									updatedstartpointerRow.getCell(UPDATED_START_COLUMN).setCellValue(ustartCol);
									FileOutputStream out = new FileOutputStream(Config.controllerFilePath);
									LinearupdatedmainWB.write(out);
									out.flush();
									out.close();
									in.close();

									webDriver.getReport().setTrasactionType(controllerTransactionType.toString());
									webDriver.getReport().setTestDescription(testDescription.toString());
									webDriver.getReport().setTestcaseId(controllerTestCaseID.toString());
									webDriver.getReport().setCycleDate(cycleDateCellValue.toString());
									if (controllerTransactionType.toString().equalsIgnoreCase("RunBatch_JBEAM")) {
										if (!cycleDateCellValue.equalsIgnoreCase("NA")) {
											report = RunBatch.RunBatch_WS();
											Date endDate = new Date();
											webDriver.getReport().setToDate(Config.dtFormat.format(endDate));
											ExcelUtility.writeReport(webDriver.getReport());
										} else {
											Date endDate = new Date();
											webDriver.getReport().setToDate(Config.dtFormat.format(endDate));
											webDriver.getReport().setStatus("Cycledate not available.");
											ExcelUtility.writeReport(webDriver.getReport());
										}

									}

									else if (controllerTransactionType.toString().equalsIgnoreCase("SC_END")
											&& !pExecution) {
										String recovery_scenario_Temp = Config.recovery_scenario;
										if (recovery_scenario_Temp.equalsIgnoreCase("FALSE")) {
											Config.recovery_scenario = "TRUE";
										}

										recoveryhandler();

										Config.recovery_scenario = recovery_scenario_Temp;
									}

									else {
										report = TransactionMapping.TransactionInputData(cycleDateCellValue.toString(),
												controllerTestCaseID.toString(), controllerTransactionType.toString(),
												Config.transactionInputFilePath);
									}
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
		} else if (executionApproach.equalsIgnoreCase("Timetravel")) // Timetravel
		{

			int columnIndex2 = startCol;
			for (int rowIndex2 = 1; rowIndex2 < rowCount; columnIndex2++) {

				if (pExecution == true) {
					return report;
				}

				ustartCol = columnIndex2 + 1;
				controllerRow = reqSheet.getRow(rowIndex2);
				CycleDateRow = reqSheet.getRow(1);
				colCount = CycleDateRow.getLastCellNum();
				if (columnIndex2 >= colCount)
					break;

				cycleDate = controllerRow.getCell(columnIndex2);
				colCount = controllerRow.getLastCellNum();
				cycleDateCellValue = cycleDate.getStringCellValue();
				try {
					ResultSet rs = null;
					Connection conn = JDBCConnection.establishHTML5BillingDBConn();
					Statement st = conn.createStatement();

					//06/08/2025 - USIC Project
					if ((Config.databaseAdminaccess != null) && (Config.databaseAdminaccess.equalsIgnoreCase("FALSE"))) {
						rs = st.executeQuery("select * from BILLING_CONTENT.business_day");
					}
					else {
					rs = st.executeQuery("select * from business_day");
					}
					while (rs.next()) {
						businessDate = rs.getString("business_date");
					}
					rs.close();
					st.close();
					JDBCConnection.closeConnection(conn);
				} catch (Exception e) {
					log.error(e.getMessage(), e);
					log.error("Failed to get Current Date from Database<-|-> LocalizeMessage " + e.getLocalizedMessage()
							+ " <-|-> Message " + e.getMessage() + " <-|-> Cause " + e.getCause());
					throw new Exception(
							"Failed to get Current Date from Database<-|-> LocalizeMessage " + e.getLocalizedMessage()
									+ " <-|-> Message" + e.getMessage() + " <-|-> Cause " + e.getCause());
				}
				SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
				Date newBDate = simpleDateFormat.parse(businessDate);
				DateFormat cycleDateFormat = new SimpleDateFormat("MM/dd/yyyy");
				cycleDateValue = cycleDateFormat.format(newBDate);
				// log.info("newBDate is :"+newBDate);

				log.info(
						"Cycle Date of MainController sheet is :" + cycleDateCellValue + " for column " + columnIndex2);
				log.info("Business Date of Application database is :" + cycleDateValue + " for column " + columnIndex2);

				if (columnIndex2 > startCol) {
					columnChanged = true;
				}

				// ***Below given code is added to avoid businessday change
				// during debugging***// 05/21/2019

				if (Config.ChangeBusinessDay != null) {
					if (Config.ChangeBusinessDay.equalsIgnoreCase("False")) {
						cycleDateValue = cycleDateCellValue;
						log.info("Debugging without changing the businessday of the application");
					}
				}

				// ***End***//

				// If Cycle Date in MainController Sheet is equal to Business
				// Date of the Application
				if (cycleDateCellValue.equalsIgnoreCase(cycleDateValue)) {
					log.info("Maincontroller cycle date is equal to application business date");
					if (columnChanged == true) {
						startRow = 1;
					}

					Sheet startPointerSheet = null;
					Sheet mainControlData = WebHelperUtil.getSheet(Config.controllerFilePath, "MainControlSheet");

					for (int rowIndex = startRow; rowIndex < rowCount; rowIndex++) {

						if (!pExecution) {
							ustartRow = rowIndex + 1;
							// log.info("start pointer Updated to STARTROW :" + ustartRow + " STARTCOLUMN :
							// " + ustartCol);
							// log.info("***************************PLEASE DONT OPEN MAINCONTROLLER
							// SHEET*********************");
							try {
								/*
								 * FileInputStream in = new FileInputStream(Config.controllerFilePath); Workbook
								 * updatedmainWB = new XSSFWorkbook(in); startPointerSheet =
								 * updatedmainWB.getSheet("StartPointer"); updatedstartpointerRow =
								 * startPointerSheet.getRow(1); updatedstartpointerValues =
								 * WebHelperUtil.getValueFromHashMap(startPointerSheet); UPDATED_START_ROW =
								 * updatedstartpointerValues.get("StartRow"); UPDATED_START_COLUMN =
								 * updatedstartpointerValues.get("StartCol");
								 * updatedstartpointerRow.getCell(UPDATED_START_ROW).setCellValue(ustartRow);
								 * updatedstartpointerRow.getCell(UPDATED_START_COLUMN).setCellValue(ustartCol);
								 * 
								 * // ****08/03/2019 code added below to minimize // MainController.xls file
								 * access Sheet mainControlData = updatedmainWB.getSheet("MainControlSheet");
								 */

								FileInputStream in = null;
								Workbook updatedmainWB = null;

								try {
									Cell tempCell = mainControlData.getRow(ustartRow - 1).getCell(ustartCol - 1);

									if (tempCell == null || tempCell.getStringCellValue().equals(""))
									// if(StringUtils.isEmpty(tempCell.getStringCellValue()))
									{
										// in.close();
										// updatedmainWB.close();
										continue;
									}
									// log.info("Transaction is: " + tempCell.getStringCellValue());
									// Cell exeFlagCell = mainControlData.getRow(ustartRow - 1).getCell(3);

									in = new FileInputStream(Config.controllerFilePath);
									updatedmainWB = new XSSFWorkbook(in);
									Sheet mainControlDataCur = updatedmainWB.getSheet("MainControlSheet");
									Cell exeFlagCell = mainControlDataCur.getRow(ustartRow - 1).getCell(3);

									if (exeFlagCell.getStringCellValue().equalsIgnoreCase("N")) {
										// log.info(tempCell.getStringCellValue() + " transaction's exe. flag is N");
										log.info("Execution flag of transaction '" + tempCell.getStringCellValue()
												+ "' at Row : " + ustartRow + ", Column : " + ustartCol + " is N");
										in.close();
										updatedmainWB.close();
										continue;
									}

									startPointerSheet = updatedmainWB.getSheet("StartPointer");
									updatedstartpointerRow = startPointerSheet.getRow(1);
									updatedstartpointerValues = WebHelperUtil.getValueFromHashMap(startPointerSheet);
									UPDATED_START_ROW = updatedstartpointerValues.get("StartRow");
									UPDATED_START_COLUMN = updatedstartpointerValues.get("StartCol");
									updatedstartpointerRow.getCell(UPDATED_START_ROW).setCellValue(ustartRow);
									updatedstartpointerRow.getCell(UPDATED_START_COLUMN).setCellValue(ustartCol);
									log.info("StartPointer Updated to STARTROW :" + ustartRow + " STARTCOLUMN : "
											+ ustartCol);
									log.info("Transaction is: " + tempCell.getStringCellValue());

								} catch (Exception e) {
									log.info("Exception occured while reading MainContolData: " + e.getMessage());
								}
								// ***********
								if (in != null) {
									in.close();
									FileOutputStream out = new FileOutputStream(Config.controllerFilePath);
									updatedmainWB.write(out);
									out.flush();
									out.close();
									updatedmainWB.close();
								}
							} catch (Exception e) {
								log.error(e.getMessage(), e);
								if (e.getMessage().toString().contains(
										"The process cannot access the file because it is being used by another process")) {
									rowIndex--;
									log.info("Handle start pointer exception Updated to STARTROW :" + ustartRow
											+ " STARTCOLUMN : " + ustartCol);
									CalendarSnippet.killProcess("EXCEL.EXE");
									Thread.sleep(500);
									continue;
								}
								log.error(
										"Failed to update StartPointer <-|-> LocalizeMessage " + e.getLocalizedMessage()
												+ " <-|-> Message " + e.getMessage() + " <-|-> Cause " + e.getCause());
								throw new Exception(
										"Failed to update StartPointer <-|-> LocalizeMessage " + e.getLocalizedMessage()
												+ " <-|-> Message" + e.getMessage() + " <-|-> Cause " + e.getCause());
							}
						}

						pauseExecution = false;

						if (recoverydone == true) {
							controllerRow = MainControlSheet.getRow(rowIndex);
							// recoverydone = false;
						} else {
							controllerRow = reqSheet.getRow(rowIndex);
						}

						CycleDateRow = reqSheet.getRow(1);
						colCount = CycleDateRow.getLastCellNum();
						testDescription = WebHelperUtil.getCellData("Test_Description", reqSheet, rowIndex,
								sheetValues);

						Cell executeFlag = controllerRow.getCell(execFlag);
						sc_no = controllerRow.getCell(scenarioNO);
						controllerTestCaseID = controllerRow.getCell(sheetValues.get("TestCaseID"));
						controllerGroupName = controllerRow.getCell(sheetValues.get("GroupName"));
						if (controllerTestCaseID == null
								|| controllerTestCaseID.getStringCellValue().equalsIgnoreCase("")) {
							// log.info("No KeyWord Found");
							continue;
						}
						if (executeFlag != null) {
							if (executeFlag.toString().equalsIgnoreCase("Y")) {

								if (columnIndex2 <= colCount && !pauseExecution && !pExecution)
								// bhaskar Pause Functionality END
								{
									controllerTransactionType = controllerRow.getCell(columnIndex2);
									log.info("Value of Main controller TestCaseID : " + controllerTestCaseID
											+ " TransactionType : " + controllerTransactionType + " Execute Flag is : "
											+ executeFlag.toString());
									// log.info("Value of Main controller sc_no : "+sc_no);
									if (controllerTransactionType != null
											&& StringUtils.isNotBlank(controllerTransactionType.getStringCellValue())) {

										webDriver.getReport().setTrasactionType(controllerTransactionType.toString());

										webDriver.getReport().setMessage("");
										webDriver.getReport().setScreenShot("");
										webDriver.getReport().setTestcaseId(controllerTestCaseID.toString());
										MainControllerBilling.frmDateNew = new Date();

										if (controllerTransactionType.toString().equalsIgnoreCase("PAUSE")) {
											// pauseFun("Do You Wish To Continue");
											String msg = "Do You Wish To Continue";
											String uInteraction = Config.userInteraction;
											if (!uInteraction.equalsIgnoreCase("FALSE")) {
												webDriver.getFrame().setVisible(true);
												webDriver.getFrame().setAlwaysOnTop(true);
												webDriver.getFrame().setLocationRelativeTo(null);

												JOptionPane.setRootFrame(webDriver.getFrame());
												int stopexecution = JOptionPane.showConfirmDialog(webDriver.getFrame(),
														msg, msg, JOptionPane.YES_NO_OPTION);
												// log.info(stopexecution);
												if (stopexecution == JOptionPane.NO_OPTION) {
													// pExecution = true;
												}
												webDriver.getFrame().dispose();
												// log.info(stopexecution);
											}

										}

										// Added keyword END for SUITE
										// Integration//
										if (controllerTransactionType.toString().equalsIgnoreCase("END")) {
											try {
												FileInputStream in = new FileInputStream(Config.controllerFilePath);

												Workbook updatedmainWB = new XSSFWorkbook(in);
												startPointerSheet = updatedmainWB.getSheet("StartPointer");

												updatedstartpointerRow = startPointerSheet.getRow(1);
												updatedstartpointerValues = WebHelperUtil
														.getValueFromHashMap(startPointerSheet);
												UPDATED_START_ROW = updatedstartpointerValues.get("StartRow");
												UPDATED_START_COLUMN = updatedstartpointerValues.get("StartCol");
												updatedstartpointerRow.getCell(UPDATED_START_ROW)
														.setCellValue((ustartRow + 1));
												updatedstartpointerRow.getCell(UPDATED_START_COLUMN)
														.setCellValue(ustartCol);
												FileOutputStream out = new FileOutputStream(Config.controllerFilePath);
												updatedmainWB.write(out);
												out.flush();
												out.close();
												updatedmainWB.close();
												in.close();
											} catch (Exception e) {
												log.error(e.getMessage(), e);
												e.printStackTrace();
											} finally {
												Date finalfrmDate = new Date();
												webDriver.getReport().setIteration(Config.cycleNumber);
												webDriver.getReport().setTestcaseId("Common");
												webDriver.getReport().setCycleDate(cycleDateValue);
												webDriver.getReport().setTrasactionType("END");
												webDriver.getReport().setFromDate(Config.dtFormat.format(finalfrmDate));
												return report;
												// System.exit(1);
											}
										}
										// END

										// Webservice transactions//

										else if (controllerTransactionType.toString().startsWith("WAITFOR")) {
											try {

												Date frmDate = new Date();
												String fromDate = Config.dtFormat.format(frmDate);
												webDriver.getReport().setFromDate(fromDate);
												webDriver.getReport().setIteration(Config.cycleNumber);
												webDriver.getReport().setTestcaseId(controllerTestCaseID.toString());
												webDriver.getReport().setCycleDate(cycleDateCellValue);
												webDriver.getReport()
														.setTrasactionType(controllerTransactionType.toString());

												String[] arrWait = controllerTransactionType.toString().split("_");
												int wSec;

												if (arrWait.length == 2) {
													wSec = Integer.parseInt(arrWait[1]);
													log.info("Wait in MainController : Waiting for " + wSec
															+ " seconds before next transaction");
													Thread.sleep((wSec) * 1000);

													Date toDate = new Date();
													String EndDate = Config.dtFormat.format(toDate);
													webDriver.getReport().setToDate(EndDate);
													webDriver.getReport().setStatus("PASS");
													ExcelUtility.writeReport(webDriver.report);

												} else {
													log.info(
															"Wait in MainController : As time is not given, waiting for 10 seconds before next transaction");
													Thread.sleep(10000);
												}
											} catch (Exception e) {
												log.error(e.getMessage(), e);
												log.info(e.getMessage());
											}

										}
										//

										// if(controllerTransactionType.toString().equalsIgnoreCase("RunBatch_JBEAM"))

										else if ((controllerTransactionType.toString()
												.equalsIgnoreCase("RunBatch_JBEAM")
												|| controllerTransactionType.toString()
														.equalsIgnoreCase("RunBatch_JBEAM_DAY")
												|| controllerTransactionType.toString()
														.equalsIgnoreCase("RunBatch_JBEAM_NIGHT"))
												&& !pExecution)
										// Batch
										// Pause Functionality END
										{
											WebHelper.testcaseID = controllerTestCaseID;
											WebHelper.transactionType = controllerTransactionType;
											report = RunBatch.RunBatch_WS();
											Date endDate = new Date();
											webDriver.getReport().setToDate(Config.dtFormat.format(endDate));
											ExcelUtility.writeReport(webDriver.getReport());

										}
										// For Copy Flat File
										else if (controllerTransactionType.toString()
												.equalsIgnoreCase("RunBatch_JBEAMFF") && !pExecution)

										{
											report = RunBatch.RunBatch_WSFF();
											Date endDate = new Date();
											webDriver.getReport().setToDate(Config.dtFormat.format(endDate));
											ExcelUtility.writeReport(webDriver.getReport());

										}

										else if (controllerTransactionType.toString().equalsIgnoreCase("Recon_Report")
												&& !pExecution) {
											boolean isReportGenerated = ReconReport.generateReconReport(
													cycleDateCellValue, Config.reconReportJarFilePath,
													Config.reconReportJarFileName, Config.reconReportGeneratedFilePath,
													Config.reconReportArchivalPath, Config.reconReportDateFormat,
													Config.reconReportArgs1, Config.reconReportArgs2);
											log.info("Recon_Report generated : " + isReportGenerated);

											if (!isReportGenerated) {
												recoveryhandler();
												log.info("Error in Recon Report, Execution is going to stop");
												webDriver.getReport().setStatus("FAIL");
												webDriver.getReport().setMessage(
														("Error in Recon Report, Execution is going to stop"));
												pExecution = true;
												webDriver.getFrame().dispose();
											}
										} else if (controllerTransactionType.toString()
												.equalsIgnoreCase("INVOICE_PROCESSING") && !pExecution) {
											try {
												InvoiceProcessing invoiceProcessing = new InvoiceProcessing(
														Config.invoiceProcessingRemotePath,
														Config.invoiceProcessingArchivePath,
														Config.invoiceProcessingLocalPath,
														Config.invoiceProcessingConfigExcelPath,
														Config.invoiceProcessingConfigExcelTab,
														Config.invoiceProcessingType, Config.controllerFilePath,
														cycleDateCellValue, Config.EnvironmentType,
														Config.flatFileUserName, Config.flatFilePassword,
														Config.flatFileHostName, Integer.parseInt(Config.flatFilePort),
														Config.invoiceProcessingReportExcelFile);

												invoiceProcessing.readConfigExcel(
														invoiceProcessing.CONFIG_EXCEL_FILE_PATH,
														invoiceProcessing.EXCEL_TAB_NAME);
												log.info(invoiceProcessing.getLabelMapping());

												boolean isCopyAndArchived;

												if (invoiceProcessing.IS_ENV_WINDOWS) {
													isCopyAndArchived = invoiceProcessing.copyAndArchivedFiles(
															invoiceProcessing.REMOTE_PATH, invoiceProcessing.DIR_PATH,
															invoiceProcessing.ARCHIVE_PATH);
												} else {
													isCopyAndArchived = invoiceProcessing.copyAndArchivedSFTPFiles(
															invoiceProcessing.REMOTE_PATH, invoiceProcessing.DIR_PATH,
															invoiceProcessing.ARCHIVE_PATH);
												}

												log.info("isCopyAndArchived : " + isCopyAndArchived);
												if (isCopyAndArchived) {
													boolean isFileExist = invoiceProcessing.isReportExcelFileExist(
															invoiceProcessing.REPORT_EXCEL_FILE);
													log.info("Report excel file exist : " + isFileExist);

													File[] files = invoiceProcessing
															.getFiles(invoiceProcessing.DIR_PATH);
													if (null != files && files.length > 0) {
														for (File file : files) {
															log.info("file.getName() : " + file.getName());
															invoiceProcessing.processFile(file,
																	invoiceProcessing.REPORT_EXCEL_FILE,
																	invoiceProcessing.DIR_PATH,
																	invoiceProcessing.FILE_EXTENSION,
																	invoiceProcessing.MAIN_CONTROLLER_PATH,
																	InvoiceProcessing.getDate(
																			invoiceProcessing.TEST_CASE_DATE,
																			InvoiceProcessing.DATE_FORMAT));
														}

														invoiceProcessing.movedProcessedFiles(
																invoiceProcessing.DIR_PATH,
																InvoiceProcessing.getDate(
																		invoiceProcessing.TEST_CASE_DATE,
																		InvoiceProcessing.DATE_FORMAT));
													} else {
														log.error("Files does not exist to process : "
																+ invoiceProcessing.DIR_PATH);
													}
												}
											} catch (Exception ex) {
												log.error("Error in invoice processing : " + ex.getMessage(), ex);
											}
										} else if (controllerTransactionType.toString().equalsIgnoreCase("Login2")
												&& !pExecution)

										{
											log.info("Transaction is Login2 - NewUser, hence performing login");
											Automation.setUp();
											TransactionMapping.TransactionInputData("Login2");
											Thread.sleep(3000);
											WebHelperUtil.saveScreenShot();
											columnChanged = false;
											webDriver.getReport()
													.setTrasactionType(controllerTransactionType.toString());

											// report =
											// TransactionMapping.TransactionInputData(cycleDateCellValue,controllerTestCaseID,controllerTransactionType,Config.transactionInputFilePath.toString());
										}

										else if (controllerTransactionType.toString().equalsIgnoreCase("SC_END")
												&& !pExecution) {
											recoveryhandler();
										} else if (controllerTransactionType.toString()
												.equalsIgnoreCase("VerifyCSVReport")) {
											verifyCSVReport();
											continue;
										}

										else if (controllerTransactionType.toString().trim()
												.equalsIgnoreCase("Task_Delete") && !pExecution) {
											try {
												Date frmDate = new Date();
												String fromDate = Config.dtFormat.format(frmDate);
												webDriver.getReport().setFromDate(fromDate);
												webDriver.getReport().setIteration(Config.cycleNumber);
												webDriver.getReport().setTestcaseId(controllerTestCaseID.toString());
												webDriver.getReport().setCycleDate(cycleDateCellValue);
												webDriver.getReport()
														.setTrasactionType(controllerTransactionType.toString());

												Connection conn = JDBCConnection.establishHTML5BillingDBConn();
												Statement st = conn.createStatement();

												st.execute("DELETE FROM [dbo].RS_WORKITEMPARTICIPANT");
												log.info("Deleted data from RS_WORKITEMPARTICIPANT");
												st.execute("DELETE FROM [dbo].RS_WORKITEM_EXT_ATTR");
												log.info("Deleted data from RS_WORKITEM_EXT_ATTR");
												st.execute("DELETE  FROM [dbo].RS_WORKITEMCACHE");
												log.info("Deleted data from RS_WORKITEMCACHE");
												Date toDate = new Date();
												String EndDate = Config.dtFormat.format(toDate);
												webDriver.getReport().setToDate(EndDate);
												webDriver.getReport().setStatus("PASS");
												ExcelUtility.writeReport(webDriver.report);

												st.close();
												JDBCConnection.closeConnection(conn);
											} catch (Exception e) {
												log.error(e.getLocalizedMessage(), e);
												webDriver.getReport().setStatus("FAIL");
												webDriver.getReport().setMessage(e.getLocalizedMessage());
												Date toDate = new Date();
												String EndDate = Config.dtFormat.format(toDate);
												webDriver.getReport().setToDate(EndDate);
												ExcelUtility.writeReport(webDriver.report);
											}
										} else if (controllerTransactionType.toString().trim()
												.equalsIgnoreCase("FF_Delete") && !pExecution) {
											try {
												Date frmDate = new Date();
												String fromDate = Config.dtFormat.format(frmDate);
												webDriver.getReport().setFromDate(fromDate);
												webDriver.getReport().setIteration(Config.cycleNumber);
												webDriver.getReport().setTestcaseId(controllerTestCaseID.toString());
												webDriver.getReport().setCycleDate(cycleDateCellValue);
												webDriver.getReport()
														.setTrasactionType(controllerTransactionType.toString());

												File InterfaceFile = new File(Config.InterfaceOutBoundPath);
												WebHelperUtil.deleteInterfaceFiles(InterfaceFile);

												webDriver.getReport()
														.setMessage("Interface files deleted successfully.");
												Date toDate = new Date();
												String EndDate = Config.dtFormat.format(toDate);
												webDriver.getReport().setToDate(EndDate);
												webDriver.getReport().setStatus("PASS");
												ExcelUtility.writeReport(webDriver.report);

											} catch (Exception e) {
												log.error(e.getLocalizedMessage(), e);
												webDriver.getReport().setStatus("FAIL");
												webDriver.getReport().setMessage(e.getLocalizedMessage());
												Date toDate = new Date();
												String EndDate = Config.dtFormat.format(toDate);
												webDriver.getReport().setToDate(EndDate);
												ExcelUtility.writeReport(webDriver.report);
											}
										} else if (controllerTransactionType.toString().trim()
												.equalsIgnoreCase("UpdateMainController")) {
											try {
												webDriver.getReport().setTrasactionType("UpdateMainController");
												webDriver.getReport().setCycleDate(cycleDateCellValue);
												String fromDate = Config.dtFormat.format(new Date());
												webDriver.getReport().setFromDate(fromDate);
												webDriver.getReport().setStatus("");
												webDriver.getReport().setMessage("");

												Sheet mainControllerSheet = WebHelperUtil
														.getSheet(Config.controllerFilePath, "MainControlSheet");
												HashMap<String, Integer> sheetHeaders = new HashMap<String, Integer>();
												sheetHeaders = WebHelperUtil.getValueFromHashMap(mainControllerSheet);
												int maincontrollersheetrowCount = mainControllerSheet.getLastRowNum()
														+ 1;
												FileInputStream in = new FileInputStream(Config.controllerFilePath);
												Workbook mainWB = null;
												if (Config.controllerFilePath.endsWith(".xls")) {
													mainWB = new HSSFWorkbook(in);
												} else if (Config.controllerFilePath.endsWith(".xlsx")
														|| Config.controllerFilePath.endsWith(".xlsm")) {
													mainWB = new XSSFWorkbook(in);
												}
												mainControllerSheet = mainWB.getSheet("MainControlSheet");
												List<String> dummyPolicyNum = new ArrayList<String>();
												int policyNumCol = sheetHeaders.get("PolicyNumber");
												int testCaseIDCol = sheetHeaders.get("TestCaseID");
												// identify dummyPolicyNum list
												HashMap<String, String> policyHashMap = new HashMap<String, String>();
												int ctr = 1;
												for (int i = 2; i < maincontrollersheetrowCount; i++) {
													Row currentRow = mainControllerSheet.getRow(i);
													String tempTestCaseID = "";
													try {
														tempTestCaseID = currentRow.getCell(testCaseIDCol).toString();
													} catch (Exception e) {
														tempTestCaseID = "";
													}
													if (tempTestCaseID
															.equalsIgnoreCase(controllerTestCaseID.toString())) {
														String curDummyPolicyNum = currentRow.getCell(policyNumCol)
																.toString();
														if (!dummyPolicyNum.contains(curDummyPolicyNum)) {
															dummyPolicyNum.add(curDummyPolicyNum);
															String policyColHeaderFromUnique = "PAS_PolicyNum" + ctr;
															policyHashMap.put(curDummyPolicyNum,
																	policyColHeaderFromUnique);
															ctr++;
														}
													}
												}
												// Update actual policy numbers in MainController from unique number
												// file
												for (int i = 2; i < maincontrollersheetrowCount; i++) {
													Row currentRow = mainControllerSheet.getRow(i);
													String tempTestCaseID = "";
													try {
														tempTestCaseID = currentRow.getCell(testCaseIDCol).toString();
													} catch (Exception e) {
														tempTestCaseID = "";
													}

													if (tempTestCaseID
															.equalsIgnoreCase(controllerTestCaseID.toString())) {
														String policyNumToUpdate = WebHelperUtil.ReadFromExcel("",
																policyHashMap.get(
																		currentRow.getCell(policyNumCol).toString()));
														currentRow.getCell(policyNumCol)
																.setCellValue(policyNumToUpdate);
													}
												}

												FileOutputStream out;
												out = new FileOutputStream(Config.controllerFilePath);
												mainWB.write(out);
												out.flush();
												out.close();
												in.close();
												webDriver.getReport().setStatus("PASS");
											} catch (Exception e) {
												webDriver.getReport().setStatus("FAIL");
											} finally {
												String toDate = Config.dtFormat.format(new Date());
												webDriver.getReport().setToDate(toDate);
												ExcelUtility.writeReport(webDriver.getReport());
												webDriver.getReport().setTrasactionType("");
											}
										} else if (!pExecution) {
											if ((columnChanged == true)
													&& (!(controllerTransactionType.toString().startsWith("WebService")
															| controllerTransactionType.toString()
																	.startsWith("RestServiceJSON")
															| controllerTransactionType.toString().contains("Outbound")
															| controllerTransactionType.toString()
																	.startsWith("OASServiceJSON")))) {
												log.info(
														"First UI transaction of the changed column occurred, hence performing login");

												Automation.setUp();
												TransactionMapping.TransactionInputData("Login");
												Thread.sleep(3000);
												WebHelperUtil.saveScreenShot();
												columnChanged = false;
												webDriver.getReport()
														.setTrasactionType(controllerTransactionType.toString());

											}
											webDriver.getReport().setCycleDate(cycleDateCellValue);
											report = TransactionMapping.TransactionInputData(
													cycleDateCellValue.toString(), controllerTestCaseID.toString(),
													controllerTransactionType.toString(),
													Config.transactionInputFilePath);
										}
									} else {
										// log.info("No Transaction Found in the Maincontroller at Cell : "
										// + columnIndex2);
									}

								}
							}
						} else {
							log.info("Execute Flag is not Set");
						}

					}
				}
				// }
				else // If Cycle Date in MainController Sheet is not equal to
						// Business Date of the Application
				{
					if (!pExecution) {
						try {
							webDriver.getReport().setTrasactionType("ChangeBusinessdate");
							webDriver.getReport().setCycleDate(cycleDateCellValue);
							String fromDate = Config.dtFormat.format(new Date());
							webDriver.getReport().setFromDate(fromDate);
							MainControllerBilling.frmDateNew = new Date();
							webDriver.getReport().setStatus("");
							webDriver.getReport().setMessage("");
							businessDateValue = cycleDateCellValue.toString();
							log.info("Maincontroller cycle date is : " + businessDateValue);
							log.info("Maincontroller cycle date is not equal to application business date");
							// TransactionMapping.TransactionInputData("ChangeBusinessDate");
							cbd_requestxml = Config.changebusinessdaterequestxmlpath;
							cbd_wsdl = Config.changebusinessdatewsdl;

							DocumentBuilderFactory cbd_request_dbFactory = DocumentBuilderFactory.newInstance();
							DocumentBuilder cbd_request_dBuilder = cbd_request_dbFactory.newDocumentBuilder();
							Document cbd_request_doc = cbd_request_dBuilder.parse(cbd_requestxml);
							cbd_request_doc.getDocumentElement().normalize();
							SimpleDateFormat cycledateformat = new SimpleDateFormat("MM/dd/yyyy");
							Date cycledate = cycledateformat.parse(businessDateValue);
							DateFormat xmlformat = new SimpleDateFormat("yyyy-MM-dd");
							xmlbusinessdate = xmlformat.format(cycledate);
							NodeList request_list = cbd_request_doc.getElementsByTagName("*");

							for (int i = 0; i < request_list.getLength(); i++) {
								Node node = request_list.item(i);
								if (node.getNodeName().equalsIgnoreCase("Businessdate")) {
									node.setTextContent(xmlbusinessdate);
									break;
								}
							}
							log.info("Date " + xmlbusinessdate + " updated in XML file : ");// Meghna
							TransformerFactory cbd_transformerFactory = TransformerFactory.newInstance();
							Transformer cbd_transformer = cbd_transformerFactory.newTransformer();
							DOMSource cbd_source = new DOMSource(cbd_request_doc);
							StreamResult cbd_result = new StreamResult(new File(cbd_requestxml));
							cbd_transformer.transform(cbd_source, cbd_result);

							String cbd_responsexml = WebService.callWebService("", cbd_wsdl, cbd_requestxml,
									Config.user, Config.password);

							DocumentBuilderFactory cbd_reponse_dbFactory = DocumentBuilderFactory.newInstance();
							DocumentBuilder cbd_response_dBuilder = cbd_reponse_dbFactory.newDocumentBuilder();
							Document cbd_response_doc = cbd_response_dBuilder.parse(cbd_responsexml);
							cbd_response_doc.getDocumentElement().normalize();

							NodeList response_list = cbd_response_doc.getElementsByTagName("*");
							String response_nodevalue1 = null;
							String response_nodevalue2 = null;
							for (int j = 0; j < response_list.getLength(); j++) {
								Node response_node = response_list.item(j);
								if (response_node.getNodeName().equalsIgnoreCase("ProcessStatusFlag")) {
									response_nodevalue1 = response_node.getTextContent();
								}
								if (response_node.getNodeName().equalsIgnoreCase("SuccessFlag")) {
									response_nodevalue2 = response_node.getTextContent();
									break;
								}
							}

							if (response_nodevalue1.equalsIgnoreCase("COMPLETED")
									&& response_nodevalue2.equalsIgnoreCase("SUCCESS")
									&& cbd_responsexml.contains(xmlbusinessdate)) {
								webDriver.getReport().setStatus("PASS");
							} else {
								webDriver.getReport().setStatus("FAIL");
							}

							webDriver.getReport().setMessage(
									response_nodevalue1 + " ; " + response_nodevalue2 + " ; " + xmlbusinessdate);
							log.info(response_nodevalue1 + " ; " + response_nodevalue2 + " ; " + xmlbusinessdate);
						} catch (Exception e) {
							log.error(e.getMessage(), e);
							webDriver.getReport().setMessage(e.getLocalizedMessage());
							webDriver.getReport().setStatus("FAIL");
						}

						// by devishree- on 16/03/2017//
						finally {
							String toDate = Config.dtFormat.format(new Date());
							webDriver.getReport().setToDate(toDate);
							ExcelUtility.writeReport(webDriver.getReport());
							columnIndex2 = columnIndex2 - 1;
							log.info("Application business date is changed and rechecking in database");
							webDriver.getReport().setTrasactionType("");
						}

					}

				}
			}
			// }
		}
		// log.info("In MainController value of pauseExecution3:"+pauseExecution);
		startCol = execFlag + 1;
		return report;
	}

	
	/*
	 * public void recoveryhandler() { HashMap<String, Integer> controlsheet = new
	 * HashMap<String, Integer>(); Row currentRow = null; int currentSC_NO = 0; int
	 * currentscenarionum = 0; Sheet reqSheet = null; Workbook mainWB = null;
	 * FileInputStream in = null; int maincontrollersheetrowCount = 0;
	 * 
	 * try {
	 * 
	 * // Recovery Scenario (Transaction) START if
	 * (Config.recovery_scenario.equalsIgnoreCase("TRUE")) { reqSheet =
	 * WebHelperUtil.getSheet(Config.controllerFilePath, "MainControlSheet");
	 * controlsheet = WebHelperUtil.getValueFromHashMap(reqSheet);
	 * maincontrollersheetrowCount = reqSheet.getLastRowNum() + 1; in = new
	 * FileInputStream(Config.controllerFilePath);
	 * 
	 * if (Config.controllerFilePath.endsWith(".xls")) { mainWB = new
	 * HSSFWorkbook(in);
	 * 
	 * } else if (Config.controllerFilePath.endsWith(".xlsx") |
	 * Config.controllerFilePath.endsWith(".xlsm")) { mainWB = new XSSFWorkbook(in);
	 * }
	 * 
	 * MainControlSheet = mainWB.getSheet("MainControlSheet"); failedscenarionum =
	 * (int) sc_no.getNumericCellValue(); log.info("Failed Scenario is  : " +
	 * "TestcaseID:" + controllerTestCaseID + " " + "SCNO:" + failedscenarionum);
	 * 
	 * for (int rowindx = 2; rowindx < maincontrollersheetrowCount; rowindx++) {
	 * currentRow = reqSheet.getRow(rowindx); currentscenarionum =
	 * controlsheet.get("SC_NO"); currentSC_NO = (int)
	 * currentRow.getCell(currentscenarionum).getNumericCellValue();
	 * 
	 * if (currentSC_NO == failedscenarionum) {
	 * log.info("Current MainController scenario number : " + currentSC_NO);
	 * HashMap<String, Integer> mainControlHashMap = WebHelperUtil
	 * .getValueFromHashMap(MainControlSheet); int executeflaghashmap =
	 * mainControlHashMap.get("ExecuteFlag"); Row ExecuteFlagRow =
	 * MainControlSheet.getRow(rowindx); Cell ExecuteFlagCell =
	 * ExecuteFlagRow.getCell(executeflaghashmap);
	 * 
	 * int ColAccountNumber = controlsheet.get("AccountNumber"); int ColPolicyNumber
	 * = controlsheet.get("PolicyNumber"); int ColAgentNumber =
	 * controlsheet.get("AgentNumber");
	 * 
	 * String STAccountNumber =
	 * currentRow.getCell(ColAccountNumber).getStringCellValue(); String
	 * STPolicyNumber = currentRow.getCell(ColPolicyNumber).getStringCellValue();
	 * String STAgentNumber =
	 * currentRow.getCell(ColAgentNumber).getStringCellValue();
	 * 
	 * //log.info("-------- The data pickup for SkipRowIdentifier ---------");
	 * //log.info("STAccountNumber = " + STAccountNumber.toString());
	 * //log.info("STPolicyNumber = " +
	 * currentRow.getCell(ColPolicyNumber).getStringCellValue());
	 * //log.info("STAgentNumber = " +
	 * currentRow.getCell(ColAgentNumber).getStringCellValue());
	 * 
	 * objFailedAccountList.add(STAccountNumber);
	 * objFailedBrokerList.add(STAgentNumber);
	 * objFailedPolicyList.add(STPolicyNumber);
	 * 
	 * if (!"GroupBilling".equalsIgnoreCase(Config.productTeam)) { if
	 * (controllerTransactionType.toString().equalsIgnoreCase("SC_END")) { //
	 * ExecuteFlagCell.setCellValue("Y");
	 * WebService.RunDBQueries("XX_Scenario_Completed", STPolicyNumber,
	 * STAgentNumber, STAccountNumber); } else { ExecuteFlagCell.setCellValue("N");
	 * log.info("Marking Execution Flag for SCNO:" + " " + currentSC_NO + " " +
	 * "as 'N'. This entire scenario will stop its further execution.");
	 * WebService.RunDBQueries("XX_Transaction_Failed", STPolicyNumber,
	 * STAgentNumber, STAccountNumber); } } else { if
	 * (controllerTransactionType.toString().equalsIgnoreCase("SC_END")) {
	 * ExecuteFlagCell.setCellValue("Y");
	 * WebService.RunDBQueries("XX_Scenario_Completed", STPolicyNumber,
	 * STAgentNumber, STAccountNumber); }
	 * 
	 * } } }
	 * 
	 * FileOutputStream out; out = new FileOutputStream(Config.controllerFilePath);
	 * mainWB.write(out); out.flush(); out.close(); in.close(); recoverydone = true;
	 * // log.info("End of Recovery"); } } catch (Exception e) { try {
	 * FileOutputStream out; out = new FileOutputStream(Config.controllerFilePath);
	 * mainWB.write(out); out.flush(); out.close(); } catch (Exception e1) {
	 * log.error("Error occured while saving file at location: " +
	 * Config.controllerFilePath); } log.error(e.getMessage(), e);
	 * log.error(e.getLocalizedMessage()); }
	 * 
	 * }
	 */
	
	
	public void recoveryhandler() {
	    HashMap<String, Integer> controlsheet = new HashMap<String, Integer>();
	    Row currentRow = null;
	    int currentSC_NO = 0;
	    int currentscenarionum = 0;
	    Sheet reqSheet = null;
	    Workbook mainWB = null;
	    FileInputStream in = null;
	    int maincontrollersheetrowCount = 0;

	    try {

	        // Recovery Scenario (Transaction) START
	        if (Config.recovery_scenario.equalsIgnoreCase("TRUE")) {
	            reqSheet = WebHelperUtil.getSheet(Config.controllerFilePath, "MainControlSheet");
	            controlsheet = WebHelperUtil.getValueFromHashMap(reqSheet);
	            maincontrollersheetrowCount = reqSheet.getLastRowNum() + 1;
	            in = new FileInputStream(Config.controllerFilePath);
	            
	            if (Config.controllerFilePath.endsWith(".xls")) {
	                mainWB = new HSSFWorkbook(in);
	                
	            } else if (Config.controllerFilePath.endsWith(".xlsx") | Config.controllerFilePath.endsWith(".xlsm")) {
	                mainWB = new XSSFWorkbook(in);
	            }
	            
	            MainControlSheet = mainWB.getSheet("MainControlSheet");
	            failedscenarionum = (int) sc_no.getNumericCellValue();
	            log.info("Failed Scenario is  : " + "TestcaseID:" + controllerTestCaseID + " " + "SCNO:"
	                    + failedscenarionum);
	            
	            boolean scenarioFound = false; // Flag to track if scenario is found
	            for (int rowindx = 2; rowindx < maincontrollersheetrowCount; rowindx++) {
	                currentRow = reqSheet.getRow(rowindx);
	                currentscenarionum = controlsheet.get("SC_NO");
	                currentSC_NO = (int) currentRow.getCell(currentscenarionum).getNumericCellValue();
	                
	                if (currentSC_NO == failedscenarionum && !scenarioFound) {
	                    log.info("Current MainController scenario number : " + currentSC_NO);
	                    scenarioFound = true; // Mark the scenario as found

	                    // Rest of your processing/logic goes here
	                    HashMap<String, Integer> mainControlHashMap = WebHelperUtil
	                            .getValueFromHashMap(MainControlSheet);
	                    int executeflaghashmap = mainControlHashMap.get("ExecuteFlag");
	                    Row ExecuteFlagRow = MainControlSheet.getRow(rowindx);
	                    Cell ExecuteFlagCell = ExecuteFlagRow.getCell(executeflaghashmap);

	                    int ColAccountNumber = controlsheet.get("AccountNumber");
	                    int ColPolicyNumber = controlsheet.get("PolicyNumber");
	                    int ColAgentNumber = controlsheet.get("AgentNumber");

	                    String STAccountNumber = currentRow.getCell(ColAccountNumber).getStringCellValue();
	                    String STPolicyNumber = currentRow.getCell(ColPolicyNumber).getStringCellValue();
	                    String STAgentNumber = currentRow.getCell(ColAgentNumber).getStringCellValue();

	                    log.info("-------- The data pickup for SkipRowIdentifier ---------");
	                    log.info("STAccountNumber = " + STAccountNumber);
	                    log.info("STPolicyNumber = " + STPolicyNumber);
	                    log.info("STAgentNumber = " + STAgentNumber);

	                    objFailedAccountList.add(STAccountNumber);
	                    objFailedBrokerList.add(STAgentNumber);
	                    objFailedPolicyList.add(STPolicyNumber);
	                    
	                    if (!"GroupBilling".equalsIgnoreCase(Config.productTeam)) {
	                        if (controllerTransactionType.toString().equalsIgnoreCase("SC_END")) {
	                            // ExecuteFlagCell.setCellValue("Y");
	                            WebService.RunDBQueries("XX_Scenario_Completed", STPolicyNumber, STAgentNumber,
	                                    STAccountNumber);
	                        } else {
	                        	// Now iterate over all rows again to mark all rows with the same SC_NO
	                            for (int i = 2; i < maincontrollersheetrowCount; i++) {
	                                Row checkRow = reqSheet.getRow(i);
	                                int checkSC_NO = (int) checkRow.getCell(currentscenarionum).getNumericCellValue();
	                                if (checkSC_NO == failedscenarionum) {
	                                    Row rowToUpdate = MainControlSheet.getRow(i);
	                                    Cell executeFlagCell = rowToUpdate.getCell(executeflaghashmap);
	                                    executeFlagCell.setCellValue("N");
	                                    log.info("Marked ExecuteFlag as 'N' for row with SC_NO: " + checkSC_NO);
	                                }
	                            }
	                             // Exit the outer loop after processing the failed scenario
	                        }
	                        //26/09/24
								/*
								 * ExecuteFlagCell.setCellValue("N");
								 * log.info("Marking Execution Flag for SCNO:" + " " + currentSC_NO + " " +
								 * "as 'N'. This entire scenario will stop its further execution.");
								 * WebService.RunDBQueries("XX_Transaction_Failed", STPolicyNumber,
								 * STAgentNumber, STAccountNumber);
								 */
	                        //}
	                    } else {
	                        if (controllerTransactionType.toString().equalsIgnoreCase("SC_END")) {
	                            ExecuteFlagCell.setCellValue("Y");
	                            WebService.RunDBQueries("XX_Scenario_Completed", STPolicyNumber, STAgentNumber,
	                                    STAccountNumber);
	                        }

	                    }
	                    
	                    break; 
	                  
	                }
	            }
	            
	            FileOutputStream out;
	            out = new FileOutputStream(Config.controllerFilePath);
	            mainWB.write(out);
	            out.flush();
	            out.close();
	            in.close();
	            recoverydone = true;
	            // log.info("End of Recovery");
	        }
	    } catch (Exception e) {
	        try {
	            FileOutputStream out;
	            out = new FileOutputStream(Config.controllerFilePath);
	            mainWB.write(out);
	            out.flush();
	            out.close();
	        } catch (Exception e1) {
	            log.error("Error occurred while saving file at location: " + Config.controllerFilePath);
	        }
	        log.error(e.getMessage(), e);
	        log.error(e.getLocalizedMessage());
	        
	        
	        
	    }
	}

	
	/** Pauses the Execution **/
	public void batchRecoveryScenario(String batchNo) {

		FileInputStream in = null;
		Workbook mainWB = null;
		ResultSet failedRecords = null;
		HashMap<String, Integer> controlsheet = new HashMap<String, Integer>();
		Row currentRow = null;
		int currentSC_NO = 0;
		int currentscenarionum = 0;
		int currentaccountnum = 0;
		String currentAccount_NO = null;
		int currentpolicynum = 0;
		String currentPolicy_NO = null;
		int currentagentnum = 0;
		String currentAgent_NO = null;
		//int scenarionum = 0;//Sel4

		String jobName = null;
		String PolicyNo = null;
		String brokerNumber = null;
		String AccountNumber = null;
		String brokerSystemCode = null;
		String accountSystemCode = null;
		String job_SEQ = null;
		String temp_job_SEQ = "";

		Connection conn11 = null;
		Statement st11 = null;
		try {
			conn11 = JDBCConnection.establishHTML5BillingCoreDBConn();
			st11 = conn11.createStatement();
		} catch (Exception exp) {
			log.info("Issue while creating Connection " + exp.getMessage());
		}

		try {
			Sheet reqSheet = WebHelperUtil.getSheet(Config.controllerFilePath, "MainControlSheet");
			controlsheet = WebHelperUtil.getValueFromHashMap(reqSheet);
			int maincontrollersheetrowCount = reqSheet.getLastRowNum() + 1;
			log.info("No of Rows in MainController sheet:" + maincontrollersheetrowCount);

			in = new FileInputStream(Config.controllerFilePath);
			// if (ITAFWebDriver.isClaimsApplication()) {
			if (Config.controllerFilePath.endsWith(".xls")) {
				mainWB = new HSSFWorkbook(in);
			} else if (Config.controllerFilePath.endsWith(".xlsx") | Config.controllerFilePath.endsWith(".xlsm")) {
				mainWB = new XSSFWorkbook(in);
			}
			MainControlSheet = mainWB.getSheet("MainControlSheet");
			if (Config.databaseType.equalsIgnoreCase("MsSQL")) {
				
				if ("GroupBilling".equalsIgnoreCase(Config.productTeam)) {
					failedRecords = st11.executeQuery(
							"SELECT BASE.JOB_NAME, CORE.POLICY_NO, BASE.BROKER_SYSTEM_CODE, BASE.ACCOUNT_SYSTEM_CODE,BASE.JOB_SEQ,BASE.GROUP_SYSTEM_CODE FROM "
									+ Config.jbeamdatabaseusername + ".dbo.LOG CORE, " + Config.databaseName
									+ ".dbo.JOB_SCHEDULE BASE WHERE CORE.BE_SEQ_NO=BASE.JOB_SEQ AND BASE.JOB_STATUS = 'FAILED' AND CORE.BATCH_NO='"
									+ batchNo
									+ "' AND BASE.JOB_NAME NOT IN ('ENTITY_INTERFACE','PREMIUM_INTERFACE','RCAN_INTERFACE','CHECK_UPDATE_INTERFACE','AP_INTERFACE','RETURNEDPAYMENT_INTERFACE','GL_INTERFACE','POLICY_STATUS_INTERFACE','RREIN_INTERFACE','GENERATE_FILE_NAME','OFS_MERGE_XML_JOB') AND (BASE.POLICY_NO is not null OR BASE.ACCOUNT_SYSTEM_CODE is not null OR BASE.BROKER_SYSTEM_CODE is not null OR BASE.GROUP_SYSTEM_CODE IS NOT NULL)");
					
				} else {

					failedRecords = st11.executeQuery(
							"SELECT BASE.JOB_NAME, CORE.POLICY_NO, BASE.BROKER_SYSTEM_CODE, BASE.ACCOUNT_SYSTEM_CODE,BASE.JOB_SEQ FROM "
									+ Config.jbeamdatabaseusername + ".dbo.LOG CORE, "
									+ Config.applicationdatabaseusername
									+ ".dbo.JOB_SCHEDULE BASE WHERE CORE.BE_SEQ_NO=BASE.JOB_SEQ AND BASE.JOB_STATUS = 'FAILED' AND CORE.TASK_NAME is not null AND CORE.BATCH_NO='"
									+ batchNo
									+ "' AND BASE.JOB_NAME NOT IN ('ENTITY_INTERFACE','PREMIUM_INTERFACE','RCAN_INTERFACE','CHECK_UPDATE_INTERFACE','AP_INTERFACE','RETURNEDPAYMENT_INTERFACE','GL_INTERFACE','POLICY_STATUS_INTERFACE','RREIN_INTERFACE','GENERATE_FILE_NAME','OFS_MERGE_XML_JOB') AND (BASE.POLICY_NO is not null OR BASE.ACCOUNT_SYSTEM_CODE is not null OR BASE.BROKER_SYSTEM_CODE is not null)");
				}
			} else if (Config.databaseType.equalsIgnoreCase("Oracle")) {
				
				/*
				 * failedRecords = st11.executeQuery(
				 * "SELECT BASE.JOB_NAME, CORE.POLICY_NO, BASE.BROKER_SYSTEM_CODE, BASE.ACCOUNT_SYSTEM_CODE,BASE.JOB_SEQ FROM "
				 * + Config.jbeamdatabaseusername + ".LOG CORE, " +
				 * Config.applicationdatabaseusername +
				 * ".JOB_SCHEDULE BASE WHERE CORE.BE_SEQ_NO=BASE.JOB_SEQ AND BASE.JOB_STATUS = 'FAILED' AND CORE.TASK_NAME is not null AND CORE.BATCH_NO='"
				 * + batchNo +
				 * "' AND BASE.JOB_NAME NOT IN ('ENTITY_INTERFACE','PREMIUM_INTERFACE','RCAN_INTERFACE','CHECK_UPDATE_INTERFACE','AP_INTERFACE','RETURNEDPAYMENT_INTERFACE','GL_INTERFACE','POLICY_STATUS_INTERFACE','RREIN_INTERFACE','GENERATE_FILE_NAME','OFS_MERGE_XML_JOB') AND (BASE.POLICY_NO is not null OR BASE.ACCOUNT_SYSTEM_CODE is not null OR BASE.BROKER_SYSTEM_CODE is not null)"
				 * );
				 */
			
				//Fall '25 Update 1 
				//query changed to use JBEAM_CORE and Config.databasename for Cloud env. having multiple usernames
				 
				failedRecords = st11.executeQuery(
						"SELECT BASE.JOB_NAME, CORE.POLICY_NO, BASE.BROKER_SYSTEM_CODE, BASE.ACCOUNT_SYSTEM_CODE,BASE.JOB_SEQ FROM "
								+ "JBEAM_CORE.LOG CORE, " + Config.databaseName
								+ ".JOB_SCHEDULE BASE WHERE CORE.BE_SEQ_NO=BASE.JOB_SEQ AND BASE.JOB_STATUS = 'FAILED' AND CORE.TASK_NAME is not null AND CORE.BATCH_NO='"
								+ batchNo
								+ "' AND BASE.JOB_NAME NOT IN ('ENTITY_INTERFACE','PREMIUM_INTERFACE','RCAN_INTERFACE','CHECK_UPDATE_INTERFACE','AP_INTERFACE','RETURNEDPAYMENT_INTERFACE','GL_INTERFACE','POLICY_STATUS_INTERFACE','RREIN_INTERFACE','GENERATE_FILE_NAME','OFS_MERGE_XML_JOB') AND (BASE.POLICY_NO is not null OR BASE.ACCOUNT_SYSTEM_CODE is not null OR BASE.BROKER_SYSTEM_CODE is not null)");
			} else {
				log.error("Database is not selected ");
			}
			while (failedRecords.next()) {
				jobName = failedRecords.getString(1).toString();
				try {
					PolicyNo = failedRecords.getString(2).toString();
				} catch (NullPointerException ne) {
					log.error(ne.getMessage(), ne);
					log.info("Policy# is null for Failed Job_Object " + jobName + "Batch # " + batchNo);
				}

				try {
					brokerSystemCode = failedRecords.getString(3).toString();
				} catch (NullPointerException ne) {
					log.error(ne.getMessage(), ne);
					log.info("brokerSystemCode is null for Failed Job_Object " + jobName + "Batch # " + batchNo);
				}
				if (brokerSystemCode != null) {
					ResultSet rs;
					Connection conn = JDBCConnection.establishHTML5BillingDBConn();
					Statement st = conn.createStatement();
					rs = st.executeQuery(
							"select system_entity_code from entity_register where entity_type= 'BROKER' and source_system_entity_code = '"
									+ brokerSystemCode + "'");
					while (rs.next()) {
						brokerNumber = rs.getString(1);
					}
					rs.close();
					st.close();
					JDBCConnection.closeConnection(conn);
				}
				try {
					accountSystemCode = failedRecords.getString(4).toString();
				} catch (NullPointerException ne) {
					log.error(ne.getMessage(), ne);
					log.info("accountSystemCode is null for failed Job_Object " + jobName + "Batch # " + batchNo);
				}
				if (accountSystemCode != null) {
					ResultSet rs1 = null;
					Connection conn = JDBCConnection.establishHTML5BillingDBConn();
					Statement st = conn.createStatement();
					rs1 = st.executeQuery(
							"select source_system_entity_code from entity_register where entity_type= 'ACCOUNT' and system_entity_code = '"
									+ accountSystemCode + "'");
					while (rs1.next()) {
						AccountNumber = rs1.getString(1);
					}
					rs1.close();
					st.close();
					JDBCConnection.closeConnection(conn);
				}
				job_SEQ = failedRecords.getString(5).toString();

				if ("GroupBilling".equalsIgnoreCase(Config.productTeam)) {

					for (int rowindx = 2; rowindx < maincontrollersheetrowCount; rowindx++) {
						log.info("###################### MainController Row ################## : " + rowindx);
						currentRow = reqSheet.getRow(rowindx);
						currentaccountnum = controlsheet.get("AccountNumber");
						currentAccount_NO = currentRow.getCell(currentaccountnum).toString();
						currentpolicynum = controlsheet.get("PolicyNumber");
						currentPolicy_NO = currentRow.getCell(currentpolicynum).getStringCellValue();
						currentagentnum = controlsheet.get("AgentNumber");
						currentAgent_NO = currentRow.getCell(currentagentnum).getStringCellValue();
						currentscenarionum = controlsheet.get("SC_NO");
						currentSC_NO = (int) currentRow.getCell(currentscenarionum).getNumericCellValue();
						//scenarionum = (int) sc_no.getNumericCellValue();//Sel4
						if ((PolicyNo != null && PolicyNo.equalsIgnoreCase(currentPolicy_NO))
								|| (AccountNumber != null && AccountNumber.equalsIgnoreCase(currentAccount_NO))
								|| (brokerNumber != null && brokerNumber.equalsIgnoreCase(currentAgent_NO))) {
							failedscenarionum = currentSC_NO;

							Connection conn = JDBCConnection.establishHTML5BillingDBConn();
							Statement st = conn.createStatement();
							st.execute(
									"UPDATE job_schedule  SET job_status = 'Batch_Failed_Object'  WHERE job_status <> 'COMPLETED' AND job_seq = '"
											+ job_SEQ + "'");

							if (Config.databaseType.equalsIgnoreCase("ORACLE")) {
								st.execute("commit");
							}
							st.close();
							JDBCConnection.closeConnection(conn);
							// This changes for multiple job_seq failure start
							if (!temp_job_SEQ.equalsIgnoreCase(job_SEQ)) {
								
								webDriver.getReport()
										.setMessage(webDriver.getReport().getMessage() + " " + "Job_Object " + jobName
												+ "|Batch # " + batchNo + "|job_seq " + job_SEQ + "|SC NO "
												+ failedscenarionum);
								temp_job_SEQ = job_SEQ;
							}
							// This changes for multiple job_seq failure End
							
							for (int scblockindex = 2; scblockindex < maincontrollersheetrowCount; scblockindex++) {
								HashMap<String, Integer> mainControlHashMap = WebHelperUtil
										.getValueFromHashMap(MainControlSheet);
								int executeflaghashmap = mainControlHashMap.get("ExecuteFlag");
								currentRow = reqSheet.getRow(scblockindex);
								currentaccountnum = controlsheet.get("AccountNumber");
								currentAccount_NO = currentRow.getCell(currentaccountnum).toString();
								currentpolicynum = controlsheet.get("PolicyNumber");
								currentPolicy_NO = currentRow.getCell(currentpolicynum).getStringCellValue();
								currentagentnum = controlsheet.get("AgentNumber");
								currentAgent_NO = currentRow.getCell(currentagentnum).getStringCellValue();
								currentscenarionum = controlsheet.get("SC_NO");
								Row ExecuteFlagRow = MainControlSheet.getRow(scblockindex);
								Cell ExecuteFlagCell = ExecuteFlagRow.getCell(executeflaghashmap);

								// Add if condition for checking scenario number
								// and block all policies, account and broker
								currentRow = reqSheet.getRow(scblockindex);
								currentscenarionum = controlsheet.get("SC_NO");
								// log.info("currentscenarionum : "+currentscenarionum);
								currentSC_NO = (int) currentRow.getCell(currentscenarionum).getNumericCellValue();
								if (failedscenarionum == currentSC_NO) {
									ExecuteFlagCell.setCellValue("N");
									log.info("Row no" + scblockindex + "out of MainController rows"
											+ maincontrollersheetrowCount);
									// log.info("currentscenarionum blocked: "+currentscenarionum);
									log.info("currentscenarionum blocked: " + currentSC_NO);
									// Changes done by Aniruddha
									Connection conn1 = JDBCConnection.establishHTML5BillingDBConn();
									Statement st1 = conn1.createStatement();
									if (currentPolicy_NO != null) {
										st1.execute(
												"UPDATE job_schedule  SET job_status = 'XX_Obejct_Failed'  WHERE job_status <> 'COMPLETED' AND policy_term_id IN (select policy_term_id from policy_register where policy_no = '"
														+ PolicyNo + "')");
										log.info("Marked objects for Policy_NO : " + currentPolicy_NO);
									}
									if (currentAgent_NO != null) {
										st1.execute(
												"UPDATE job_schedule SET job_status = 'XX_Obejct_Failed' WHERE job_status <> 'COMPLETED' AND account_system_code = (select system_entity_code from entity_register where entity_type= 'BROKER' and source_system_entity_code = '"
														+ brokerNumber + "')");
										log.info("Marked objects for Agent_NO : " + currentAgent_NO);
									}
									if (currentAccount_NO != null) {
										st1.execute(
												"UPDATE job_schedule SET job_status = 'XX_Obejct_Failed'  WHERE job_status <> 'COMPLETED' AND account_system_code = (select system_entity_code from entity_register where entity_type= 'ACCOUNT' and source_system_entity_code = '"
														+ AccountNumber + "')");
										log.info("Marked objects for Account_NO : " + currentAccount_NO);
									}
									// rowindx = scblockindex;
									if (Config.databaseType.equalsIgnoreCase("ORACLE")) {
										st.execute("commit");
									}
									st1.close();
									JDBCConnection.closeConnection(conn1);
								}
							}
						}
					}
				} else {
					// ***GB Change - Recovery for GB*** 26/07/2018
					for (int rowindx = 2; rowindx < maincontrollersheetrowCount; rowindx++) {

						if (!temp_job_SEQ.equalsIgnoreCase(job_SEQ)) {

							webDriver.getReport().setMessage(webDriver.getReport().getMessage() + " " + "Job_Object "
									+ jobName + "|Batch # " + batchNo + "|job_seq " + job_SEQ);
							temp_job_SEQ = job_SEQ;
							// ((JavascriptExecutor)
							// Automation.driver).executeScript("alert('batch Ojects failed');");//Mandar
						}
						
					}
				}
			}
		} catch (SQLException se) {
			log.error(se.getMessage(), se);
			log.error(se.getLocalizedMessage());
		} catch (ClassNotFoundException ce) {
			log.error(ce.getMessage(), ce);
			log.error(ce.getLocalizedMessage());
		} catch (IOException io) {
			log.error(io.getMessage(), io);
			log.error(io.getLocalizedMessage());
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			log.error(e.getLocalizedMessage());
		} finally {
			try { // changes for saving main controller 
				failedRecords.close();
				st11.close();
				JDBCConnection.closeConnection(conn11);
				FileOutputStream out;
				out = new FileOutputStream(Config.controllerFilePath);
				mainWB.write(out);
				out.flush();
				out.close();
				in.close();
			} catch (Exception e) {
				log.error(e.getMessage(), e);
				log.error("Exception thrown while saving main controller  after marking SC" + failedscenarionum
						+ " in file  <-|-> LocalizeMessage " + e.getLocalizedMessage() + " <-|-> Message "
						+ e.getMessage() + " <-|-> Cause " + e.getCause());
			}
		}
	}

	public boolean pauseFun(String message) {

		/**
		 * DS:18-07-2014:Replacing timeout in msg String tempMsg = "Timed out after
		 * CONFIGTIMEOUT seconds waiting for presence of element located by" ; tempMsg =
		 * tempMsg.replace("CONFIGTIMEOUT", Config.timeOut.toString()); if(message!=
		 * null) message = message.replace(tempMsg, "Element not found");
		 **/

		String scrshotbody = "";
		String strScrCompleteData = "";
		String scrEmail = "";
		String scrreportsheet = "";
		String scrdetailreportsheet = "";
		String userInteraction = "TRUE";

		log.info("In PauseFun method");
		webDriver.getReport().setGroupName(controllerGroupName.toString());
		webDriver.getReport().setTestcaseId(controllerTestCaseID.toString());
		webDriver.getReport().setTestDescription(testDescription);
		webDriver.getReport().setTrasactionType(controllerTransactionType.toString());
		// WebHelperBilling.toDate = new Date();
		webDriver.getReport().setMessage(message);
		webDriver.getReport().setToDate(Config.dtFormat.format(new Date()));

		if (!(controllerTransactionType.toString().startsWith("WebService"))) {
			WebHelperUtil.saveScreenShot();
		}
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
				JOptionPane.showConfirmDialog(webDriver.getFrame(), "Null Value Found for UserInteractioin Parameter");
			}
		}

		/** Don't mark status as FAIL if transaction name is PAUSE **/
		if (!controllerTransactionType.toString().equalsIgnoreCase("PAUSE")) {
			log.info("Start Recovery");
			webDriver.getReport().setStatus("FAIL");
			webDriver.getReport().setCycleDate(cycleDateCellValue);

			recoveryhandler();
			log.info("End Recovery");
			scrshotbody = "PFA the Screenshot for failed Transaction";
			scrEmail = Config.emailId;
			if (!scrEmail.toString().equalsIgnoreCase("NA") && !scrEmail.equalsIgnoreCase("")
					&& !scrEmail.equalsIgnoreCase(null)) {
				strScrCompleteData = scrEmail + "#" + scrshotbody + "#" + scrreportsheet + "#" + scrdetailreportsheet
						+ "#" + FailScreen + "#" + "";
				try {
					Runtime.getRuntime().exec("wscript SendMail.vbs " + (char) 34 + strScrCompleteData + (char) 34);
				} catch (IOException e) {
					// TODO Auto-generated catch block
					log.error(e.getMessage(), e);
					e.printStackTrace();
				}
			}

		}

		if (!userInteraction.equalsIgnoreCase("FALSE")) {
			webDriver.getFrame().setVisible(true);
			webDriver.getFrame().setAlwaysOnTop(true);
			webDriver.getFrame().setLocationRelativeTo(null);

			JOptionPane.setRootFrame(webDriver.getFrame());
			int response = JOptionPane.showConfirmDialog(webDriver.getFrame(), message, "iTAF - Do you want to STOP...",
					JOptionPane.YES_NO_OPTION);
			// log.info(response);
			if (response == JOptionPane.YES_OPTION) {
				// pauseExecution = true;
				pExecution = true;
			} else if (response == 1) {

			} else {
				log.info("You have pressed cancel" + response);
				// pauseExecution =true;
			}
		} else {
			webDriver.getReport().setMessage(message);
			// pauseExecution = true;
		}
		return pauseExecution;

	}
}
