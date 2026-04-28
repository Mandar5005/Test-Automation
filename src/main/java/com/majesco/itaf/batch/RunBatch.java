package com.majesco.itaf.batch;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Locale;
import javax.swing.JOptionPane;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.ITAFWebDriver;
import com.majesco.itaf.main.MainController;
import com.majesco.itaf.main.MainControllerSuite;
import com.majesco.itaf.main.ObjectFactory;
import com.majesco.itaf.main.WebHelperUtil;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.util.JDBCConnection;
import com.majesco.itaf.vo.Reporter;
import com.majesco.itaf.webservice.WebService;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class RunBatch {

	private final static Logger log = LogManager.getLogger(RunBatch.class.getName());
	public static String XMLDataFile = null;
	public static String Business_Date = null;
	public static String Batch_Date = null;
	public static String tempDate = null;
	public static String PDC_Count = null;
	public static String BatchExecutionDate = null;
	public static String Job_Status = null;
	public static String sWSDL_URL = null;
	public static String sContent_Type = null;
	public static String batchNo = null;
	public static int Batch_No = 0;
	public static HashMap<String, Object> BatchHashMap = new HashMap<String, Object>();
	public static String Batch_Status = null;
	public static String Batch_End = null;
	public static int batchTimer = 1;
	public static String stopBatch = null;
	public static int retVal = 0;
	public static String strBatchCompleteData = null;
	public static HashMap<String, Integer> HeaderValues = new HashMap<String, Integer>();
	private static boolean isMsSQLDB = Config.databaseType.equalsIgnoreCase("MsSQL");
	private static boolean isOracleDB = Config.databaseType.equalsIgnoreCase("Oracle");
	public static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	private static MainController controller = ObjectFactory.getMainController();
	private static String batchEndReasonQry = "Select BATCH_END_REASON from BATCH where BATCH_NO IN (Select MAX(BATCH_NO) from BATCH where INSTRUCTION_SEQ_NO = ";
	private static String batchEndReasonQryNew = "Select BATCH_END_REASON from JBEAM_CORE.BATCH where BATCH_NO IN (Select MAX(BATCH_NO) from JBEAM_CORE.BATCH where INSTRUCTION_SEQ_NO = ";
	private static Connection billingDBCon;
	private static Connection billingCoreDBCon;

	public static Reporter RunBatch_WS() {
		return runBatch(true);
	}

	public static Reporter RunBatch_WSFF() {
		return runBatch(false);
	}

	private static Reporter runBatch(boolean considerPDCCount) {
		Date startDate = new Date();
		int Job_Status_count = 0;
		Batch_End = null;
		webDriver.getReport().setFromDate(Config.dtFormat.format(startDate));
		webDriver.getReport().setIteration(Config.cycleNumber);
		webDriver.getReport().setTestcaseId(controller.controllerTestCaseID.toString());
		webDriver.getReport().setGroupName(controller.controllerGroupName.toString());
		webDriver.getReport().setTrasactionType(controller.controllerTransactionType.toString());
		webDriver.getReport().setCycleDate(controller.cycleDateCellValue);
		webDriver.getReport().setTestDescription(controller.testDescription);
		Reporter report = new Reporter();

		// For Day and Night Batch//
		String batch_trans_type = controller.controllerTransactionType.toString();
		String testCaseId = controller.controllerTestCaseID.toString();

		if (batch_trans_type.equalsIgnoreCase("RunBatch_JBEAM_NIGHT")) {
			XMLDataFile = Config.runbatchxmlpath_night;
		} else if (batch_trans_type.equalsIgnoreCase("RunBatch_JBEAM_DAY")) {
			XMLDataFile = Config.runbatchxmlpath_day;
		} else {
			XMLDataFile = Config.runbatchxmlpath;
		}
		if (!StringUtils.equalsIgnoreCase(Config.executionApproach, "linear")) {
			Business_Date = controller.cycleDateValue;

			log.info("Business Date to start run batch :" + Business_Date);
			Calendar calendar = Calendar.getInstance();
			try {
				calendar.setTime(new SimpleDateFormat("MM/dd/yyyy").parse(Business_Date));
			} catch (ParseException e1) {
				log.error(e1.getLocalizedMessage(), e1);
				webDriver.getReport().setStatus("FAIL");
				webDriver.getReport().setMessage(e1.getLocalizedMessage());
			}
			int weekday = calendar.get(Calendar.DAY_OF_WEEK);

			log.info("day of the week is : " + weekday);
			log.info("Modified Business Date : " + Business_Date);
			log.info("batch_trans_type : " + batch_trans_type);
		}
		BatchHashMap.put("STOP_BATCH", "FALSE");
		try {

			billingDBCon = JDBCConnection.establishHTML5BillingDBConn();
			billingCoreDBCon = JDBCConnection.establishHTML5BillingCoreDBConn();

			if (batch_trans_type.equalsIgnoreCase("RunBatch_JBEAM_NIGHT")) {
				Batch_Date = getNextBatchDate();
				Business_Date = getNextBusinessDate();
			} else if (batch_trans_type.equalsIgnoreCase("RunBatch_JBEAM_DAY")) {
				Batch_Date = getBatchDate();
				Business_Date = getBusinessDate();

			} else if (batch_trans_type.equalsIgnoreCase("RunBatch_JBEAMFF")) {

				if (testCaseId.equalsIgnoreCase("RunBatch_JBEAM_NIGHT")) {
					Batch_Date = getNextBatchDate();
					Business_Date = getNextBusinessDate();
					log.info("RunBatch_JBEAM_NIGHT:RunBatch_JBEAMFF:Business_Date : " + Business_Date);

				} else if (testCaseId.equalsIgnoreCase("RunBatch_JBEAM_DAY")) {
					Batch_Date = getBatchDate();
					Business_Date = getBusinessDate();
					log.info("RunBatch_JBEAM_DAY:RunBatch_JBEAMFF:Business_Date : " + Business_Date);
				} else {
					Batch_Date = getNextBatchDate();
					Business_Date = getNextBusinessDate();
					log.info("RunBatch_JBEAM_Default:RunBatch_JBEAMFF:Business_Date : " + Business_Date);
				}
			} else if (batch_trans_type.equalsIgnoreCase("RunBatch_JBEAM")) {

				if (testCaseId.equalsIgnoreCase("RunBatch_JBEAM_NIGHT")) {
					Batch_Date = getNextBatchDate();
					Business_Date = getNextBusinessDate();
					log.info("RunBatch_JBEAM_NIGHT:RunBatch_JBEAM:Business_Date : " + Business_Date);

				} else if (testCaseId.equalsIgnoreCase("RunBatch_JBEAM_DAY")) {
					Batch_Date = getBatchDate();
					Business_Date = getBusinessDate();
					log.info("RunBatch_JBEAM_DAY:RunBatch_JBEAM:Business_Date : " + Business_Date);
				} else {
					Batch_Date = getNextBatchDate();
					Business_Date = getNextBusinessDate();
					log.info("RunBatch_JBEAM_Default:RunBatch_JBEAM:Business_Date : " + Business_Date);
				}
			}

			log.info("Modified Batch Date : " + Batch_Date);

			BatchExecutionDate = Business_Date;

			log.info("BatchExecutionDate : " + BatchExecutionDate);

			Job_Status_count = getBatchJobCount();

			if (considerPDCCount) {
				int PDC_Count_Num = getPDCCount();

				if (Job_Status_count == 0 && PDC_Count_Num == 0) {
					log.info("Batch object as well as PDC object Not Found for the BatchExecutionDate : "
							+ BatchExecutionDate);
					webDriver.getReport().setStatus("PASS");
					webDriver.getReport()
							.setMessage("Batch object as well as PDC object Not Found for the BatchExecutionDate : "
									+ BatchExecutionDate);
					return report;
				} else if (Job_Status_count > 0 || PDC_Count_Num > 0) {
					executeBatch();
				}
			} else {

				executeBatch();

			}
		} catch (SQLException sqle) {
			log.error(sqle.getLocalizedMessage(), sqle);
			webDriver.getReport().setStatus("FAIL");
			webDriver.getReport().setMessage(sqle.getLocalizedMessage());
		} catch (ClassNotFoundException cnfe) {
			log.error(cnfe.getLocalizedMessage(), cnfe);
			webDriver.getReport().setStatus("FAIL");
			webDriver.getReport().setMessage(cnfe.getLocalizedMessage());
		} catch (IOException ioe) {
			log.error(ioe.getLocalizedMessage(), ioe);
			webDriver.getReport().setStatus("FAIL");
			webDriver.getReport().setMessage(ioe.getLocalizedMessage());
		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
			webDriver.getReport().setStatus("FAIL");
			webDriver.getReport().setMessage(e.getLocalizedMessage());
		} finally {
			if (null != billingDBCon)
				JDBCConnection.closeConnection(billingDBCon);

			if (null != billingCoreDBCon)
				JDBCConnection.closeConnection(billingCoreDBCon);
		}

		return report;
	}

	@SuppressWarnings("deprecation")
	private static void executeBatch() throws ClassNotFoundException, SQLException, Exception {
		String sequenceNo = null;
		String errorDescription = "";
		int waitcount = 0;

		DateFormat simpleDateFormat = new SimpleDateFormat("MM-dd-yyyy", Locale.ENGLISH);
		Date newBatchDate = simpleDateFormat.parse(Batch_Date);
		DateFormat cycleDateFormat = new SimpleDateFormat("MM/dd/yyyy");
		Batch_Date = cycleDateFormat.format(newBatchDate);

		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document doc = dBuilder.parse(XMLDataFile);
		doc.getDocumentElement().normalize();

		Node configuration = doc.getElementsByTagName("instructionParameters").item(2);
		NodeList list = configuration.getChildNodes();

		String NodeValue = "value";
		for (int i = 0; i < list.getLength(); i++) {
			Node node = list.item(i);

			if (NodeValue.equals(node.getNodeName())) {
				node.setTextContent(Batch_Date + " 23:59:59");
				break;
			}
		}

		TransformerFactory transformerFactory = TransformerFactory.newInstance();
		Transformer transformer = transformerFactory.newTransformer();
		DOMSource source = new DOMSource(doc);
		StreamResult result = new StreamResult(new File(XMLDataFile));
		transformer.transform(source, result);

		if (!StringUtils.equalsIgnoreCase(Config.executionApproach, "linear")) {
			Sheet valuesSheet = ExcelUtility.GetSheet(Config.runbatchinputsheetpath, "Values");
			if (HeaderValues.isEmpty() == true) {
				HeaderValues = WebHelperUtil.getValueFromHashMap(valuesSheet);
			}

			Row currentRow = valuesSheet.getRow(1);

			sWSDL_URL = currentRow.getCell(HeaderValues.get("WSDL_URL")).toString();
			sContent_Type = currentRow.getCell(HeaderValues.get("ContentType")).toString();
		} else {
			sWSDL_URL = Config.jbeamWsdlUrl;
			sContent_Type = Config.jbeamContentType;
		}

		String sRequestXML = XMLDataFile;
		String sResponseXML = null;
		Boolean isServiceHit = false;
		try {
			sResponseXML = WebService.callWebService_JBeam(sWSDL_URL, sRequestXML.toString(), sContent_Type);
			isServiceHit = true;
			sequenceNo = WebService.getXMLResponseTagValue(sResponseXML, "batchDetails", "instructionSeqNo", 0);
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			if (isServiceHit)
				errorDescription = WebService.getXMLResponseTagValue(sResponseXML, "*", "description", 0);
			else {
				controller.stopExecution = true;
				controller.pauseFun(
						"Failure in batch execution. Recheck parameters - JBEAM_USERID, JBEAM_PASSWORD, JBEAMDAYBATCHINPUT excel -> WSDL_URL, Batch XML file -> INSTALLATION_CODE."
								+ e.getMessage());
				if (ITAFWebDriver.isSuiteApplication())
					MainControllerSuite.reportWritten = true;
			}
		}
		log.info("Sequence Number of Response XML is: " + sequenceNo);

		if ((sequenceNo != null) == true) {
			waitcount = 0;
			webDriver.getReport().setMessage("Instruction Sequence No is : " + sequenceNo); // Mandar

			String tableName = "BATCH";
			if ((Config.databaseAdminaccess != null) && (Config.databaseAdminaccess.equalsIgnoreCase("FALSE"))) {
				tableName = "JBEAM_CORE.BATCH";
			}

			do {
				batchNo = executeQueryOnCoreDB(
						"Select BATCH_NO from " + tableName + " where INSTRUCTION_SEQ_NO =" + sequenceNo);
				log.info("Batch No is : " + batchNo);

				try {
					Thread.sleep(1000);
				} catch (InterruptedException e) {
					log.error(e.getMessage(), e);
					Thread.currentThread().interrupt();
				}
				waitcount++;
			} while (batchNo == null && waitcount <= 60);

			if (batchNo == null) {
				controller.pExecution = true;
				controller.pauseExecution = true;
				controller.stopExecution = true;
				log.info("Instruction sequence number is not generated, Execution is going to stop");
				log.info(errorDescription);
				BatchHashMap.put("STOP_BATCH", "TRUE");
				String BatchError = "Instruction Sequence number is not generated, Framework stops execution";
				if (!(webDriver.getEmail().equalsIgnoreCase("NA")) && !(webDriver.getEmail().equalsIgnoreCase(""))
						&& !(webDriver.getEmail().equalsIgnoreCase(null))) {
					strBatchCompleteData = webDriver.getEmail() + "#" + BatchError + "#"
							+ "Batch No not Fetched for sequence No:" + sequenceNo;
					Runtime.getRuntime()
							.exec("wscript BatchFailMail.vbs " + (char) 34 + strBatchCompleteData + (char) 34);
				} else {
					webDriver.getFrame().setVisible(true);
					webDriver.getFrame().setAlwaysOnTop(true);
					webDriver.getFrame().setLocationRelativeTo(null);
					JOptionPane.setRootFrame(webDriver.getFrame());
					JOptionPane.showMessageDialog(webDriver.getFrame(),
							"Batch not working and email id not configured in config sheet");
					controller.pExecution = true;
					controller.pauseExecution = true;
					webDriver.getFrame().dispose();
				}
			}

			stopBatch = BatchHashMap.get("STOP_BATCH").toString();

			boolean bBatchRunning = true;
			batchTimer = 0;
			log.info("Waiting for Batch to complete...This will take few minutes");

			// Pick query based on Config.databaseAdminaccess//21/08/2025 - For USIC and
			// SELECTIVE Projects
			String activeBatchEndReasonQry;
			if ((Config.databaseAdminaccess != null) && (Config.databaseAdminaccess.equalsIgnoreCase("FALSE"))) {
				activeBatchEndReasonQry = batchEndReasonQryNew; // use JBEAM_CORE.BATCH
			} else {
				activeBatchEndReasonQry = batchEndReasonQry; // use default BATCH
			}

			if (!stopBatch.equalsIgnoreCase("TRUE")) {
				// while (bBatchRunning || batchTimer <= 60 ||
				// !stopBatch.equalsIgnoreCase("TRUE")) {
				while (bBatchRunning || batchTimer <= 60) {
					// Batch_End = executeQueryOnCoreDB(batchEndReasonQry + sequenceNo + ")");
					// //21/08/2025
					Batch_End = executeQueryOnCoreDB(activeBatchEndReasonQry + sequenceNo + ")");
					//Batch_End = executeQueryOnCoreDB(batchEndReasonQry + sequenceNo + ")");
					if (Batch_End != null && Batch_End.trim().length() > 0) {
						break;
					}
					Thread.sleep(1000 * 3);
					batchTimer = batchTimer + 1;
				}
			}
			log.info("Batch End Reason : " + Batch_End);

			if (StringUtils.equalsIgnoreCase(Batch_End, "USER_INTERRUPTED")) {
				log.info("Batch End Reason Fetched from Database : " + Batch_End);
				webDriver.getReport().setStatus("BATCH USER INTERRUPTED");
			} else if (StringUtils.equalsIgnoreCase(Batch_End, "BATCH_COMPLETED")) {
				webDriver.getReport().setStatus("PASS");
				log.info("Batch End Reason Fetched from Database : " + Batch_End);
				webDriver.getReport().setMessage("Batch Completed -- Batch No is :" + batchNo);// Mandar
			} else {
				webDriver.getReport().setStatus("FAIL");
				log.info("Batch End Reason Fetched from Database : " + Batch_End);
				webDriver.getReport().setMessage("Batch Failed -- Batch No is :" + batchNo);// Mandar
			}

			if (StringUtils.equalsIgnoreCase(Config.retrieveFailedObjectList, "true")
					|| Config.retrieveFailedObjectList == null) {
				String failedRecords = getFailedRecords();
				if (failedRecords != null) {
					if (!StringUtils.equalsIgnoreCase(Config.executionApproach, "linear")) {
						controller.batchRecoveryScenario(batchNo);
					} else {
						webDriver.getReport().setStatus("FAIL");
						webDriver.getReport().setMessage("Batch Failed -- Batch No is :" + batchNo);
					}
				}
			}
		} else {
			log.info("Instruction sequence number is not generated, Execution is going to stop");
			BatchHashMap.put("STOP_BATCH", "TRUE");
			webDriver.getReport().setStatus("FAIL");
			webDriver.getReport().setMessage(errorDescription);
			// webDriver.getReport().setMessage(("Instruction sequence number is not
			// generated, Execution is going to stop"));
			controller.pExecution = true;
			controller.pauseExecution = true;
			controller.stopExecution = true;
		}
	}

	private static String getNextBatchDate() throws ClassNotFoundException, SQLException, Exception {
		String tableName = "business_day";
		String query = null;
		if ((Config.databaseAdminaccess != null) && (Config.databaseAdminaccess.equalsIgnoreCase("FALSE"))) {
			tableName = "BILLING_CONTENT.business_day";
		}

		if (isMsSQLDB) {
			query = "SELECT CONVERT(VARCHAR(10), CAST(NEXT_BUSINESS_DATE AS DATETIME), 110) FROM " + tableName;
		} else if (isOracleDB) {
			query = "SELECT TO_CHAR(NEXT_BUSINESS_DATE, 'mm-dd-yyyy') FROM " + tableName;
		} else {
			throw new RuntimeException("No database selected");
		}

		return executeQueryOnBillingDB(query);
	}

	private static String getBatchDate() throws ClassNotFoundException, SQLException, Exception {
		String query = null;
		String tableName = "business_day";
		if ((Config.databaseAdminaccess != null) && (Config.databaseAdminaccess.equalsIgnoreCase("FALSE"))) {
			tableName = "BILLING_CONTENT.business_day";
		}
		if (isMsSQLDB) {
			query = "Select convert(varchar(10),CAST(BUSINESS_DATE AS datetime),110) FROM " + tableName;// Meghna--Selecting
			// next_business_date
		} else if (isOracleDB) {
			query = "Select TO_CHAR(business_date,'mm-dd-yyyy') FROM " + tableName; // Meghna--Selecting
			// next_business_date
		} else {
			throw new RuntimeException("No database selected");
		}

		return executeQueryOnBillingDB(query);
	}

	private static String getBusinessDate() throws ClassNotFoundException, SQLException, Exception {
		String query = null;
		String tableName = "business_day";
		if ((Config.databaseAdminaccess != null) && (Config.databaseAdminaccess.equalsIgnoreCase("FALSE"))) {
			tableName = "BILLING_CONTENT.business_day";
		}
		if (isMsSQLDB) {
			query = "Select Replace(convert(varchar(10),CAST(BUSINESS_DATE AS datetime),6),' ','-') FROM " + tableName;
		} else if (isOracleDB) {
			query = "Select TO_CHAR(business_date,'dd-mon-yy') FROM " + tableName;
			// next_business_date
		}
		return executeQueryOnBillingDB(query);
	}

	private static String getNextBusinessDate() throws ClassNotFoundException, SQLException, Exception {
		String query = null;
		String tableName = "business_day";
		if ((Config.databaseAdminaccess != null) && (Config.databaseAdminaccess.equalsIgnoreCase("FALSE"))) {
			tableName = "BILLING_CONTENT.business_day";
		}
		if (isMsSQLDB) {
			query = "Select Replace(convert(varchar(10),CAST(NEXT_BUSINESS_DATE AS datetime),6),' ','-') FROM "
					+ tableName;// Mandar

		} else if (isOracleDB) {
			query = "Select TO_CHAR(next_business_date,'dd-mon-yy') FROM " + tableName; // Meghna--Selecting

		}
		return executeQueryOnBillingDB(query);
	}

	private static int getBatchJobCount() {
		String query = null;
		String tableName = "job_schedule";
		int jobCount = 0;
		if ((Config.databaseAdminaccess != null) && (Config.databaseAdminaccess.equalsIgnoreCase("FALSE"))) {
			tableName = "BILLING_CONTENT.job_schedule";
		}
		if (isMsSQLDB) {
			query = "Select count(*) from " + tableName
					+ " where Replace(convert(varchar(10),CAST(EXECUTION_DATE AS date),6),' ','-') <= CAST(" + "'"
					+ BatchExecutionDate + "' AS DATE)" + " " + "and" + " job_status = 'SCHEDULED'";
		} else if (isOracleDB) {
			query = "select count(*) from " + tableName + " where trunc(execution_date) <= " + "'" + BatchExecutionDate
					+ "'" + " " + "and" + " job_status = 'SCHEDULED'";
		}
		log.info("Query:" + query);

		try {
			Job_Status = executeQueryOnBillingDB(query);
			jobCount = Integer.parseInt(Job_Status);
			log.info(Job_Status + " number of object are indetified for batch | execution date " + BatchExecutionDate);
		} catch (Exception e) {
			log.error("Job_Status returns null", e);
			jobCount = 0;
		}
		return jobCount;
	}

	private static int getPDCCount() {
		String query = null;
		int PDC_Count_Num = 0;
		String tableName = "payment_batch";
		if ((Config.databaseAdminaccess != null) && (Config.databaseAdminaccess.equalsIgnoreCase("FALSE"))) {
			tableName = "BILLING_CONTENT.payment_batch";
		}

		if (isMsSQLDB) {
			query = "Select count(*) from " + tableName
					+ " where Replace(convert(varchar(10),CAST(deposit_date AS date),6),' ','-') <= CAST(" + "'"
					+ BatchExecutionDate + "' AS DATE)" + " and payment_batch_Status= 'OPEN'";
		} // 17-May 2017 : to add MSSql query : End
		else if (isOracleDB) {
			query = "select count(*) from " + tableName + " where pdc_flag = 'Y' and deposit_date <= '"
					+ BatchExecutionDate + "' and payment_batch_Status = 'OPEN'";
		}

		try {
			PDC_Count = executeQueryOnBillingDB(query);
			PDC_Count_Num = Integer.parseInt(PDC_Count);
			log.info(PDC_Count + " number of PDC objects are indetified for batch | execution date "
					+ BatchExecutionDate);
		} catch (Exception e) {
			log.error("PDC_Count returns null", e);
			PDC_Count_Num = 0;
		}
		return PDC_Count_Num;
	}

	private static String getFailedRecords() throws ClassNotFoundException, SQLException, Exception {
		String query = null, failedRecords = null;

		// *** Default values from config - FOR USIC and SELECTIVE projects - 21/08/2025
		String jbeamSchema = Config.jbeamdatabaseusername;
		String billingSchema = Config.databaseName;

		// Override schemas if admin access is FALSE
		if ((Config.databaseAdminaccess != null) && (Config.databaseAdminaccess.equalsIgnoreCase("FALSE"))) {
			jbeamSchema = "JBEAM_CORE";
			billingSchema = "BILLING_CONTENT";
		}
		// ***

		if (isMsSQLDB) {
			if ("GroupBilling".equalsIgnoreCase(Config.productTeam)) {
				query = "SELECT CORE.BATCH_NO, BASE.JOB_NAME, CORE.POLICY_NO, CORE.BROKER, BASE.ACCOUNT_SYSTEM_CODE, BASE.BROKER_SYSTEM_CODE,BASE.GROUP_SYSTEM_CODE FROM "
						+ jbeamSchema + ".dbo.LOG CORE, " + billingSchema
						+ ".dbo.JOB_SCHEDULE BASE WHERE CORE.BE_SEQ_NO=BASE.JOB_SEQ AND BASE.JOB_STATUS = 'FAILED' AND CORE.BATCH_NO='"
						+ batchNo
						+ "' AND BASE.JOB_NAME NOT IN ('ENTITY_INTERFACE','PREMIUM_INTERFACE','RCAN_INTERFACE','CHECK_UPDATE_INTERFACE','AP_INTERFACE','RETURNEDPAYMENT_INTERFACE','GL_INTERFACE','POLICY_STATUS_INTERFACE','RREIN_INTERFACE','GENERATE_FILE_NAME','OFS_MERGE_XML_JOB') AND (BASE.POLICY_NO is not null OR BASE.ACCOUNT_SYSTEM_CODE is not null OR BASE.BROKER_SYSTEM_CODE is not null OR BASE.GROUP_SYSTEM_CODE IS NOT NULL)";
				failedRecords = executeQueryOnCoreDB(query);
			} else {
				query = "SELECT CORE.BATCH_NO, BASE.JOB_NAME, CORE.POLICY_NO, CORE.BROKER, BASE.ACCOUNT_SYSTEM_CODE, BASE.BROKER_SYSTEM_CODE FROM "
						+ jbeamSchema + ".dbo.LOG CORE, " + billingSchema
						+ ".dbo.JOB_SCHEDULE BASE WHERE CORE.BE_SEQ_NO=BASE.JOB_SEQ AND BASE.JOB_STATUS = 'FAILED' AND CORE.TASK_NAME is not null AND CORE.BATCH_NO='"
						+ batchNo
						+ "' AND BASE.JOB_NAME NOT IN ('ENTITY_INTERFACE','PREMIUM_INTERFACE','RCAN_INTERFACE','CHECK_UPDATE_INTERFACE','AP_INTERFACE','RETURNEDPAYMENT_INTERFACE','GL_INTERFACE','POLICY_STATUS_INTERFACE','RREIN_INTERFACE','GENERATE_FILE_NAME','OFS_MERGE_XML_JOB') AND (BASE.POLICY_NO is not null OR BASE.ACCOUNT_SYSTEM_CODE is not null OR BASE.BROKER_SYSTEM_CODE is not null)";
				failedRecords = executeQueryOnCoreDB(query);

			}
		} else if (isOracleDB) {
			query = "SELECT CORE.BATCH_NO, BASE.JOB_NAME, CORE.POLICY_NO, CORE.BROKER, BASE.ACCOUNT_SYSTEM_CODE, BASE.BROKER_SYSTEM_CODE FROM "
					+ jbeamSchema + ".LOG CORE, " + billingSchema
					+ ".JOB_SCHEDULE BASE WHERE CORE.BE_SEQ_NO=BASE.JOB_SEQ AND BASE.JOB_STATUS = 'FAILED' AND CORE.TASK_NAME is not null AND CORE.BATCH_NO='"
					+ batchNo
					+ "' AND BASE.JOB_NAME NOT IN ('ENTITY_INTERFACE','PREMIUM_INTERFACE','RCAN_INTERFACE','CHECK_UPDATE_INTERFACE','AP_INTERFACE','RETURNEDPAYMENT_INTERFACE','GL_INTERFACE','POLICY_STATUS_INTERFACE','RREIN_INTERFACE','GENERATE_FILE_NAME','OFS_MERGE_XML_JOB') AND (BASE.POLICY_NO is not null OR BASE.ACCOUNT_SYSTEM_CODE is not null OR BASE.BROKER_SYSTEM_CODE is not null)";// Mandar

			failedRecords = executeQueryOnCoreDB(query);
		}
		return failedRecords;
	}

	private static String executeQueryOnCoreDB(String query) throws ClassNotFoundException, SQLException, Exception {
		ResultSet rs = null;
		String returnVale = null;

		Statement st = billingCoreDBCon.createStatement();

		rs = st.executeQuery(query);
		if (rs.next()) {
			returnVale = rs.getString(1);
		}
		st.close();
		rs.close();

		return returnVale;
	}

	private static String executeQueryOnBillingDB(String query) throws ClassNotFoundException, SQLException, Exception {
		String returnVale = null;
		ResultSet rs;

		Statement st = billingDBCon.createStatement();
		rs = st.executeQuery(query);
		if (rs.next()) {
			returnVale = rs.getString(1);
		}
		rs.close();
		st.close();

		return returnVale;
	}

}
