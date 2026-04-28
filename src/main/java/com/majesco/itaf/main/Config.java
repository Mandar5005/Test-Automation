package com.majesco.itaf.main;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * This class is responsible for loading all the configurations from Config xls
 * and making them available to the application to access.
 * 
 * 
 * 
 */
public class Config {

	private final static Logger log = LogManager.getLogger(Config.class);
	private static HashMap<String, Object> configHashMap = new HashMap<String, Object>();
	public static String projectName = "";
	public static String seleniumGrid = "";// Selenium 4
	public static String incognitoMode = "";// Selenium 4
	public static String environment = null;// Selenium 4
	public static String productTeam = null;
	public static String SikuliScreenValue = null;
	public static String SikuliScr = null;
	public static String runbatchinputsheetpath = null;
	public static String errorlog = "";
	public static String applicationdatabaseusername = null;
	public static String applicationdatabasepassword = null;
	public static String applicationdatabaseusername_pas = null, applicationdatabasepassword_pas = null,
			applicationdatabaseusername_claims = null, applicationdatabasepassword_claims = null,
			applicationdatabaseusername_billing = null, applicationdatabasepassword_billing = null;
	public static String databaseHost = null;
	public static String databasePort = null;
	public static String databaseSID = null;
	public static String databaseAdminaccess = null;//06/08/2025 - USIC Project
	public static String EnvironmentType = null;
	public static String jbeamdatabaseusername = null;
	public static String jbeamdatabasepassword = null;
	public static String jbeamHost = null;
	public static String jbeamPort = null;
	public static String jbeamSID = null;
	public static String jbeamWsdlUrl = null;
	public static String jbeamContentType = null;
	public static String retrieveFailedObjectList = null;
	public static String jbeamusername = null;
	public static String jbeampassword = null;
	public static String jbeamVersion = null;
	public static String changebusinessdaterequestxmlpath = null;
	public static String changebusinessdatewsdl = null;
	public static String databaseType = null;
	public static String databaseName = null;
	public static String jbeamdatabaseName = null;
	public static String user = null;
	public static String password = null;
	public static String ChangeBusinessDay = null;
	public static String InterfaceOutBoundPath = null;
	public static String endComparison = null;
	public static String expectedValuesPath = null;
	public static String appendData = null;
	public static String recovery_scenario = null, recovery_scenario_billing = null, recovery_scenario_claims = null;
	public static String runbatchxmlpath = null;
	public static String runbatchxmlpath_day = null;
	public static String runbatchxmlpath_night = null;
	public static String verificationResultPath = null, appendVerificationResultPath = "", appendResultOutput = "",
			appendResultOutputGeneric = "", appendVerificationResultPathGeneric = "";;
	public static String dashboardResultPath, executionScope = null, openapiHost = null,
			validateResponseIgnoreKeys = null, securedJdbcConnection = null, validateResponseIgnoreKeysWebservice;
	public static String seleniumExecution, transactionInfo, policyUniqueNoPath, suiteDataPath, structureSheetFilePath;
	public static String defaultLogin, applicationName, verificationTableISTPath, verificationTemplatePath,
			actualValuesValuesPath, applicationPlatform, inputMasterFilePath;
	public static String expectedValuesValuesPath, invoiceComparision, invoiceXmlComparisionUtilityPath,
			executionStatusReportUtility, createVisualReportUtility, webserviceComparisonUtilityPath, authTokenURL,
			restfulServiceComparisonUtilityPath, openAPIServiceComparisonUtilityPath, webServiceComparisonResultPath;
	public static String emailId, resultOutput, suiteTesting, dashboardFilepath, webserviceComparison,
			restfulserviceComparison, openAPIserviceComparison;
	public static String authScope, authTokenUsername, authGrantType, authTokenPassword, authTokenClientId,
			authTokenClientSecret, authEnabled;
	public static String controllerFilePath, resetStartPointer, transactionInputFilePath, inputDataFilePath,
			applyStaticWait, copyServerRemotePath, resultFilePath, executionApproach, flatFileHostName,
			flatFileUserName, flatFilePassword, flatFilePort, flatFilePath = "", applyStaticWait_pas,
			applyStaticWait_claims, applyStaticWait_billing;
	public static String timeOut, userInteraction, cycleNumber, restComparisonResultPath, openAPIComparisonResultPath,
			headlessBrowser, Token_URL;
	public static String browserType, baseURL;
	public static String emailTransactionInfo = "false", emailOnFailure = "false";
	public static String basicAuthUsername, basicAuthPassword;
	public static DateFormat dtFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss:SSSS");
	private static String[] replaceSequences = { " & $B$2 & ", " &$B$2& ", "&$B$2& ", "&$B$19&", "$B$4&" };
	public static String defaultLocation1 = null;
	public static String defaultLocation2 = null;
	public static String expectedLocation = null;
	public static String actualLocation = null;
	public static String pdfBackupLocation = null;
	public static String pdfParamPath = null;
	public static String pdfParamResultPath = null;
	public static String expPdfFilePath = null;
	public static String actualPdfDownloadPath = null;
	public static String pdfCompResultPath = null;
	public static String serverFileLocation = null;
	public static String csvCompareReportFormat, csvCompareReportExpectedPath, csvCompareReportActualPath,
			csvCompareReportConfigPath, csvCompareReportResultPath;
	public static String externalscriptsfolder = null;
	public static String runtimeFileDownloadFolder = null;
	public static String expXMLFilePath = null;
	public static String actualXMLDownloadPath = null;
	public static String xmlCompResultPath = null;
	public static String xmldownloads = null;
	public static String xmlDefaultNodeLevelTagsToIgnore = null;
	public static String xmlDefaultAttributeLevelTagsToIgnore = null;
	public static String updateDynamicPath = null;
	public static String reconReportJarFileName = null;
	public static String reconReportJarFilePath = null;
	public static String reconReportGeneratedFilePath = null;
	public static String reconReportDateFormat = null;
	public static String reconReportArgs1 = null;
	public static String reconReportArgs2 = null;
	public static String reconReportArchivalPath = null;
	public static String invoiceProcessingRemotePath = null;
	public static String invoiceProcessingArchivePath = null;
	public static String invoiceProcessingLocalPath = null;
	public static String invoiceProcessingConfigExcelPath = null;
	public static String invoiceProcessingConfigExcelTab = null;
	public static String invoiceProcessingType = null;
	public static String invoiceProcessingReportExcelFile = null, reportSortByGroupID = null;
	public static String enableBackup, backupFilepath, backupFileList;
	public static String runMode = null;
	public static String digitCountUniqueNum = null;
	public static String verificationFailureHandling = "TRUE";
	public static String appHostAppPort = "";
	public static String openapidynamicOASURL = "";
	public static String openapidynamicFolder = "";
	// Below given parameters added for FTP unix server requirement -10/11/22
	public static String ftpserverhost = null;
	public static String ftpserverport = null;
	public static String ftpserveruserid = null;
	public static String ftpserverpassword = null;
	public static String copyServerRemotePathLinux = null;
	// WebOutlook Email
	public static String webOutlook_Email, executionStatusReportPath, outlookUsername, outlookApppassword,
			outlookemailFrom, outlookemailTo, outlookemailCC, outlookemailSubject;

	/*
	 * public static void initializeConfigMapNew(String configFilePathFileName)
	 * throws IOException, Exception { Sheet configSheet = null; Row rowActual =
	 * null; String parameterName = null; String value = null;
	 * 
	 * DataFormatter format = new DataFormatter(); String projectPath =
	 * System.getProperty("user.dir"); String configPath = projectPath + "\\" +
	 * configFilePathFileName;
	 * 
	 * configSheet = getExcelSheet(configPath, "Config"); int rowCount =
	 * configSheet.getLastRowNum() + 1;
	 * 
	 * for (int rowIndex = 1; rowIndex < rowCount; rowIndex++) { rowActual =
	 * configSheet.getRow(rowIndex); parameterName =
	 * format.formatCellValue(rowActual.getCell(0)); value =
	 * format.formatCellValue(rowActual.getCell(1));
	 * 
	 * value = value.replace("\"", ""); if (StringUtils.isNotBlank(parameterName) ||
	 * StringUtils.isNotBlank(value)) { configHashMap.put(parameterName, value);
	 * log.info(parameterName + " " + ":" + " " + value); } } //projectName =
	 * getConfigValue("PROJECT_NAME");
	 * 
	 * for (int rowIndex2 = 1; rowIndex2 < rowCount; rowIndex2++) {
	 * 
	 * rowActual = configSheet.getRow(rowIndex2); parameterName =
	 * format.formatCellValue(rowActual.getCell(0)); value =
	 * format.formatCellValue(rowActual.getCell(1)); value = value.replace("\"",
	 * "");
	 * 
	 * for (String replaceSeq : replaceSequences) { if (value.contains(replaceSeq))
	 * { value = value.replace(replaceSeq, projectName); break; } }
	 * 
	 * if (StringUtils.isNotBlank(parameterName) || StringUtils.isNotBlank(value)) {
	 * configHashMap.put(parameterName, value);
	 * 
	 * } }
	 * 
	 * // Update dynamic path [UPDATE_DYNAMIC_PATH] updateDynamicPath =
	 * getConfigValue("UPDATE_DYNAMIC_PATH"); if (updateDynamicPath != null &&
	 * updateDynamicPath.equalsIgnoreCase("True")) { String currentDir =
	 * System.getProperty("user.dir"); for (String configProp :
	 * configHashMap.keySet()) { String configPropValue =
	 * getConfigValue(configProp); if (configPropValue != null &&
	 * configPropValue.contains("%DYNAMIC_PATH%")) configHashMap.replace(configProp,
	 * configPropValue.replace("%DYNAMIC_PATH%", currentDir)); } } // Update dynamic
	 * path END.
	 * 
	 * loadConfigData();
	 * 
	 * }
	 */

	public static void initializeConfigMap(String configFilePathFileName) throws IOException, Exception {
		Sheet configSheet = null;
		Row rowActual = null;
		String parameterName = null;
		String value = null;

		DataFormatter format = new DataFormatter();
		String projectPath = System.getProperty("user.dir");
		String configPath = projectPath + "\\" + configFilePathFileName;

		configSheet = getExcelSheet(configPath, "Config");
		FormulaEvaluator evaluator = configSheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		int rowCount = configSheet.getLastRowNum() + 1;

		
		for (int rowIndex = 1; rowIndex < rowCount; rowIndex++) {
			rowActual = configSheet.getRow(rowIndex);
			parameterName = format.formatCellValue(rowActual.getCell(0), evaluator);
			 value = format.formatCellValue(rowActual.getCell(1), evaluator);
	         value = value.replace("\"", "");
			
	         if (StringUtils.isNotBlank(parameterName) || StringUtils.isNotBlank(value)) {
	                configHashMap.put(parameterName, value);
	                log.info(parameterName + " " + ":" + " " + value); // Replace log.info with System.out.println for simplicity
	            }
		}

		for (int rowIndex2 = 1; rowIndex2 < rowCount; rowIndex2++) {
			rowActual = configSheet.getRow(rowIndex2);
			parameterName = format.formatCellValue(rowActual.getCell(0), evaluator);
			value = format.formatCellValue(rowActual.getCell(1), evaluator);
            value = value.replace("\"", "");

			for (String replaceSeq : replaceSequences) {
				if (value.contains(replaceSeq)) {
					value = value.replace(replaceSeq, projectName);
					break;
				}
			}

			if (StringUtils.isNotBlank(parameterName) || StringUtils.isNotBlank(value)) {
				configHashMap.put(parameterName, value);
			}
		}

		updateDynamicPath = getConfigValue("UPDATE_DYNAMIC_PATH");
		if (updateDynamicPath != null && updateDynamicPath.equalsIgnoreCase("True")) {
			String currentDir = System.getProperty("user.dir");
			for (String configProp : configHashMap.keySet()) {
				String configPropValue = getConfigValue(configProp);
				if (configPropValue != null && configPropValue.contains("%DYNAMIC_PATH%"))
					configHashMap.replace(configProp, configPropValue.replace("%DYNAMIC_PATH%", currentDir));
			}
		}

		loadConfigData();
	}

	/** Loads the Config sheet into HashMap **/
	public static void loadConfigData() throws Exception {

		try {
			projectName = getConfigValue("PROJECT_NAME");
			seleniumGrid = getConfigValue("SELENIUM_GRID");
			incognitoMode = getConfigValue("INCOGNITO_MODE");
			environment = getConfigValue("ENVIRONMENT");//For Claims Cloud env Refresh issue -23/10/2024

			endComparison = getConfigValue("END_COMPARISON");
			appendData = getConfigValue("APPEND_DATA_RUNTIME");
			recovery_scenario = getConfigValue("RECOVERY_SCENARIO");
			recovery_scenario_billing = getConfigValue("RECOVERY_SCENARIO_BILLING");
			recovery_scenario_claims = getConfigValue("RECOVERY_SCENARIO_CLAIMS");

			errorlog = getConfigValue("WEBSERVICE_ERRORLOG_FILEPATH");
			// added for LNA product for token url
			Token_URL = getConfigValue("Token_URL");
			headlessBrowser = getConfigValue("HEADLESS_BROWSER");
			applicationdatabaseusername = getConfigValue("APPLICATIONDATABASEUSERNAME");
			applicationdatabasepassword = getConfigValue("APPLICATIONDATABASEPASSWORD");
			databaseHost = getConfigValue("DATABASE_HOST_IP_ADDRESS");
			databasePort = getConfigValue("DATABASE_PORT");
			databaseSID = getConfigValue("DATABASE_SID");
			databaseName = getConfigValue("DATABASE_NAME");
			databaseAdminaccess = getConfigValue("DATABASE_ADMIN_ACCESS");//06/08/2025 - USIC Project
			jbeamdatabaseName = getConfigValue("JBEAM_DATABASE_NAME");// For LNA Billing Cloud Env - 29\03\2023
			databaseType = getConfigValue("DATABASE_TYPE");

			// Suite specific DB credentials
			applicationdatabaseusername_pas = getConfigValue("APPLICATIONDATABASEUSERNAME_PAS");
			applicationdatabasepassword_pas = getConfigValue("APPLICATIONDATABASEPASSWORD_PAS");
			applicationdatabaseusername_claims = getConfigValue("APPLICATIONDATABASEUSERNAME_CLAIMS");
			applicationdatabasepassword_claims = getConfigValue("APPLICATIONDATABASEPASSWORD_CLAIMS");
			applicationdatabaseusername_billing = getConfigValue("APPLICATIONDATABASEUSERNAME_BILLING");
			applicationdatabasepassword_billing = getConfigValue("APPLICATIONDATABASEPASSWORD_BILLING");

			jbeamdatabaseusername = getConfigValue("JBEAMDATABASEUSERNAME");
			jbeamdatabasepassword = getConfigValue("JBEAMDATABASEPASSWORD");
			jbeamHost = getConfigValue("JBEAM_HOST_IP_ADDRESS");
			jbeamPort = getConfigValue("JBEAM_PORT");
			jbeamSID = getConfigValue("JBEAM_SID");
			jbeamusername = getConfigValue("JBEAM_USERID");
			jbeampassword = getConfigValue("JBEAM_PASSWORD");
			jbeamVersion = getConfigValue("JBEAM_VERSION");
			jbeamWsdlUrl = getConfigValue("JBEAM:WSDL_URL");
			jbeamContentType = getConfigValue("JBEAM:CONTENTTYPE");
			user = getConfigValue("APPLICATION_USER_ID");
			password = getConfigValue("APPLICATION_PASSWORD");

			ChangeBusinessDay = getConfigValue("CHANGE_BUSINESSDAY");
			InterfaceOutBoundPath = getConfigValue("INTERFACE_OUTBOUND_PATH");// Added
			expectedValuesPath = expectedValuesValuesPath;
			runbatchxmlpath = getConfigValue("JBEAMDAYBATCHXMLPATH");
			runbatchxmlpath_day = getConfigValue("JBEAM_DAY_BATCH_XMLPATH");
			runbatchxmlpath_night = getConfigValue("JBEAM_NIGHT_BATCH_XMLPATH");
			retrieveFailedObjectList = getConfigValue("RETRIEVE_FAILED_OBJECT_LIST");
			runbatchinputsheetpath = getConfigValue("JBEAMDAYBATCHINPUT");
			changebusinessdaterequestxmlpath = getConfigValue("CHANGEBUSINESSDATE_REQUEST_XML");
			changebusinessdatewsdl = getConfigValue("CHANGEBUSINESSDATE_WSDL");
			verificationResultPath = getConfigValue("VERIFICATIONRESULTSPATH");
			appendVerificationResultPath = getConfigValue("PAS_APPEND_VERIFICATIONRESULTSPATH");
			appendVerificationResultPathGeneric = getConfigValue("APPEND_VERIFICATIONRESULTSPATH");
			dashboardResultPath = getConfigValue("DASHBOARDRESULTSPATH");
			timeOut = getConfigValue("TIMEOUT");
			controllerFilePath = getConfigValue("CONTROLLER_FILEPATH");
			resetStartPointer = getConfigValue("RESET_START_POINTER"); // Pradeep-added
			transactionInputFilePath = getConfigValue("TRANSACTION_INPUT_FILEPATH");
			productTeam = getConfigValue("PRODUCT_TEAM");
			EnvironmentType = getConfigValue("ENVIRONMENT_TYPE");
			resultFilePath = getConfigValue("RESULT_FILEPATH");
			executionApproach = getConfigValue("EXECUTION_APPROACH");
			userInteraction = getConfigValue("USERINTERACTION");
			cycleNumber = getConfigValue("CYCLENUMBER");
			flatFileHostName = getConfigValue("FLATFILE_HOSTNAME");
			flatFileUserName = getConfigValue("FLATFILE_USERNAME");
			flatFilePassword = getConfigValue("FLATFILE_PASSWORD");
			flatFilePort = getConfigValue("FLATFILE_PORT");
			flatFilePath = getConfigValue("FLATFILE_PATH");
			inputDataFilePath = getConfigValue("INPUT_DATA_FILEPATH");
			applyStaticWait = getConfigValue("APPLY_STATIC_WAIT");
			applyStaticWait_pas = getConfigValue("APPLY_STATIC_WAIT_PAS");
			applyStaticWait_claims = getConfigValue("APPLY_STATIC_WAIT_CLAIMS");
			applyStaticWait_billing = getConfigValue("APPLY_STATIC_WAIT_BILLING");

			// Below given parameters added for FTP unix server requirement -10/11/22
			copyServerRemotePath = getConfigValue("COPY_SERVER_REMOTE_PATH");
			copyServerRemotePathLinux = getConfigValue("COPY_SERVER_REMOTE_PATH_LINUX");
			ftpserverhost = getConfigValue("FTP_SERVER_HOST");
			ftpserverport = getConfigValue("FTP_SERVER_PORT");
			ftpserveruserid = getConfigValue("FTP_SERVER_USER_ID");
			ftpserverpassword = getConfigValue("FTP_SERVER_PASSWORD");

			webserviceComparisonUtilityPath = getConfigValue("WEBSERVICE_COMPARISION_UTILITYPATH");
			webServiceComparisonResultPath = getConfigValue("WEBSERVICE_COMPARISON_RESULT_PATH");
			restComparisonResultPath = getConfigValue("REST_COMPARISON_RESULT_PATH");
			restfulServiceComparisonUtilityPath = getConfigValue("RESTFULSERVICE_COMPARISON_UTILITYPATH");
			restfulserviceComparison = getConfigValue("RESTFULSERVICE_COMPARISON");

			openAPIServiceComparisonUtilityPath = getConfigValue("OPENAPISERVICE_COMPARISON_UTILITYPATH");
			openAPIserviceComparison = getConfigValue("OPENAPISERVICE_COMPARISON");
			openapiHost = getConfigValue("OPENAPI_HOST");
			validateResponseIgnoreKeys = getConfigValue("VALIDATE_RESPONSE_IGNORE_KEYS");
			validateResponseIgnoreKeysWebservice = getConfigValue("VALIDATE_RESPONSE_IGNORE_KEYS_WEBSERVICE");// For New

			openapidynamicOASURL = getConfigValue("OPENAPI_DYNAMIC_OASURL");// OpenAPI change 20-10-2022
			openapidynamicFolder = getConfigValue("OPENAPI_DYNAMIC_FOLDER");// OpenAPI change 20-10-2022

			basicAuthUsername = getConfigValue("BASIC_AUTH_USERNAME");
			basicAuthPassword = getConfigValue("BASIC_AUTH_PASSWORD");

			authGrantType = Config.getConfigValue("AUTH_GRANT_TYPE");
			authScope = Config.getConfigValue("AUTH_SCOPE");
			authTokenUsername = Config.getConfigValue("AUTH_TOKEN_USERNAME");
			authTokenPassword = Config.getConfigValue("AUTH_TOKEN_PASSWORD");
			authTokenClientId = Config.getConfigValue("AUTH_TOKEN_CLIENT_ID");
			authTokenClientSecret = Config.getConfigValue("AUTH_TOKEN_CLIENT_SECRET");
			authEnabled = Config.getConfigValue("AUTH_ENABLED");
			authTokenURL = Config.getConfigValue("AUTH_TOKEN_URL");
			emailId = getConfigValue("EMAIL_ID");
			emailOnFailure = getConfigValue("EMAIL_ON_FAILURE");
			resultOutput = getConfigValue("RESULTOUTPUT");
			appendResultOutput = getConfigValue("PAS_APPEND_RESULTOUTPUT");
			appendResultOutputGeneric = getConfigValue("APPEND_RESULTOUTPUT");
			suiteTesting = getConfigValue("SUITE_TESTING");
			dashboardFilepath = getConfigValue("DASHBOARD_FILEPATH");
			webserviceComparison = getConfigValue("WEBSERVICE_COMPARISION");
			executionStatusReportUtility = getConfigValue("EXECUTION_STATUS_REPORT_UTILITY");
			createVisualReportUtility = getConfigValue("VISUAL_REPORT_UTILITY");
			webOutlook_Email = getConfigValue("WEBOUTLOOK_EMAIL");
			executionStatusReportPath = getConfigValue("EXECUTION_STATUS_REPORT_PATH");
			outlookUsername = getConfigValue("OUTLOOK_USERNAME");
			outlookApppassword = getConfigValue("OUTLOOK_APP_PASSWORD");
			outlookemailFrom = getConfigValue("OUTLOOK_EMAIL_FROM");
			outlookemailTo = getConfigValue("OUTLOOK_EMAIL_TO");
			outlookemailCC = getConfigValue("OUTLOOK_EMAIL_CC");
			outlookemailSubject = getConfigValue("OUTLOOK_EMAIL_SUBJECT");

			if (getConfigValue("VISUAL_REPORT_UTILITY") == null || getConfigValue("VISUAL_REPORT_UTILITY").isEmpty()
					|| getConfigValue("VISUAL_REPORT_UTILITY").equalsIgnoreCase("TRUE"))
				createVisualReportUtility = "TRUE";
			else
				createVisualReportUtility = "FALSE";

			defaultLogin = getConfigValue("DEFAULTLOGIN");
			applicationName = getConfigValue("APPLICATION_NAME");
			verificationTableISTPath = getConfigValue("VERIFICATIONTABLELISTPATH");
			verificationTemplatePath = getConfigValue("VERIFCATIONTEMPLATEPATH");
			actualValuesValuesPath = getConfigValue("ACTUALVALUESPATH");
			expectedValuesValuesPath = getConfigValue("EXPECTEDVALUESPATH");
			seleniumExecution = getConfigValue("SELENIUMEXECUTION");
			transactionInfo = getConfigValue("TRANSACTION_INFO");
			emailTransactionInfo = getConfigValue("EMAIL_TRANSACTION_INFO");
			policyUniqueNoPath = getConfigValue("POLICY_UNIQUE_NUMBER_PATH");
			suiteDataPath = getConfigValue("SUITE_DATA_PATH");
			structureSheetFilePath = getConfigValue("STRUCTURE_SHEET_FILEPATH");
			inputMasterFilePath = getConfigValue("INPUT_MASTER_FILEPATH");
			browserType = getConfigValue("BROWSERTYPE");
			baseURL = getConfigValue("BASEURL");
			// pdf utility properties added below
			defaultLocation1 = getConfigValue("DEFAULT_LOCATION_AGENCY");
			defaultLocation2 = getConfigValue("DEFAULT_LOCATION_SINGLE");
			expectedLocation = getConfigValue("EXPECTED_LOCATION");
			actualLocation = getConfigValue("ACTUAL_LOCATION");
			pdfBackupLocation = getConfigValue("PDFBACKUP_LOCATION");
			pdfParamPath = getConfigValue("PDF_CONFIG_PATH");
			pdfParamResultPath = getConfigValue("PDF_RESULT_PATH");

			// PAS specific changes for pdf comparision
			externalscriptsfolder = getConfigValue("EXT_SCRIPT_EXE_FOLDER");
			expPdfFilePath = getConfigValue("PDF_COMP_ExpPDF");
			runtimeFileDownloadFolder = getConfigValue("RUN_FILE_LOCAL_FOLDER");
			if (getConfigValue("PDF_COMP_ActPDF_folder") != null)
				actualPdfDownloadPath = runtimeFileDownloadFolder + "\\" + getConfigValue("PDF_COMP_ActPDF_folder");
			else
				actualPdfDownloadPath = runtimeFileDownloadFolder;

			pdfCompResultPath = getConfigValue("PDF_COMP_PDF_Results");
			serverFileLocation = getConfigValue("SERVER_FILE_LOCATION");
			expXMLFilePath = getConfigValue("XML_COMP_ExpXML");

			if (getConfigValue("XML_COMP_ActXML_folder") != null)
				actualXMLDownloadPath = runtimeFileDownloadFolder + "\\" + getConfigValue("XML_COMP_ActXML_folder");
			else
				actualXMLDownloadPath = runtimeFileDownloadFolder;
			if (getConfigValue("XML_Download_folder") != null)
				xmldownloads = runtimeFileDownloadFolder + "\\" + getConfigValue("XML_Download_folder");

			else
				xmldownloads = runtimeFileDownloadFolder;

			// downloaded xml's
			xmlCompResultPath = getConfigValue("XML_COMP_XML_Results");
			xmlDefaultAttributeLevelTagsToIgnore = getConfigValue("XML_Default_AttributeLevelTagsToIgnore");
			xmlDefaultNodeLevelTagsToIgnore = getConfigValue("XML_Default_NodeLevelTagsToIgnore");
			csvCompareReportFormat = getConfigValue("REPORTFORMAT");
			csvCompareReportExpectedPath = getConfigValue("REPORTEXPECTEDPATH");
			csvCompareReportActualPath = getConfigValue("REPORTACTUALPATH");
			csvCompareReportConfigPath = getConfigValue("REPORTCONFIGPATH");
			csvCompareReportResultPath = getConfigValue("REPORTRESULTPATH");
			reconReportJarFileName = getConfigValue("RECON_REPORT_JAR_FILE_NAME");
			reconReportJarFilePath = getConfigValue("RECON_REPORT_JAR_FILE_PATH");
			reconReportGeneratedFilePath = getConfigValue("RECON_REPORT_GENERATED_FILE_PATH");
			reconReportDateFormat = getConfigValue("RECON_REPORT_DATE_FORMAT");
			reconReportArgs1 = getConfigValue("RECON_REPORT_ARGS_1");
			reconReportArgs2 = getConfigValue("RECON_REPORT_ARGS_2");
			reconReportArchivalPath = getConfigValue("RECON_REPORT_ARCHIVAL_PATH");
			invoiceProcessingRemotePath = getConfigValue("INVOICE_PROCESSING_REMOTE_PATH");
			invoiceProcessingArchivePath = getConfigValue("INVOICE_PROCESSING_ARCHIVE_PATH");
			invoiceProcessingLocalPath = getConfigValue("INVOICE_PROCESSING_LOCAL_PATH");
			invoiceProcessingConfigExcelPath = getConfigValue("INVOICE_PROCESSING_CONFIG_EXCEL_PATH");
			invoiceProcessingConfigExcelTab = getConfigValue("INVOICE_PROCESSING_CONFIG_EXCEL_TAB");
			invoiceProcessingType = getConfigValue("INVOICE_PROCESSING_TYPE");
			invoiceProcessingReportExcelFile = getConfigValue("INVOICE_PROCESSING_REPORT_EXCEL_FILE");
			invoiceComparision = getConfigValue("INVOICE_COMPARISION");
			invoiceXmlComparisionUtilityPath = getConfigValue("INVOICE_XML_COMPARISION_UTILITY_PATH");

			reportSortByGroupID = getConfigValue("REPORT_SORT_BY_GROUP_ID");

			// Added by Sheetal PAS
			executionScope = getConfigValue("EXECUTION_SCOPE");

			// for secured JDBC Connection
			securedJdbcConnection = getConfigValue("SECURED_JDBC_CONNECTION");

			// update dynamic path in Config.xls
			updateDynamicPath = getConfigValue("UPDATE_DYNAMIC_PATH");

			// iTAF Backup Configuration
			enableBackup = getConfigValue("ENABLE_BACKUP");
			backupFilepath = getConfigValue("BACKUP_FILEPATH");
			backupFileList = getConfigValue("BACKUP_FILE_LIST");

			// Application platform for processing icon handling in WebHelperPAS
			applicationPlatform = getConfigValue("APPLICATION_PLATFORM");

			// Suite implementation
			runMode = getConfigValue("RUN_MODE");

			// Claim verification handling
			verificationFailureHandling = getConfigValue("VERIFICATION_FAILURE_HANDLING");

			// Suite Billing Webservice
			appHostAppPort = getConfigValue("APPHOST\\:APPPORT");

			digitCountUniqueNum = getConfigValue("DIGIT_COUNT_UNIQUE_NUM");

		} catch (Exception e) {
			log.error(e.getMessage(), e);
			throw e;
		}
	}

	public static int getConfgiMapSize() {
		return configHashMap.size();
	}

	public static String getConfigValue(String string) {
		Object value = configHashMap.get(string);
		if (value != null) {
			return value.toString();
		}
		return null;
	}

	public static Sheet getExcelSheet(String FilePath, String SheetName) throws Exception {
		Sheet workSheet = null;
		Workbook workBook = null;
		InputStream myXls = null;
		try {
			log.info("FilePath is:" + FilePath);
			log.info("SheetName is:" + SheetName);
			myXls = new FileInputStream(FilePath);

			String configSheetExtension = FilenameUtils.getExtension(FilePath);
			try {
				if (configSheetExtension.equalsIgnoreCase("xls")) {
					workBook = new HSSFWorkbook(myXls);
				} else if (configSheetExtension.equalsIgnoreCase("xlsx")
						|| configSheetExtension.equalsIgnoreCase("xlsm")) {
					workBook = new XSSFWorkbook(myXls);
				} else {
					throw new IllegalArgumentException(
							"Invalid file type. File type should be one of 'xls', 'xlsx' or 'xlsm'");
				}
			} finally {
				if (myXls != null) {
					myXls.close();
				}
			}
			workSheet = workBook.getSheet(SheetName);
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			throw e;
		} finally {
			try {
				if (workBook != null) {
					workBook.close();
				}
			} catch (IOException e) {
				// Handle IOException
			}
		}
		return workSheet;
	}

}
