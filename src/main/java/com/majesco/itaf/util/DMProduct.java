package com.majesco.itaf.util;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Vector;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import com.jcraft.jsch.Channel;
import com.jcraft.jsch.ChannelSftp;
import com.jcraft.jsch.ChannelSftp.LsEntry;
import com.jcraft.jsch.JSch;
import com.jcraft.jsch.JSchException;
import com.jcraft.jsch.Session;
import com.jcraft.jsch.SftpException;
import com.majesco.itaf.main.Automation;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.ITAFWebDriver;
import com.majesco.itaf.main.MainController;
import com.majesco.itaf.main.ObjectFactory;
import com.majesco.itaf.main.WebHelper;
import com.majesco.itaf.main.WebHelperDM;
import com.majesco.itaf.main.WebHelperUtil;

public class DMProduct {
	private final static Logger log = LogManager.getLogger(DMProduct.class.getName());
	static MainController controller = ObjectFactory.getMainController();
	static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	public static String sqlQuery = "";
	public static String readFromColName = "";
	public static String toBeFetchedDBColName = "";
	public static String expectedDBStatus = "";
	public static String writeToColName = "";
	public static String Local_Path = "";
	public static String Template_Path = "";
	public static String Remote_Path = "";
	public static String Remote_Path_Outbound = "";
	public static String EffectiveDate = "";
	public static String EffectiveDate1 = "";
	public static String EffectiveDate2 = "";
	public static String EffectiveDate3 = "";
	public static String EffectiveDate4 = "";
	public static String EffectiveDate5 = "";
	public static String EffectiveDate6 = "";
	public static String EffectiveDate7 = "";
	public static String EffectiveDate8 = "";
	public static String EffectiveDate9 = "";
	public static String EffectiveDate_10 = "";
	public static String EffectiveDate_11 = "";
	public static String RiskCommencementDate = "";
	public static String RiskCommencementDate1 = "";
	public static String RiskCommencementDate2 = "";
	public static String PolicyEffectiveToDate = "";
	public static String ProposalApplDate = "";
	public static String ProposalSubmDate = "";
	public static String PolicyNumber = "";
	public static String PolicyNumber1 = "";
	public static String PolicyNumber2 = "";
	public static String PolicyNumber3 = "";
	public static String PolicyNumber4 = "";
	public static String PolicyNumber5 = "";
	public static String GroupFileLoadNumber = "";
	public static String PolicyFileLoadNumber = ""; // 01-05-2019
	public static String PremiumFileLoadNumber = ""; // 01-05-2019
	public static String EntityCode1 = "";
	public static String EntityCode2 = "";
	public static String EntityCode3 = "";
	public static String EntityCode4 = "";
	public static String EntityCode5 = "";
	public static String PayoutHeaderID = "";// 01-01-2020
	public static String TemplateFileName = "";
	public static String FileNameToBeCreated = "";
	public static String FileName = "";// 06-16-2022
	public static String CreateFile = "";
	public static String FromLocalToServer = "";
	public static String FromServerToLocal = "";
	public static String FileNameCopiedToLocal = "";
	public static String SSED1 = "";
	public static String SSN1 = "";
	public static String CSCI1 = "";
	public static String MemberName1 = "";
	public static String SSED2 = "";
	public static String SSN2 = "";
	public static String CSCI2 = "";
	public static String MemberName2 = "";
	public static String ApxGrpAddr1 = "";
	public static String ApxGrpEmailId = "";
	public static String ApxGrpNo = "";
	public static String SubGrpNo = "";
	public static String ApxGrpName = "";
	public static String ApxSSPK = "";
	public static String SSRN = "";
	public static String RFPN = "";
	public static String ApxGrpNo0001 = "";
	public static String ApxGrpNo0002 = "";
	public static String SubGrpAddr1 = "";
	public static String SubGrpAddr2 = "";
	public static String SubGrpAddr3 = "";
	public static String SubGrpSSEC1 = "";
	public static String SubGrpSSEC2 = "";
	public static String SubGrpSSEC3 = "";
	public static String SASubGrpNo1 = "";
	public static String SASubGrpNo2 = "";
	public static String SASubGrpNo3 = "";
	public static String SASubGrpNo4 = "";
	public static String SubGrpName1 = "";
	public static String SubGrpName2 = "";
	public static String SubGrpName3 = "";
	public static String SubGrpSSPK1 = "";
	public static String SubGrpSSPK2 = "";
	public static String SubGrpSSPK3 = "";
	public static String AGENCY_ID = "";
	public static String NPN1 = "";
	public static String NPN2 = "";
	public static String NPN3 = "";
	public static String NPN4 = "";
	public static String NPN5 = "";
	public static String FEIN = "";
	public static String SSN3 = "";
	public static String SSN4 = "";
	public static String SSN5 = "";
	public static String FN1 = "";
	public static String FN2 = "";
	public static String FN3 = "";
	public static String FN4 = "";
	public static String FN5 = "";
	public static String MN1 = "";
	public static String MN2 = "";
	public static String MN3 = "";
	public static String MN4 = "";
	public static String MN5 = "";
	public static String LN1 = "";
	public static String LN2 = "";
	public static String LN3 = "";
	public static String LN4 = "";
	public static String LN5 = "";
	public static String EP_CODE1 = "";
	public static String EP_CODE2 = "";
	public static String EP_CODE3 = "";
	public static String EP_CODE4 = "";
	public static String EP_CODE5 = "";
	public static String NSCC_NO1 = "";
	public static String NSCC_NO2 = "";
	public static String NSCC_NO3 = "";
	public static String NSCC_NO4 = "";
	public static String NSCC_NO5 = "";
	public static String ACCOUNT_NO1 = "";
	public static String ACCOUNT_NO2 = "";
	public static String ACCOUNT_NO3 = "";
	public static String ACCOUNT_NO4 = "";
	public static String ACCOUNT_NO5 = "";
	public static String POL_NO1 = "";
	public static String POL_NO2 = "";
	public static String POL_NO3 = "";
	public static String POL_NO4 = "";
	public static String POL_NO5 = "";
	public static String EAOEF_DATE1 = "";
	public static String EAOEF_DATE2 = "";
	public static String EAOEF_DATE3 = "";
	public static String EAOEF_DATE4 = "";
	public static String EAOEF_DATE5 = "";
	public static String EAOEF_DATE = "";
	public static String EF_DATE1 = "";
	public static String EF_DATE2 = "";
	public static String EF_DATE3 = "";
	public static String EF_DATE4 = "";
	public static String EF_DATE5 = "";
	public static String EF_DATE = "";
	public static String ET_DATE1 = "";
	public static String ET_DATE2 = "";
	public static String ET_DATE3 = "";
	public static String ET_DATE4 = "";
	public static String ET_DATE5 = "";
	public static String LICENSE_NO1 = "";
	public static String LICENSE_NO2 = "";
	public static String LICENSE_NO3 = "";
	public static String LICENSE_NO4 = "";
	public static String LICENSE_NO5 = "";
	public static String SC_CODE1 = "";
	public static String SC_CODE2 = "";
	public static String SC_CODE3 = "";
	public static String SC_CODE4 = "";
	public static String SC_CODE5 = "";
	public static String COURSE_COMPLETION_DATE = "";
	public static String COURSE_CERTIFICATION_DATE = "";
	public static String COURSE_SUBMISSION_DATE = "";
	public static String HOUR = "";
	public static String ENROLL_DATE = "";
	public static String STATUS_DATE = "";
	public static String STATUSDATE1 = "";
	public static File copySource;
	public static File copyDestination;
	public static String source;
	public static String destination;

	public static String ReadFromExcelUsingColumnName(String controlValue, String ColumnName) throws IOException {
		Sheet uniqueNumberSheet = null;
		String uniqueTestcaseID = "";
		HashMap<String, Integer> uniqueValuesHashMap = null;
		String uniqueNumber = null;
		String SearchTestcaseID = null;
		try {
			uniqueNumberSheet = WebHelperUtil.getSheet(Config.transactionInfo, "DataSheet");
			uniqueValuesHashMap = WebHelperUtil.getValueFromHashMap(uniqueNumberSheet);
			int rowCount = uniqueNumberSheet.getPhysicalNumberOfRows();

			if (controlValue.equals("")) {
				SearchTestcaseID = controller.controllerTestCaseID.toString();
			} else {
				SearchTestcaseID = controlValue;
			}

			for (int rIndex = 1; rIndex < rowCount; rIndex++) {
				uniqueTestcaseID = WebHelperUtil.getCellData("TestCaseID", uniqueNumberSheet, rIndex,
						uniqueValuesHashMap);

				if (SearchTestcaseID.equals(uniqueTestcaseID)) {
					return uniqueNumber = WebHelperUtil.getCellData(ColumnName, uniqueNumberSheet, rIndex,
							uniqueValuesHashMap);
				}
			}

		} catch (Exception e) {
			log.error(e.getMessage(), e);
			controller.pauseFun(e.getMessage() + " from ReadFromExcel Function");
		}
		return uniqueNumber;
	}

	// Added this function to write SQL queries to Unique Number
	// sheet for reference
	public static void writeDataToUniqueNumberSheet(String ctrlValue, String columnName, String srchTCID)
			throws Exception {
		Workbook uniqueWB = null;
		String SearchTestcaseID = null;
		try {
			FileInputStream in = new FileInputStream(Config.transactionInfo.toString());
			uniqueWB = WebHelperUtil.createWorkbook(in, Config.transactionInfo);

			Sheet uniqueNumberSheet = uniqueWB.getSheet("DataSheet");
			HashMap<String, Integer> uniqueValuesHashMap = WebHelperUtil.getValueFromHashMap(uniqueNumberSheet);
			Row uniqueRow = null;
			int rowNum = uniqueNumberSheet.getPhysicalNumberOfRows();
			log.info("%%%%%%%%*********" + rowNum);

			if (srchTCID.equals("")) {
				SearchTestcaseID = controller.controllerTestCaseID.toString();
			} else {
				SearchTestcaseID = srchTCID;
			}
			for (int rIndex = 0; rIndex < rowNum; rIndex++) {
				uniqueRow = uniqueNumberSheet.getRow(rIndex);
				String uniqueTestcaseID = WebHelperUtil.getCellData("TestCaseID", uniqueNumberSheet, rIndex,
						uniqueValuesHashMap);

				if (SearchTestcaseID.equals(uniqueTestcaseID)) {
					uniqueRow = uniqueNumberSheet.getRow(rIndex);
					break;
				} else if (rIndex == rowNum - 1) {
					uniqueRow = uniqueNumberSheet.createRow(rowNum);
				}
			}

			Cell uniqueTestCaseID = uniqueRow.createCell(uniqueValuesHashMap.get("TestCaseID").intValue());
			Cell uniqueCell = uniqueRow.createCell(uniqueValuesHashMap.get(columnName).intValue());
			uniqueTestCaseID.setCellValue(SearchTestcaseID);
			uniqueCell.setCellValue(ctrlValue);
			in.close();

			FileOutputStream out = new FileOutputStream(Config.transactionInfo);
			uniqueWB.write(out);

		} catch (FileNotFoundException e) {
			log.error(e.getMessage(), e);
			throw e;
		} catch (Exception e) {
			log.error(e.getMessage(), e);
			throw e;
		} finally {
			try {
				uniqueWB.close();
			} catch (Exception ex) {
			}
		}
	}

	public static Boolean writeRequestIDToExcel(String ctrlValue, WebElement webElement, String controlId,
			String controlType, String controlName, String rowNo, String colNo) throws Exception {
		Workbook uniqueWB = null;
		try {
			FileInputStream in = new FileInputStream(Config.transactionInfo.toString());
			uniqueWB = WebHelperUtil.createWorkbook(in, Config.transactionInfo);
			Sheet uniqueNumberSheet = uniqueWB.getSheet("DataSheet");
			HashMap<String, Integer> uniqueValuesHashMap = WebHelperUtil.getValueFromHashMap(uniqueNumberSheet);
			Row uniqueRow = null;
			int rowNum = uniqueNumberSheet.getPhysicalNumberOfRows();
			log.info("%%%%%%%%*********" + rowNum);

			for (int rIndex = 0; rIndex < rowNum; rIndex++) {
				uniqueRow = uniqueNumberSheet.getRow(rIndex);
				String uniqueTestcaseID = WebHelperUtil.getCellData("TestCaseID", uniqueNumberSheet, rIndex,
						uniqueValuesHashMap);

				if (controller.controllerTestCaseID.toString().equals(uniqueTestcaseID)) {
					uniqueRow = uniqueNumberSheet.getRow(rIndex);
					break;
				} else if (rIndex == rowNum - 1) {
					uniqueRow = uniqueNumberSheet.createRow(rowNum);
				}
			}

			ctrlValue = webElement.getText();
			String arrString[] = ctrlValue.split("\\s+");

			Cell uniqueTestCaseID = uniqueRow.createCell(uniqueValuesHashMap.get("TestCaseID").intValue());
			Cell uniqueCell = uniqueRow.createCell(uniqueValuesHashMap.get(WebHelper.columnName).intValue());
			uniqueTestCaseID.setCellValue(controller.controllerTestCaseID.toString());

			for (int i = 0; i < arrString.length; i++) {
				if (arrString[i].toString().equalsIgnoreCase("request")
						|| arrString[i].toString().equalsIgnoreCase("request#")
						|| arrString[i].toString().equalsIgnoreCase("Id")
						|| (arrString[i].toString().equalsIgnoreCase("termination")
								&& !arrString[i + 1].toString().equalsIgnoreCase("request"))) {
					if (arrString[i + 1].toString().contains(".")) {
						uniqueCell.setCellValue(arrString[i + 1].toString().replace(".", ""));
					} else {
						uniqueCell.setCellValue(arrString[i + 1].toString());
					}
					break;
				}
			}

			in.close();
			FileOutputStream out = new FileOutputStream(Config.transactionInfo);
			uniqueWB.write(out);
		} catch (FileNotFoundException e) {
			throw e;
		} catch (Exception e) {
			throw e;
		}
		return true;
	}

	public static String convertDate(String strValue) {
		String strTemp = null;
		DateFormat inputFormat = new SimpleDateFormat("MM/dd/yyyy");
		DateFormat outputFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.S");
		try {
			Date date = inputFormat.parse(strValue);
			strTemp = outputFormat.format(date);
		} catch (ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return strTemp;

	}

	public static String getDynamicSQLQuery(String tableID, String ctrlName) throws IOException {

		String uniqueNumber = "";
		String uniqueNumber1 = "";
		String uniqueNumber2 = "";

		if (ctrlName.contains("|")) {
			String[] arrTempCntrlVal = ctrlName.split("\\|");
			uniqueNumber = ReadFromExcelUsingColumnName("", arrTempCntrlVal[0].toString());
			uniqueNumber1 = ReadFromExcelUsingColumnName("", arrTempCntrlVal[1].toString());
			uniqueNumber2 = ReadFromExcelUsingColumnName("", arrTempCntrlVal[2].toString());
		} else {
			uniqueNumber = ReadFromExcelUsingColumnName("", ctrlName);
		}

		if (uniqueNumber.contains(",")) {
			String[] arrTempVal = uniqueNumber.split(",");
			if (arrTempVal.length <= 2) { // 03-01-2019
				if (tableID.contains("$value1") || tableID.contains("$value2")) {// Minaakshi
																					// :
																					// 01-05-2019
					tableID = tableID.replace("$value1", arrTempVal[0].toString());
					tableID = tableID.replace("$value2", arrTempVal[1].toString());
				} else {
					tableID = tableID.replace("$value", uniqueNumber);
				}
			} else {
				tableID = tableID.replace("$value", uniqueNumber);
			}
		} else if (tableID.contains(" OR ")) {
			tableID = tableID.replace("$value1", uniqueNumber);
			tableID = tableID.replace("$value2", uniqueNumber);
		} else if (tableID.contains(" AND ") && !tableID.contains("value1") && !tableID.contains("value2")
				&& !tableID.contains("value3")) {
			tableID = tableID.replace("$value", uniqueNumber);
		} else if (tableID.contains(" AND ")) {
			tableID = tableID.replace("$value1", uniqueNumber);
			tableID = tableID.replace("$value2", uniqueNumber1);
			tableID = tableID.replace("$value3", uniqueNumber2);
		} else {
			tableID = tableID.replace("$value", uniqueNumber);
		}

		return tableID;

	}

	public static String getDataFetchingSQLQuery(String sqlQuery, String ctrlValue) throws IOException {//
																										// 01-10-2019

		String tempVal = "";
		String finalVal = "";
		String[] arrCtrlValue;
		@SuppressWarnings("unused")
		String[] arrCtrlValue1;
		int i;
		int j;
		@SuppressWarnings("unused")
		int k;
		int n;

		if (ctrlValue.contains("*")) { // PRMPREM_291020955723*2697,2698,2699
			arrCtrlValue = ctrlValue.split("\\*");
			for (n = 0; n < arrCtrlValue.length; n++) {
				if (sqlQuery.contains("STRING-$value")) {
					String[] arrTempVal1 = arrCtrlValue[n].split(",");
					j = 0;
					k = arrTempVal1.length - 1;
					for (i = 0; i < arrTempVal1.length; i++) {
						tempVal = "'" + arrTempVal1[i].toString() + "'";
						if (i == j) {
							finalVal = tempVal;
						}
						/*
						 * else if (i==k){ finalVal = finalVal + tempVal; }
						 */
						else {
							finalVal = finalVal + "," + tempVal;
						}
					}
					sqlQuery = sqlQuery.replace("STRING-$value", finalVal);

				} else if (sqlQuery.contains("NUMBER-$value")) {
					String[] arrTempVal1 = arrCtrlValue[n].split(",");
					j = 0;
					k = arrTempVal1.length - 1;
					for (i = 0; i < arrTempVal1.length; i++) {
						tempVal = arrTempVal1[i].toString();
						if (i == j) {
							finalVal = tempVal;
						}
						/*
						 * else if (i==k){ finalVal = finalVal + tempVal; }
						 */
						else {
							finalVal = finalVal + "," + tempVal;
						}
					}
					sqlQuery = sqlQuery.replace("NUMBER-$value", finalVal);
				}
			}

		} else if (sqlQuery.contains("IN (") && sqlQuery.contains("$value1") && sqlQuery.contains("$value2")) {
			String[] arrTempVal = ctrlValue.split(",");
			sqlQuery = sqlQuery.replace("$value1", arrTempVal[0].toString());
			sqlQuery = sqlQuery.replace("$value2", arrTempVal[1].toString());
		} else if (sqlQuery.contains("STRING-$value")) {
			String[] arrTempVal1 = ctrlValue.split(",");
			j = 0;
			k = arrTempVal1.length - 1;
			for (i = 0; i < arrTempVal1.length; i++) {
				tempVal = "'" + arrTempVal1[i].toString() + "'";
				if (i == j) {
					finalVal = tempVal;
				}
				/*
				 * else if (i==k){ finalVal = finalVal + tempVal; }
				 */
				else {
					finalVal = finalVal + "," + tempVal;
				}
			}
			sqlQuery = sqlQuery.replace("STRING-$value", finalVal);
		} else if (sqlQuery.contains("NUMBER-$value")) {
			String[] arrTempVal1 = ctrlValue.split(",");
			j = 0;
			k = arrTempVal1.length - 1;
			for (i = 0; i < arrTempVal1.length; i++) {
				tempVal = arrTempVal1[i].toString();
				if (i == j) {
					finalVal = tempVal;
				}
				/*
				 * else if (i==k){ finalVal = finalVal + tempVal; }
				 */
				else {
					finalVal = finalVal + "," + tempVal;
				}
			}
			sqlQuery = sqlQuery.replace("NUMBER-$value", finalVal);
		} else {
			sqlQuery = sqlQuery.replace("$value", ctrlValue);
		}

		return sqlQuery;

	}

	public static String getDataFetchingSQLQueryPAS(String sqlQuery, String ctrlValue) throws IOException {

		if (sqlQuery.contains("AND (")) {
			String[] arrTempVal = ctrlValue.split(",");
			sqlQuery = sqlQuery.replace("$value1", arrTempVal[0].toString());
			sqlQuery = sqlQuery.replace("$value2", arrTempVal[1].toString());
		} else {
			sqlQuery = sqlQuery.replace("$value", ctrlValue);
		}
		return sqlQuery;
	}

	public static String getDynamicExpectedValue(String strValue, String ActValue) throws IOException {
		String strExptValue = "";
		String tempExpecetedValue = "";
		String tempstrValue = "";
		String tempstrValue1 = "";

		int i = 0;

		if (strValue.contains("$")) {
			String[] arrOfStr1 = strValue.split("\\$");
			tempstrValue = arrOfStr1[0].toString();
			tempstrValue1 = arrOfStr1[1].toString();
		} else {
			tempstrValue = strValue;
		}
		if (strValue.contains("-") && !strValue.contains("$") && !strValue.contains(",")) { //
			String[] arrOfStr1 = tempstrValue.split("\\-");

			for (i = 0; i < arrOfStr1.length; i++) {
				String[] arrOfStr11 = arrOfStr1[i].split("\\|");
				tempstrValue = ReadFromExcelUsingColumnName(arrOfStr11[0].toString(), arrOfStr11[1].toString());
				strValue = strValue.replace(arrOfStr1[i], tempstrValue);
				strExptValue = strValue;
			}

		} else if (strValue.contains("|") && strValue.contains(",")) {
			String[] temp1 = strValue.split("\\,");

			String[] temp2 = temp1[0].split("\\|");
			String tempCtrlValue1 = DMProduct.ReadFromExcelUsingColumnName(temp2[0].toString().trim(),
					temp2[1].toString());

			String[] temp3 = temp1[1].split("\\|");
			String tempCtrlValue2 = DMProduct.ReadFromExcelUsingColumnName(temp3[0].toString().trim(),
					temp3[1].toString());

			strExptValue = tempCtrlValue1 + ", " + tempCtrlValue2 + "," + temp1[2].toString() + ","
					+ temp1[3].toString();

		} else {
			String[] arrOfStr = tempstrValue.split("\\|");

			if (arrOfStr.length > 2) {// 03-10-2018

				for (i = 1; i < arrOfStr.length; i++) {
					tempExpecetedValue = tempExpecetedValue + " " + arrOfStr[i].toString();
					tempExpecetedValue = tempExpecetedValue.trim();
				}

				i = 0;
				for (i = 1; i < arrOfStr.length; i++) {
					tempstrValue = ReadFromExcelUsingColumnName(arrOfStr[0].toString(), arrOfStr[i].toString());
					tempExpecetedValue = tempExpecetedValue.replace(arrOfStr[i].toString(), tempstrValue);
				}
			}

			else {
				tempExpecetedValue = ReadFromExcelUsingColumnName(arrOfStr[0].toString(), arrOfStr[1].toString());
			}

			if ((arrOfStr[1].toString().contains("SSN#")) || (arrOfStr[1].toString().contains("FEIN#"))) {
				String replacedtempExpecetedValue = tempExpecetedValue.replace("-", "");
				tempExpecetedValue = replacedtempExpecetedValue;
			} else if ((arrOfStr[1].toString().contains("Date")) && (ActValue.contains("-"))) {
				String replacedtempExpecetedValue = convertDate(tempExpecetedValue);
				tempExpecetedValue = replacedtempExpecetedValue;
			}

			if (strValue.contains("$")) {
				strExptValue = tempExpecetedValue + tempstrValue1;
			} else {
				strExptValue = tempExpecetedValue;
			}
		}

		return strExptValue;

	}

	public static String getExpectedValueByFormula(String expectedValue, String actualValue, Sheet expectedSheet,
			Cell expectedCell) {

		String req = "";

		DataFormatter fmt = new DataFormatter();

		// 04-01-2019
		HSSFFormulaEvaluator.evaluateAllFormulaCells(expectedSheet.getWorkbook());

		FormulaEvaluator evaluator = expectedSheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

		// CellValue cellValue = evaluator.evaluate(expectedCell);
		req = fmt.formatCellValue(expectedCell, evaluator);

		if (actualValue.contains("-") && !actualValue.contains(".csv")) {
			String replacedtempExpecetedValue = convertDate(req);
			req = replacedtempExpecetedValue;
		}

		return req;
	}

	public static void CopyFilesToServer(String copyFromLocal, String copyToRemote, String fileName) {
		String hostName = Config.flatFileHostName;
		String userName = Config.flatFileUserName;
		String password = Config.flatFilePassword;
		int port = Integer.parseInt(Config.flatFilePort);
		JSch jsch = new JSch();

		Session session = null;
		log.info("Trying to connect.....");

		try {
			// session = jsch.getSession(userName, hostName, 22);
			session = jsch.getSession(userName, hostName, port);
			session.setConfig("StrictHostKeyChecking", "no");
			session.setPassword(password);
			session.connect();

			Channel channel = session.openChannel("sftp");
			channel.connect();
			ChannelSftp sftpChannel = (ChannelSftp) channel;

			log.info("Connection SHTP Established !!");

			Vector<ChannelSftp.LsEntry> filesList = sftpChannel.ls(copyToRemote);

			/* if (!controller.controllerTestCaseID.equals("RGRS_SC_179")) { */// Sel4
			if (!String.valueOf(controller.controllerTestCaseID).equals("RGRS_SC_179")) {
				if ((!filesList.isEmpty() || filesList != null)) {

					for (LsEntry file : filesList) {
						log.info("In File Delete Loop");
						log.info("In File Delete Loop");
						if (!".".equals(file.getFilename()) && !"..".equals(file.getFilename())) {
							sftpChannel.rm(copyToRemote + "//" + file.getFilename());
							log.info("File has been deleted");
							log.info("File has been deleted");
						}
					}
				}
			}
			sftpChannel.lcd(copyFromLocal);

			// Copy flat from Local to remote linux server
			sftpChannel.cd(copyToRemote);

			sftpChannel.put(copyFromLocal + "\\" + fileName, copyToRemote);
			log.info("File has been copied to FTP -> Inbound folder.");
			log.info("File has been copied to FTP -> Inbound folder.");

			sftpChannel.exit();
			session.disconnect();

		} catch (JSchException e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
		} catch (SftpException e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
		}
	}

	public static void startDMRecovery() {
		log.info("IN : DM Recovery");
		log.info("IN : DM Recovery");

		Automation.driver.quit();
		try {
			Automation.setUp();
			Thread.sleep(5000);

			WebElement userName = Automation.driver.findElement(By.xpath("//input[@id='userID']"));
			WebElement password = Automation.driver.findElement(By.xpath("//input[@id='userPass']"));
			WebElement loginBtn = Automation.driver.findElement(By.xpath("//button[@id='btnLogin']"));

			userName.sendKeys("minaakshi");
			password.sendKeys("icd123");
			loginBtn.click();
			Thread.sleep(5000);

			WebElement multiAppLnk = Automation.driver.findElement(By.xpath("//a[@id='multiAppLink']"));
			multiAppLnk.click();
			Thread.sleep(2000);

			WebElement disableMultiAppLnk = Automation.driver.findElement(By.xpath(".//*[@id='disableMultiApp']"));
			disableMultiAppLnk.click();
			Thread.sleep(1000);

			WebElement yesBtn = Automation.driver
					.findElement(By.xpath("//button[@class='btn btn-popup' and @name='Ok']"));
			yesBtn.click();
			Thread.sleep(2000);

			log.info("OUT : DM Recovery");
			log.info("OUT : DM Recovery");

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void CopyFilesToServer() {
		String hostName = Config.flatFileHostName;
		String userName = Config.flatFileUserName;
		String password = Config.flatFilePassword;
		int port = Integer.parseInt(Config.flatFilePort);
		JSch jsch = new JSch();

		Session session = null;
		log.info("Trying to connect.....");

		source = Config.inputDataFilePath.toString() + Local_Path;
		destination = Remote_Path;

		try {
			// session = jsch.getSession(userName, hostName, 22);
			session = jsch.getSession(userName, hostName, port);
			session.setConfig("StrictHostKeyChecking", "no");
			session.setConfig("PreferredAuthentications", "password");//For keberos error - 07/02/2025
			session.setPassword(password);
			session.connect();

			Channel channel = session.openChannel("sftp");
			channel.connect();
			ChannelSftp sftpChannel = (ChannelSftp) channel;

			log.info("Connection SHTP Established !!");

			Vector<ChannelSftp.LsEntry> filesList = sftpChannel.ls(destination);

			if (!controller.controllerTestCaseID.toString().equals("RGRS_SC_179")) {// : 01-02-2021
				if (!filesList.isEmpty() || filesList != null) {
					
					//Below given code changed for DM SFTP - BSP-59651 - Issue - 17/02/2025
					for (ChannelSftp.LsEntry file : filesList) {
					    String fileName = file.getFilename();
					    log.info("fileName is : " + fileName);
					    // Skip "." and ".."
					    if (".".equals(fileName) || "..".equals(fileName)) {
					        continue;
					    }

					    // Check if it's a file before deleting
					    if (file.getAttrs().isReg()) {
					        log.info("Processing file: " + fileName);
					        sftpChannel.rm(destination + "/" + fileName);
					        log.info("File deleted successfully.");
					    } else {
					        log.info("Skipping directory: " + fileName);
					    }
					}
					
					//Old code - 17/02/2025
					/*	for (LsEntry file : filesList) {
					
					log.info("IFile Name is: " + file );
					log.info("In File Delete Loop");
					//log.info("In File Delete Loop");
					if (!".".equals(file.getFilename()) && !"..".equals(file.getFilename())) {
						log.info(destination + "/" + file.getFilename());

						if (!file.getFilename().equals("RegEdUploadFolder")) {
							sftpChannel.rm(destination + "/" + file.getFilename());
						}
						//log.info("File has been deleted");
						log.info("File has been deleted");
					}
				}*/

				}
			}
			sftpChannel.lcd(source);
			log.info("Source path is : " + source);
			// Copy flat from Local to remote linux server
			sftpChannel.cd(destination);
			log.info("Destination path is : " + destination);

			sftpChannel.put(source + "\\" + FileNameToBeCreated, destination);
			//log.info("File has been copied to FTP -> Inbound folder.");
			log.info("File: " + FileNameToBeCreated + " - has been copied from: " + source + " to destination : " +  destination);

			sftpChannel.exit();
			session.disconnect();

		} catch (JSchException e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
		} catch (SftpException e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
		}
	}

	public static void CopyFilesToWindowsServer() throws IOException {
		// TODO Auto-generated method stub
		File src = new File(Config.inputDataFilePath.toString() + Local_Path + "\\" + FileNameToBeCreated);
		File dest = new File(Config.copyServerRemotePath.toString() + "/" + FileNameToBeCreated);

		if (src.exists()) {
			FileUtils.copyFile(src, dest);

		} else {
			webDriver.getReport().setMessage("File has not been copied");
			log.error("File has not been copied");
			throw new IOException("File has not been copied");
		}
	}

	public static void CopyFilesToLocal() {
		String hostName = Config.flatFileHostName;
		String userName = Config.flatFileUserName;
		String password = Config.flatFilePassword;
		int port = Integer.parseInt(Config.flatFilePort);
		JSch jsch = new JSch();

		Session session = null;
		log.info("Trying to connect.....");

		source = Remote_Path_Outbound;
		destination = Config.inputDataFilePath.toString() + Local_Path;

		try {
			// session = jsch.getSession(userName, hostName, 22);
			session = jsch.getSession(userName, hostName, port);
			session.setConfig("StrictHostKeyChecking", "no");
			session.setConfig("PreferredAuthentications", "password");//For keberos error - 07/02/2025
			session.setPassword(password);
			session.connect();
			Channel channel = session.openChannel("sftp");
			channel.connect();
			ChannelSftp sftpChannel = (ChannelSftp) channel;
			log.info("Connection SFTP Established !!");
			sftpChannel.cd(source);

			Vector<ChannelSftp.LsEntry> entries = sftpChannel.ls(source);

			// download all files (except the ., .. and folders) from given folder
			for (ChannelSftp.LsEntry en : entries) {

				if (en.getFilename().contains(FileName)) {

					try {
						sftpChannel.get(en.getFilename(), destination);
						log.info("Downloading " + (source + en.getFilename()) + " ----> " + "download"
								+ File.separator + en.getFilename());
						sftpChannel.get(source + en.getFilename(), "download" + File.separator + en.getFilename());

					} catch (SftpException e) {
						// log.error(e.getMessage(), e);
						e.printStackTrace();
					}

				} else {
					continue;
				}

			}
			sftpChannel.exit();
			session.disconnect();
		} catch (JSchException e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
		} catch (SftpException e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
		}
	}

	public static void CreateFlatFile(String logicalName, String ctrlValue) throws IOException {

		// TODO Auto-generated method stub
		if (logicalName.equalsIgnoreCase("Local_Path") && !ctrlValue.equals(""))
			Local_Path = ctrlValue;
		else if (logicalName.equalsIgnoreCase("Template_Path") && !ctrlValue.equals(""))
			Template_Path = ctrlValue;
		else if (logicalName.equalsIgnoreCase("Remote_Path") && !ctrlValue.equals(""))
			Remote_Path = ctrlValue;
		else if (logicalName.equalsIgnoreCase("Remote_Path_Outbound") && !ctrlValue.equals(""))
			Remote_Path_Outbound = ctrlValue;
		else if (logicalName.equalsIgnoreCase("FileName") && !ctrlValue.equals(""))
			FileName = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EffectiveDate") && !ctrlValue.equals(""))// 01-10-2019
			if (ctrlValue.contains("$value")) {
				String strTemp = null;
				String strTemp1 = null;
				strTemp = DMProduct.ReadFromExcelUsingColumnName("", "SchemeEvaluationToDate");
				String arrTemp[] = strTemp.split("\\/");
				strTemp1 = arrTemp[2] + arrTemp[0] + arrTemp[1];// 10/17/2020 :
																// Required
																// 17/10/2020
				ctrlValue = ctrlValue.replace("$value", strTemp1);
				EffectiveDate = ctrlValue;
			} else {
				EffectiveDate = ctrlValue;
			}
		else if (logicalName.equalsIgnoreCase("EffectiveDate1") && !ctrlValue.equals(""))
			EffectiveDate1 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EffectiveDate2") && !ctrlValue.equals(""))
			EffectiveDate2 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EffectiveDate3") && !ctrlValue.equals(""))
			EffectiveDate3 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EffectiveDate4") && !ctrlValue.equals(""))
			EffectiveDate4 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EffectiveDate5") && !ctrlValue.equals(""))
			EffectiveDate5 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EffectiveDate6") && !ctrlValue.equals(""))
			EffectiveDate6 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EffectiveDate7") && !ctrlValue.equals(""))
			EffectiveDate7 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EffectiveDate8") && !ctrlValue.equals(""))
			EffectiveDate8 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EffectiveDate9") && !ctrlValue.equals(""))
			EffectiveDate9 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EffectiveDate_10") && !ctrlValue.equals(""))
			EffectiveDate_10 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EffectiveDate_11") && !ctrlValue.equals(""))
			EffectiveDate_11 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("RiskCommencementDate") && !ctrlValue.equals(""))
			if (ctrlValue.contains("$value")) {
				String strTemp = null;
				String strTemp1 = null;
				strTemp = DMProduct.ReadFromExcelUsingColumnName("", "SchemeEvaluationToDate");
				String arrTemp[] = strTemp.split("\\/");
				strTemp1 = arrTemp[2] + arrTemp[0] + arrTemp[1];
				ctrlValue = ctrlValue.replace("$value", strTemp1);
				RiskCommencementDate = ctrlValue;
			} else {
				RiskCommencementDate = ctrlValue;
			}
		else if (logicalName.equalsIgnoreCase("RiskCommencementDate1") && !ctrlValue.equals(""))
			RiskCommencementDate1 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("RiskCommencementDate2") && !ctrlValue.equals(""))
			RiskCommencementDate2 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("PolicyEffectiveToDate") && !ctrlValue.equals(""))
			PolicyEffectiveToDate = ctrlValue;
		else if (logicalName.equalsIgnoreCase("ProposalApplDate") && !ctrlValue.equals(""))
			ProposalApplDate = ctrlValue;
		else if (logicalName.equalsIgnoreCase("ProposalSubmDate") && !ctrlValue.equals(""))
			ProposalSubmDate = ctrlValue;
		else if (logicalName.equalsIgnoreCase("TemplateFileName") && !ctrlValue.equals(""))
			TemplateFileName = ctrlValue;
		else if (logicalName.equalsIgnoreCase("FileNameToBeCreated") && !ctrlValue.equals(""))
			if (ctrlValue.contains("$value")) {
				String strTemp = null;
				String strTemp1 = null;
				strTemp = ReadFromExcelUsingColumnName("", "SchemeEvaluationToDate");
				String arrTemp[] = strTemp.split("\\/");
				strTemp1 = arrTemp[2] + "-" + arrTemp[0] + "-" + arrTemp[1];
				ctrlValue = ctrlValue.replace("$value", strTemp1);
				FileNameToBeCreated = ctrlValue;
			} else {
				FileNameToBeCreated = ctrlValue;
			}
		else if (logicalName.equalsIgnoreCase("FileNameCopiedToLocal") && !ctrlValue.equals(""))
			FileNameCopiedToLocal = ctrlValue;
		// 01-05-2019
		else if (logicalName.equalsIgnoreCase("PolicyNumber") && !ctrlValue.equals(""))
			PolicyNumber = ctrlValue;
		// 01-01-2021
		else if (logicalName.equalsIgnoreCase("GroupFileLoadNumber") && !ctrlValue.equals(""))
			GroupFileLoadNumber = ctrlValue;
		// 01-05-2019
		else if (logicalName.equalsIgnoreCase("PolicyFileLoadNumber") && !ctrlValue.equals(""))
			PolicyFileLoadNumber = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);

		// 04 Aug 22
		else if (logicalName.equalsIgnoreCase("SASubGrpNo1") && !ctrlValue.equals(""))
			SASubGrpNo1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName) + "0001";
		else if (logicalName.equalsIgnoreCase("SASubGrpNo2") && !ctrlValue.equals(""))
			SASubGrpNo2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName) + "0002";
		else if (logicalName.equalsIgnoreCase("SASubGrpNo3") && !ctrlValue.equals(""))
			SASubGrpNo3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName) + "0001";
		else if (logicalName.equalsIgnoreCase("SASubGrpNo4") && !ctrlValue.equals(""))
			SASubGrpNo4 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName) + "0002";
		else if (logicalName.equalsIgnoreCase("PremiumFileLoadNumber") && !ctrlValue.equals(""))
			PremiumFileLoadNumber = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		if (logicalName.equalsIgnoreCase("EntityCode1") && !ctrlValue.equals(""))
			EntityCode1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("EntityCode2") && !ctrlValue.equals(""))
			EntityCode2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("EntityCode3") && !ctrlValue.equals(""))
			EntityCode3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("EntityCode4") && !ctrlValue.equals(""))
			EntityCode4 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("EntityCode5") && !ctrlValue.equals(""))
			EntityCode5 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("PayoutHeaderID") && !ctrlValue.equals(""))
			PayoutHeaderID = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);// 01-01-2020
		else if (logicalName.equalsIgnoreCase("SSED1") && !ctrlValue.equals(""))
			SSED1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SSN1") && !ctrlValue.equals(""))
			SSN1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("CSCI1") && !ctrlValue.equals(""))
			CSCI1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("MemberName1") && !ctrlValue.equals(""))
			MemberName1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SSED2") && !ctrlValue.equals(""))
			SSED2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SSN2") && !ctrlValue.equals(""))
			SSN2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("CSCI2") && !ctrlValue.equals(""))
			CSCI2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("MemberName2") && !ctrlValue.equals(""))
			MemberName2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("PolicyNumber1") && !ctrlValue.equals(""))
			PolicyNumber1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("PolicyNumber2") && !ctrlValue.equals(""))
			PolicyNumber2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("PolicyNumber3") && !ctrlValue.equals(""))
			PolicyNumber3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("PolicyNumber4") && !ctrlValue.equals(""))
			PolicyNumber4 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("PolicyNumber5") && !ctrlValue.equals(""))
			PolicyNumber5 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);

		else if (logicalName.equalsIgnoreCase("ApxGrpNo") && !ctrlValue.equals(""))
			ApxGrpNo = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("ApxGrpNo0002") && !ctrlValue.equals(""))
			ApxGrpNo0002 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName) + "0002";
		else if (logicalName.equalsIgnoreCase("ApxGrpNo0001") && !ctrlValue.equals(""))
			ApxGrpNo0001 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName) + "0001";

		else if (logicalName.equalsIgnoreCase("SubGrpNo") && !ctrlValue.equals(""))// 01-09-2020
			SubGrpNo = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName) + "0001";
		// Mass Contracting variable : START
		else if (logicalName.equalsIgnoreCase("AGENCY_ID") && !ctrlValue.equals(""))
			AGENCY_ID = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("NPN1") && !ctrlValue.equals(""))
			NPN1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("NPN2") && !ctrlValue.equals(""))
			NPN2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("NPN3") && !ctrlValue.equals(""))
			NPN3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("NPN4") && !ctrlValue.equals(""))
			NPN4 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("NPN5") && !ctrlValue.equals(""))
			NPN5 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("FEIN") // : 06-17-2022
				&& !ctrlValue.equals(""))
			FEIN = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SSN3") && !ctrlValue.equals(""))
			SSN3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SSN4") && !ctrlValue.equals(""))
			SSN4 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SSN5") && !ctrlValue.equals(""))
			SSN5 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("FN1") && !ctrlValue.equals(""))
			FN1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("FN2") && !ctrlValue.equals(""))
			FN2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("FN3") && !ctrlValue.equals(""))
			FN3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("FN4") && !ctrlValue.equals(""))
			FN4 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("FN5") && !ctrlValue.equals(""))
			FN5 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("MN1") && !ctrlValue.equals(""))
			MN1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("MN2") && !ctrlValue.equals(""))
			MN2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("MN3") && !ctrlValue.equals(""))
			MN3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("MN4") && !ctrlValue.equals(""))
			MN4 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("MN5") && !ctrlValue.equals(""))
			MN5 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("LN1") && !ctrlValue.equals(""))
			LN1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("LN2") && !ctrlValue.equals(""))
			LN2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("LN3") && !ctrlValue.equals(""))
			LN3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("LN4") && !ctrlValue.equals(""))
			LN4 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("LN5") && !ctrlValue.equals(""))
			LN5 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("EP_CODE1") && !ctrlValue.equals(""))
			EP_CODE1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("EP_CODE2") && !ctrlValue.equals(""))
			EP_CODE2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("EP_CODE3") && !ctrlValue.equals(""))
			EP_CODE3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("EP_CODE4") && !ctrlValue.equals(""))
			EP_CODE4 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("EP_CODE5") && !ctrlValue.equals(""))
			EP_CODE5 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("NSCC_NO1") && !ctrlValue.equals(""))
			NSCC_NO1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("NSCC_NO2") && !ctrlValue.equals(""))
			NSCC_NO2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("NSCC_NO3") && !ctrlValue.equals(""))
			NSCC_NO3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("NSCC_NO4") && !ctrlValue.equals(""))
			NSCC_NO4 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("NSCC_NO5") && !ctrlValue.equals(""))
			NSCC_NO5 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("ACCOUNT_NO1") && !ctrlValue.equals(""))
			ACCOUNT_NO1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("ACCOUNT_NO2") && !ctrlValue.equals(""))
			ACCOUNT_NO2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("ACCOUNT_NO3") && !ctrlValue.equals(""))
			ACCOUNT_NO3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("ACCOUNT_NO4") && !ctrlValue.equals(""))
			ACCOUNT_NO4 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("ACCOUNT_NO5") && !ctrlValue.equals(""))
			ACCOUNT_NO5 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("POL_NO1") && !ctrlValue.equals(""))
			POL_NO1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("POL_NO2") && !ctrlValue.equals(""))
			POL_NO2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("POL_NO3") && !ctrlValue.equals(""))
			POL_NO3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("POL_NO4") && !ctrlValue.equals(""))
			POL_NO4 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("POL_NO5") && !ctrlValue.equals(""))
			POL_NO5 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("EAOEF_DATE1") && !ctrlValue.equals("")) {
			log.info("Testing massonboarding" + ctrlValue);
			log.info("Testing massonboarding" + ctrlValue);
			EAOEF_DATE1 = ctrlValue;
			log.info("Testing massonboarding" + EAOEF_DATE1);
		} else if (logicalName.equalsIgnoreCase("EAOEF_DATE2") && !ctrlValue.equals(""))
			EAOEF_DATE2 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EAOEF_DATE3") && !ctrlValue.equals(""))
			EAOEF_DATE3 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EAOEF_DATE4") && !ctrlValue.equals(""))
			EAOEF_DATE4 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EAOEF_DATE5") && !ctrlValue.equals(""))
			EAOEF_DATE5 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EAOEF_DATE") && !ctrlValue.equals(""))
			EAOEF_DATE = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EF_DATE1") && !ctrlValue.equals(""))
			EF_DATE1 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EF_DATE2") && !ctrlValue.equals(""))
			EF_DATE2 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EF_DATE3") && !ctrlValue.equals(""))
			EF_DATE3 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EF_DATE4") && !ctrlValue.equals(""))
			EF_DATE4 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EF_DATE5") && !ctrlValue.equals(""))
			EF_DATE5 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EF_DATE") && !ctrlValue.equals(""))
			EF_DATE = ctrlValue;
		else if (logicalName.equalsIgnoreCase("ET_DATE1") && !ctrlValue.equals(""))
			ET_DATE1 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("ET_DATE2") && !ctrlValue.equals(""))
			ET_DATE2 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("ET_DATE3") && !ctrlValue.equals(""))
			ET_DATE3 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("ET_DATE4") && !ctrlValue.equals(""))
			ET_DATE4 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("ET_DATE5") && !ctrlValue.equals(""))
			ET_DATE5 = ctrlValue;
		else if (logicalName.equalsIgnoreCase("LICENSE_NO1") && !ctrlValue.equals(""))
			LICENSE_NO1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("LICENSE_NO2") && !ctrlValue.equals(""))
			LICENSE_NO2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("LICENSE_NO3") && !ctrlValue.equals(""))
			LICENSE_NO3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("LICENSE_NO4") && !ctrlValue.equals(""))
			LICENSE_NO4 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("LICENSE_NO5") && !ctrlValue.equals(""))
			LICENSE_NO5 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SC_CODE1") && !ctrlValue.equals(""))
			SC_CODE1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SC_CODE2") && !ctrlValue.equals(""))
			SC_CODE2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SC_CODE3") && !ctrlValue.equals(""))
			SC_CODE3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SC_CODE4") && !ctrlValue.equals(""))
			SC_CODE4 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SC_CODE5") && !ctrlValue.equals(""))
			SC_CODE5 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		// Mass Contracting Variable : END

		// Training Variable : START
		else if (logicalName.equalsIgnoreCase("Course_Completion_Date") && !ctrlValue.equals(""))
			COURSE_COMPLETION_DATE = ctrlValue;
		else if (logicalName.equalsIgnoreCase("Course_Certification_Date") && !ctrlValue.equals(""))
			COURSE_CERTIFICATION_DATE = ctrlValue;
		else if (logicalName.equalsIgnoreCase("Course_Submission_Date") && !ctrlValue.equals(""))
			COURSE_SUBMISSION_DATE = ctrlValue;
		else if (logicalName.equalsIgnoreCase("Hour") && !ctrlValue.equals(""))
			HOUR = ctrlValue;
		else if (logicalName.equalsIgnoreCase("Enroll_Date") && !ctrlValue.equals(""))
			ENROLL_DATE = ctrlValue;
		else if (logicalName.equalsIgnoreCase("Status_Date") && !ctrlValue.equals(""))
			STATUS_DATE = ctrlValue;
		else if (logicalName.equalsIgnoreCase("StatusDate1") && !ctrlValue.equals(""))
			STATUSDATE1 = ctrlValue;
		// Training Variable : END
		else if (logicalName.equalsIgnoreCase("CreateFile") && ctrlValue.equals("Yes")) {
			copySource = new File(Config.inputDataFilePath.toString() + Template_Path + "//" + TemplateFileName);
			copyDestination = new File(Config.inputDataFilePath.toString() + Local_Path + "//" + FileNameToBeCreated);
			FileUtils.copyFile(copySource, copyDestination);

			// Open created and replace Transaction Date and
			// EntityCodes

			String originalFileContent = "";
			BufferedReader reader = null;
			BufferedWriter writer = null;

			try {
				reader = new BufferedReader(new FileReader(copyDestination));
				String currentReadingLine = reader.readLine();
				while (currentReadingLine != null) {
					originalFileContent += currentReadingLine;
					originalFileContent += System.lineSeparator();
					currentReadingLine = reader.readLine();
				}

				String modifiedFileContent = originalFileContent.replaceFirst("[\n\r]+$", "");
				modifiedFileContent = modifiedFileContent.replaceAll("\\$DATE", EffectiveDate);

				if (!EffectiveDate1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE", EffectiveDate1);

				if (!EffectiveDate2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE2", EffectiveDate2);

				if (!EffectiveDate3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE3", EffectiveDate3);

				if (!EffectiveDate4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE4", EffectiveDate4);

				if (!EffectiveDate5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE5", EffectiveDate5);

				if (!EffectiveDate6.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE6", EffectiveDate6);

				if (!EffectiveDate7.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE7", EffectiveDate7);

				if (!EffectiveDate8.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE8", EffectiveDate8);

				if (!EffectiveDate9.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE9", EffectiveDate9);

				if (!EffectiveDate_10.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE_10", EffectiveDate_10);

				if (!EffectiveDate_11.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE_11", EffectiveDate_11);

				if (!RiskCommencementDate1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$RC_DATE1", RiskCommencementDate1);

				if (!RiskCommencementDate2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$RC_DATE2", RiskCommencementDate2);

				if (!RiskCommencementDate.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$RC_DATE", RiskCommencementDate);

				if (!PolicyEffectiveToDate.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$PE_DATE", PolicyEffectiveToDate);

				if (!ProposalApplDate.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$PA_DATE", ProposalApplDate);

				if (!ProposalSubmDate.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$PS_DATE", ProposalSubmDate);

				if (WebHelperDM.transactionType.toString().equalsIgnoreCase("PolicyInterface")) {// Minaakshi
					// :
					// 01-05-2019
					if (!StringUtils.equalsIgnoreCase(PolicyFileLoadNumber, "")
							|| !StringUtils.equalsIgnoreCase(PolicyFileLoadNumber, null)) {
						modifiedFileContent = modifiedFileContent.replaceAll("\\$POLFLN", PolicyFileLoadNumber);
					}
				}

				if (WebHelperDM.transactionType.toString().equalsIgnoreCase("PremiumInterface")) {// Minaakshi
					// :
					// 01-05-2019
					if (!StringUtils.equalsIgnoreCase(PremiumFileLoadNumber, "")
							|| !StringUtils.equalsIgnoreCase(PremiumFileLoadNumber, null)) {
						modifiedFileContent = modifiedFileContent.replaceAll("\\$PREFLN", PremiumFileLoadNumber);
					}
				}

				// Minaakshi : 01-01-2021
				if (!GroupFileLoadNumber.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$GFLN", GroupFileLoadNumber);

				if (!StringUtils.equalsIgnoreCase(PolicyNumber, "")
						|| !StringUtils.equalsIgnoreCase(PolicyNumber, null))// Minaakshi
																				// :
																				// 01-05-2019
					modifiedFileContent = modifiedFileContent.replaceAll("\\$POL_NUMBER", PolicyNumber);
				// Added by on 8/4/2022
				if (!SASubGrpNo1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SASubGrpNo1", SASubGrpNo1);
				if (!SASubGrpNo2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SASubGrpNo2", SASubGrpNo2);

				if (!SASubGrpNo3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SASubGrpNo3", SASubGrpNo3);
				if (!SASubGrpNo4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SASubGrpNo4", SASubGrpNo4);
				if (!PolicyNumber1.equals("")) // 01-05-2019
					modifiedFileContent = modifiedFileContent.replaceAll("\\$PolicyNumber1", PolicyNumber1);

				if (!PolicyNumber2.equals("")) // 01-05-2019
					modifiedFileContent = modifiedFileContent.replaceAll("\\$PolicyNumber2", PolicyNumber2);

				if (!PolicyNumber3.equals("")) // 01-05-2019
					modifiedFileContent = modifiedFileContent.replaceAll("\\$PolicyNumber3", PolicyNumber3);

				if (!PolicyNumber4.equals("")) // 01-05-2019
					modifiedFileContent = modifiedFileContent.replaceAll("\\$PolicyNumber4", PolicyNumber4);

				if (!PolicyNumber5.equals("")) // 01-05-2019
					modifiedFileContent = modifiedFileContent.replaceAll("\\$PolicyNumber5", PolicyNumber5);

				if (!ApxGrpNo.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ApxGrpNo", ApxGrpNo);
				if (!ApxGrpNo0002.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ApxGrpNo0002", ApxGrpNo0002);

				if (!ApxGrpNo0001.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ApxGrpNo0001", ApxGrpNo0001);
				if (!SubGrpNo.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SubGrpNo", SubGrpNo);

				if (!EntityCode1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ENTITYCODE1", EntityCode1);

				if (!EntityCode2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ENTITYCODE2", EntityCode2);

				if (!EntityCode3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ENTITYCODE3", EntityCode3);

				if (!EntityCode4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ENTITYCODE4", EntityCode4);

				if (!EntityCode5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ENTITYCODE5", EntityCode5);

				if (!PayoutHeaderID.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$PHID", PayoutHeaderID);// 01-01-2020

				if (!SSED1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SSED1", SSED1);

				if (!SSN1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SSN1", SSN1);

				if (!CSCI1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$CSCI1", CSCI1);

				if (!MemberName1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$MemberName1", MemberName1);

				if (!SSED2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SSED2", SSED2);

				if (!SSN2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SSN2", SSN2);

				if (!CSCI2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$CSCI2", CSCI2);

				if (!MemberName2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$MemberName2", MemberName2);
				// Mass Contracting Variable Replacement : START
				if (!AGENCY_ID.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$AGENCY_ID", AGENCY_ID);

				if (!NPN1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$NPN1", NPN1);

				if (!NPN2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$NPN2", NPN2);

				if (!NPN3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$NPN3", NPN3);

				if (!NPN4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$NPN4", NPN4);

				if (!NPN5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$NPN5", NPN5);

				if (!SSN3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SSN3", SSN3);

				if (!SSN4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SSN4", SSN4);

				if (!SSN5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SSN5", SSN5);

				if (!FN1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$FN1", FN1);

				if (!FN2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$FN2", FN2);

				if (!FN3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$FN3", FN3);

				if (!FN4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$FN4", FN4);

				if (!FN5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$FN5", FN5);

				if (!MN1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$MN1", MN1);

				if (!MN2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$MN2", MN2);

				if (!MN3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$MN3", MN3);

				if (!MN4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$MN4", MN4);

				if (!MN5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$MN5", MN5);

				if (!LN1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$LN1", LN1);

				if (!LN2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$LN2", LN2);

				if (!LN3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$LN3", LN3);

				if (!LN4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$LN4", LN4);

				if (!LN5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$LN5", LN5);

				if (!EP_CODE1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EP_CODE1", EP_CODE1);

				if (!EP_CODE2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EP_CODE2", EP_CODE2);

				if (!EP_CODE3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EP_CODE3", EP_CODE3);

				if (!EP_CODE4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EP_CODE4", EP_CODE4);

				if (!EP_CODE5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EP_CODE5", EP_CODE5);

				if (!NSCC_NO1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$NSCC_NO1", NSCC_NO1);

				if (!NSCC_NO2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$NSCC_NO2", NSCC_NO2);

				if (!NSCC_NO3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$NSCC_NO3", NSCC_NO3);

				if (!NSCC_NO4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$NSCC_NO4", NSCC_NO4);

				if (!NSCC_NO5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$NSCC_NO5", NSCC_NO5);

				if (!ACCOUNT_NO1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ACCOUNT_NO1", ACCOUNT_NO1);

				if (!ACCOUNT_NO2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ACCOUNT_NO2", ACCOUNT_NO2);

				if (!ACCOUNT_NO3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ACCOUNT_NO3", ACCOUNT_NO3);

				if (!ACCOUNT_NO4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ACCOUNT_NO4", ACCOUNT_NO4);

				if (!ACCOUNT_NO5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ACCOUNT_NO5", ACCOUNT_NO5);

				if (!POL_NO1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$POL_NO1", POL_NO1);

				if (!POL_NO2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$POL_NO2", POL_NO2);

				if (!POL_NO3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$POL_NO3", POL_NO3);

				if (!POL_NO4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$POL_NO4", POL_NO4);

				if (!POL_NO5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$POL_NO5", POL_NO5);

				/*
				 * if (!EAOEF_DATE1.equals("")) modifiedFileContent =
				 * modifiedFileContent.replaceAll( "\\$EAOEF_DATE1", EAOEF_DATE1);
				 */
				if (!EAOEF_DATE1.equals("")) {
					log.info("checking EAOEF_DATE1 in DMProduct" + EAOEF_DATE1);
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EAOEF_DATE1", EAOEF_DATE1);
				}
				if (!EAOEF_DATE2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EAOEF_DATE2", EAOEF_DATE2);

				if (!EAOEF_DATE3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EAOEF_DATE3", EAOEF_DATE3);

				if (!EAOEF_DATE4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EAOEF_DATE4", EAOEF_DATE4);

				if (!EAOEF_DATE5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EAOEF_DATE5", EAOEF_DATE5);

				if (!EAOEF_DATE.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EAOEF_DATE", EAOEF_DATE);

				if (!EF_DATE1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE1", EF_DATE1);

				if (!EF_DATE2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE2", EF_DATE2);

				if (!EF_DATE3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE3", EF_DATE3);

				if (!EF_DATE4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE4", EF_DATE4);

				if (!EF_DATE5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE5", EF_DATE5);

				if (!EF_DATE.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EF_DATE", EF_DATE);

				if (!ET_DATE1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ET_DATE1", ET_DATE1);

				if (!ET_DATE2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ET_DATE2", ET_DATE2);

				if (!ET_DATE3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ET_DATE3", ET_DATE3);

				if (!ET_DATE4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ET_DATE4", ET_DATE4);

				if (!ET_DATE5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ET_DATE5", ET_DATE5);

				if (!LICENSE_NO1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$LICENSE_NO1", LICENSE_NO1);

				if (!LICENSE_NO2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$LICENSE_NO2", LICENSE_NO2);

				if (!LICENSE_NO3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$LICENSE_NO3", LICENSE_NO3);

				if (!LICENSE_NO4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$LICENSE_NO4", LICENSE_NO4);

				if (!LICENSE_NO5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$LICENSE_NO5", LICENSE_NO5);

				if (!SC_CODE1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SC_CODE1", SC_CODE1);

				if (!SC_CODE2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SC_CODE2", SC_CODE2);

				if (!SC_CODE3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SC_CODE3", SC_CODE3);

				if (!SC_CODE4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SC_CODE4", SC_CODE4);

				if (!SC_CODE5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SC_CODE5", SC_CODE5);

				// Mass Contracting Variable Replacement : END

				// Training Variable Replacement : START

				if (!COURSE_COMPLETION_DATE.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$COURSE_COMPLETION_DATE",
							COURSE_COMPLETION_DATE);
				if (!COURSE_CERTIFICATION_DATE.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$COURSE_CERTIFICATION_DATE",
							COURSE_CERTIFICATION_DATE);
				if (!COURSE_SUBMISSION_DATE.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$COURSE_SUBMISSION_DATE",
							COURSE_SUBMISSION_DATE);
				if (!HOUR.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$HOUR", HOUR);
				if (!ENROLL_DATE.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ENROLL_DATE", ENROLL_DATE);
				if (!STATUS_DATE.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$STATUS_DATE", STATUS_DATE);
				if (!STATUSDATE1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$STATUSDATE1", STATUSDATE1);
				// Training Variable Replacement : END

				writer = new BufferedWriter(new FileWriter(copyDestination));

				writer.write(modifiedFileContent);
			} catch (IOException e) {
				log.error(e.getMessage(), e);
				e.printStackTrace();
			} finally {

				try {
					if (reader != null) {
						reader.close();
					}
					if (writer != null) {
						writer.close();
					}
				} catch (IOException e) {
					log.error(e.getMessage(), e);
					e.printStackTrace();
				}
			}
		}

	}

	public static void CreateWebServiceRequest(String logicalName, String ctrlValue) throws IOException {
		// TODO Auto-generated method stub

		if (logicalName.equalsIgnoreCase("ApxGrpAddr1") && !ctrlValue.equals(""))
			ApxGrpAddr1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("ApxGrpEmailId") && !ctrlValue.equals(""))
			ApxGrpEmailId = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("ApxGrpNo") && !ctrlValue.equals(""))
			ApxGrpNo = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("ApxGrpName") && !ctrlValue.equals(""))
			ApxGrpName = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("ApxSSPK") && !ctrlValue.equals(""))
			ApxSSPK = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SSRN") && !ctrlValue.equals(""))
			SSRN = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("RFPN") && !ctrlValue.equals(""))
			RFPN = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("PolicyNumber1") && !ctrlValue.equals(""))
			PolicyNumber1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("PolicyNumber2") && !ctrlValue.equals(""))
			PolicyNumber2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("CSCI1") && !ctrlValue.equals(""))
			CSCI1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("CSCI2") && !ctrlValue.equals(""))
			CSCI2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("EntityCode1") && !ctrlValue.equals(""))
			EntityCode1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("EntityCode2") && !ctrlValue.equals(""))
			EntityCode2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("EntityCode3") && !ctrlValue.equals(""))
			EntityCode3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("EntityCode4") && !ctrlValue.equals(""))
			EntityCode4 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("EntityCode5") && !ctrlValue.equals(""))
			EntityCode5 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SubGrpAddr1") && !ctrlValue.equals(""))
			SubGrpAddr1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SubGrpAddr2") && !ctrlValue.equals(""))
			SubGrpAddr2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SubGrpAddr3") && !ctrlValue.equals(""))
			SubGrpAddr3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SubGrpSSEC1") && !ctrlValue.equals(""))
			SubGrpSSEC1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SubGrpSSEC2") && !ctrlValue.equals(""))
			SubGrpSSEC2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SubGrpSSEC3") && !ctrlValue.equals(""))
			SubGrpSSEC3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);

		//
		else if (logicalName.equalsIgnoreCase("SubGrpName1") && !ctrlValue.equals(""))
			SubGrpName1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SubGrpName2") && !ctrlValue.equals(""))
			SubGrpName2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SubGrpName3") && !ctrlValue.equals(""))
			SubGrpName3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SubGrpSSPK1") && !ctrlValue.equals(""))
			SubGrpSSPK1 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SubGrpSSPK2") && !ctrlValue.equals(""))
			SubGrpSSPK2 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("SubGrpSSPK3") && !ctrlValue.equals(""))
			SubGrpSSPK3 = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		else if (logicalName.equalsIgnoreCase("CreateFile") && ctrlValue.equals("Yes")) {
			// Create file
			File copySource = new File(Config.inputDataFilePath.toString() + "WebService\\WebserviceFiles\\"
					+ WebHelper.request_xml + "_Template.xml");

			File copyDestination = new File(Config.inputDataFilePath.toString() + "WebService\\WebserviceFiles\\"
					+ WebHelper.request_xml + ".xml");
			FileUtils.copyFile(copySource, copyDestination);

			// Open created and replace Transaction Date and
			// EntityCodes

			String originalFileContent = "";
			BufferedReader reader = null;
			BufferedWriter writer = null;

			try {
				reader = new BufferedReader(new FileReader(copyDestination));
				String currentReadingLine = reader.readLine();
				while (currentReadingLine != null) {
					originalFileContent += currentReadingLine;
					originalFileContent += System.lineSeparator();
					currentReadingLine = reader.readLine();
				}

				String modifiedFileContent = originalFileContent.replaceFirst("[\n\r]+$", "");
				// modifiedFileContent =
				// modifiedFileContent.replaceAll("\\$DATE", EffectiveDate);

				if (!ApxGrpAddr1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ApxGrpAddr1", ApxGrpAddr1);

				if (!ApxGrpEmailId.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ApxGrpEmailId", ApxGrpEmailId);

				if (!ApxGrpNo.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ApxGrpNo", ApxGrpNo);

				if (!ApxGrpName.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ApxGrpName", ApxGrpName);

				if (!ApxSSPK.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$ApxSSPK", ApxSSPK);

				if (!SSRN.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SSRN", SSRN);

				if (!RFPN.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$RFPN", RFPN);

				if (!PolicyNumber1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$PolicyNumber1", PolicyNumber1);

				if (!PolicyNumber2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$PolicyNumber2", PolicyNumber2);

				if (!CSCI1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$CSCI1", CSCI1);

				if (!CSCI2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$CSCI2", CSCI2);

				if (!EntityCode1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EntityCode1", EntityCode1);

				if (!EntityCode2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EntityCode2", EntityCode2);

				if (!EntityCode3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EntityCode3", EntityCode3);

				if (!EntityCode4.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EntityCode4", EntityCode4);

				if (!EntityCode5.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$EntityCode5", EntityCode5);

				if (!SubGrpAddr1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SubGrpAddr1", SubGrpAddr1);

				if (!SubGrpAddr2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SubGrpAddr2", SubGrpAddr2);

				if (!SubGrpAddr3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SubGrpAddr3", SubGrpAddr3);

				if (!SubGrpSSEC1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SubGrpSSEC1", SubGrpSSEC1);

				if (!SubGrpSSEC2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SubGrpSSEC2", SubGrpSSEC2);

				if (!SubGrpSSEC3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SubGrpSSEC3", SubGrpSSEC3);

				if (!SubGrpName1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SubGrpName1", SubGrpName1);

				if (!SubGrpName2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SubGrpName2", SubGrpName2);

				if (!SubGrpName3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SubGrpName3", SubGrpName3);

				if (!SubGrpSSPK1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SubGrpSSPK1", SubGrpSSPK1);

				if (!SubGrpSSPK2.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SubGrpSSPK2", SubGrpSSPK2);

				if (!SubGrpSSPK3.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$SubGrpSSPK3", SubGrpSSPK3);

				writer = new BufferedWriter(new FileWriter(copyDestination));

				writer.write(modifiedFileContent);
			} catch (IOException e) {
				log.error(e.getMessage(), e);
				e.printStackTrace();
			} finally {

				try {
					if (reader != null) {
						reader.close();
					}
					if (writer != null) {
						writer.close();
					}
				} catch (IOException e) {
					log.error(e.getMessage(), e);
					e.printStackTrace();

				}
			}
		}
		;
	}

	public static void ReadTxtFile() throws IOException { // 6-16-2022
		String filepath = Config.inputDataFilePath.toString() + Local_Path + FileName;
		log.info(filepath);
		@SuppressWarnings("unused")
		int failCount = 0;
		int passCount = 0;
		String readLine;
		File f = new File(filepath);
		// StringBuffer sb = new StringBuffer();

		try (BufferedReader br = new BufferedReader(new FileReader(f))) {

			// FileReader fr= new FileReader(filepath);

			while ((readLine = br.readLine()) != null) {
				String line = readLine;
				if (FileName.contains("Majesco_ASSIGNMENTS_ATP")) {
					if (line.trim().equalsIgnoreCase(FEIN + "|" + NPN1 + "|")) {

						log.info("pass" + line);
						log.info("Match found for NPN" + NPN1);
						webDriver.getReport().setStatus("PASS");
						passCount = 1;
					} else if (FileName.contains("Majesco_INDIVIDUAL_ATP")) {

					}

				}

			}
			if (passCount < 1) {
				failCount = 1;
				//log.info();
				webDriver.getReport().setStatus("FAIL");
				log.info("Match Not found for NPN" + NPN1);
				log.info("Fail");
			}

		} catch (Exception e) {
			// TODO Auto-generated catch block
			log.error(e.getMessage(), e);
			e.printStackTrace();
		}
		// return passCount;
	}
}
