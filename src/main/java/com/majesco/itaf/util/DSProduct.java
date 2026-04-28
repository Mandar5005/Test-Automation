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
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Wait;
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

public class DSProduct {

	private final static Logger log = LogManager.getLogger(WebHelperUtil.class.getName());
	static MainController controller = ObjectFactory.getMainController();
	static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	public static String sqlQuery = "";
	public static String readFromColName = "";
	public static String toBeFetchedDBColName = "";
	public static String expectedDBStatus = "";
	public static String writeToColName = "";
	public static Wait<WebDriver> waitForElementPresence;
	public static Wait<WebDriver> dmgebtWait;
	public static String Local_Path = "";
	public static String Template_Path = "";
	public static String Remote_Path = "";
	public static String Remote_Path_Outbound = "";
	public static String EffectiveDate = "";
	public static String EffectiveDate1 = "";
	public static String RiskCommencementDate = "";
	public static String RiskCommencementDate1 = "";
	public static String PolicyEffectiveToDate = "";
	public static String ProposalApplDate = "";
	public static String ProposalSubmDate = "";
	public static String PolicyNumber = "";
	public static String PolicyNumber1 = "";
	public static String PolicyNumber2 = "";
	public static String PolicyNumber3 = "";
	public static String PolicyNumber4 = "";
	public static String PolicyNumber5 = "";
	public static String PolicyFileLoadNumber = "";
	public static String PremiumFileLoadNumber = "";
	public static String EntityCode1 = "";
	public static String EntityCode2 = "";
	public static String EntityCode3 = "";
	public static String EntityCode4 = "";
	public static String EntityCode5 = "";
	public static String PayoutHeaderID = "";
	public static String TemplateFileName = "";
	public static String FileNameToBeCreated = "";
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
	public static File copySource;
	public static File copyDestination;
	public static String source;
	public static String destination;
	public static String ApxGrpAddr1 = "";
	public static String ApxGrpEmailId = "";
	public static String ApxGrpNo = "";
	public static String SubGrpNo = "";
	public static String ApxGrpName = "";
	public static String ApxSSPK = "";
	public static String SSRN = "";
	public static String RFPN = "";
	public static String SubGrpAddr1 = "";
	public static String SubGrpAddr2 = "";
	public static String SubGrpAddr3 = "";
	public static String SubGrpSSEC1 = "";
	public static String SubGrpSSEC2 = "";
	public static String SubGrpSSEC3 = "";
	public static String SubGrpName1 = "";
	public static String SubGrpName2 = "";
	public static String SubGrpName3 = "";
	public static String SubGrpSSPK1 = "";
	public static String SubGrpSSPK2 = "";
	public static String SubGrpSSPK3 = "";

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

	public static void writeDataToUniqueNumberSheet(String ctrlValue, String columnName) throws Exception {
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

			Cell uniqueTestCaseID = uniqueRow.createCell(uniqueValuesHashMap.get("TestCaseID").intValue());
			Cell uniqueCell = uniqueRow.createCell(uniqueValuesHashMap.get(columnName).intValue());
			uniqueTestCaseID.setCellValue(controller.controllerTestCaseID.toString());
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
			if (arrTempVal.length <= 2) {
				if (tableID.contains("$value1") || tableID.contains("$value2")) {
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
		} else if (tableID.contains(" AND ")) {
			tableID = tableID.replace("$value1", uniqueNumber);
			tableID = tableID.replace("$value2", uniqueNumber1);
			tableID = tableID.replace("$value3", uniqueNumber2);
		} else {
			tableID = tableID.replace("$value", uniqueNumber);
		}

		return tableID;

	}

	public static String getDataFetchingSQLQuery(String sqlQuery, String ctrlValue) throws IOException {

		String tempVal = "";
		String finalVal = "";
		String[] arrCtrlValue;
		int i;
		int j;
		@SuppressWarnings("unused")
		int k;
		int n;

		if (ctrlValue.contains("*")) {
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
						} else {
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
						} else {
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
				} else {
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
				} else {
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
		if (strValue.contains("-") && !strValue.contains("$")) {
			String[] arrOfStr1 = tempstrValue.split("\\-");

			for (i = 0; i < arrOfStr1.length; i++) {
				String[] arrOfStr11 = arrOfStr1[i].split("\\|");
				tempstrValue = ReadFromExcelUsingColumnName(arrOfStr11[0].toString(), arrOfStr11[1].toString());
				strValue = strValue.replace(arrOfStr1[i], tempstrValue);
				strExptValue = strValue;
			}

		} else {
			String[] arrOfStr = tempstrValue.split("\\|");

			if (arrOfStr.length > 2) {

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
		HSSFFormulaEvaluator.evaluateAllFormulaCells(expectedSheet.getWorkbook());

		FormulaEvaluator evaluator = expectedSheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

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
		System.out.println("Trying to connect.....");

		try {
			session = jsch.getSession(userName, hostName, port);
			session.setConfig("StrictHostKeyChecking", "no");
			session.setPassword(password);
			session.connect();

			Channel channel = session.openChannel("sftp");
			channel.connect();
			ChannelSftp sftpChannel = (ChannelSftp) channel;

			System.out.println("Connection SHTP Established !!");

			Vector<ChannelSftp.LsEntry> filesList = sftpChannel.ls(copyToRemote);
			if (!String.valueOf(controller.controllerTestCaseID).equals("RGRS_SC_179")) {
				if ((!filesList.isEmpty() || filesList != null)) {

					for (LsEntry file : filesList) {
						System.out.println("In File Delete Loop");
						log.info("In File Delete Loop");
						if (!".".equals(file.getFilename()) && !"..".equals(file.getFilename())) {
							sftpChannel.rm(copyToRemote + "//" + file.getFilename());
							System.out.println("File has been deleted");
							log.info("File has been deleted");
						}
					}
				}
			}
			sftpChannel.lcd(copyFromLocal);

			// Copy flat from Local to remote linux server
			sftpChannel.cd(copyToRemote);

			sftpChannel.put(copyFromLocal + "\\" + fileName, copyToRemote);
			System.out.println("File has been copied to FTP -> Inbound folder.");
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
		System.out.println("IN : DM Recovery");
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

			System.out.println("OUT : DM Recovery");
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
		System.out.println("Trying to connect.....");

		source = Config.inputDataFilePath.toString() + Local_Path;
		destination = Remote_Path;

		try {
			// session = jsch.getSession(userName, hostName, 22);
			session = jsch.getSession(userName, hostName, port);
			session.setConfig("StrictHostKeyChecking", "no");
			session.setPassword(password);
			session.connect();

			Channel channel = session.openChannel("sftp");
			channel.connect();
			ChannelSftp sftpChannel = (ChannelSftp) channel;

			System.out.println("Connection SHTP Established !!");

			Vector<ChannelSftp.LsEntry> filesList = sftpChannel.ls(destination);

			// Condition for handling file placing for EOD Batch scenario
			if (!controller.controllerTestCaseID.toString().equals("RGRS_SC_179")) {
				if (!filesList.isEmpty() || filesList != null) {

					for (LsEntry file : filesList) {
						System.out.println("In File Delete Loop");
						log.info("In File Delete Loop");
						if (!".".equals(file.getFilename()) && !"..".equals(file.getFilename())) {
							sftpChannel.rm(destination + "//" + file.getFilename());
							System.out.println("File has been deleted");
							log.info("File has been deleted");
						}
					}
				}
			}
			sftpChannel.lcd(source);

			// Copy flat from Local to remote linux server
			sftpChannel.cd(destination);

			sftpChannel.put(source + "\\" + FileNameToBeCreated, destination);
			System.out.println("File has been copied to FTP -> Inbound folder.");
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

	public static void CreateFlatFile(String logicalName, String ctrlValue) throws IOException {
		// TODO Auto-generated method stub
		if (logicalName.equalsIgnoreCase("Local_Path") && !ctrlValue.equals(""))
			Local_Path = ctrlValue;
		else if (logicalName.equalsIgnoreCase("Template_Path") && !ctrlValue.equals(""))
			Template_Path = ctrlValue;
		else if (logicalName.equalsIgnoreCase("Remote_Path") && !ctrlValue.equals(""))
			Remote_Path = ctrlValue;
		else if (logicalName.equalsIgnoreCase("EffectiveDate") && !ctrlValue.equals(""))// Minaakshi : 01-10-2019
			if (ctrlValue.contains("$value")) {
				String strTemp = null;
				String strTemp1 = null;
				strTemp = DSProduct.ReadFromExcelUsingColumnName("", "SchemeEvaluationToDate");
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
		else if (logicalName.equalsIgnoreCase("RiskCommencementDate") && !ctrlValue.equals(""))
			if (ctrlValue.contains("$value")) {
				String strTemp = null;
				String strTemp1 = null;
				strTemp = DSProduct.ReadFromExcelUsingColumnName("", "SchemeEvaluationToDate");
				String arrTemp[] = strTemp.split("\\/");
				strTemp1 = arrTemp[2] + arrTemp[0] + arrTemp[1];
				ctrlValue = ctrlValue.replace("$value", strTemp1);
				RiskCommencementDate = ctrlValue;
			} else {
				RiskCommencementDate = ctrlValue;
			}
		else if (logicalName.equalsIgnoreCase("RiskCommencementDate1") && !ctrlValue.equals(""))
			RiskCommencementDate1 = ctrlValue;
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
		// Minaakshi : 01-05-2019
		else if (logicalName.equalsIgnoreCase("PolicyNumber") && !ctrlValue.equals(""))
			PolicyNumber = ctrlValue;
		// Minaakshi : 01-05-2019
		else if (logicalName.equalsIgnoreCase("PolicyFileLoadNumber") && !ctrlValue.equals(""))
			PolicyFileLoadNumber = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
		// Minaakshi : 01-05-2019
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
			PayoutHeaderID = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);// Minaakshi : 01-01-2020
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
		else if (logicalName.equalsIgnoreCase("SubGrpNo") && !ctrlValue.equals(""))// Minaakshi : 01-09-2020
			SubGrpNo = ReadFromExcelUsingColumnName(ctrlValue, WebHelper.columnName);
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

				if (!RiskCommencementDate1.equals(""))
					modifiedFileContent = modifiedFileContent.replaceAll("\\$RC_DATE1", RiskCommencementDate1);

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

				if (!StringUtils.equalsIgnoreCase(PolicyNumber, "")
						|| !StringUtils.equalsIgnoreCase(PolicyNumber, null))// Minaakshi
																				// :
																				// 01-05-2019
					modifiedFileContent = modifiedFileContent.replaceAll("\\$POL_NUMBER", PolicyNumber);

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
					modifiedFileContent = modifiedFileContent.replaceAll("\\$PHID", PayoutHeaderID);// Minaakshi :
																									// 01-01-2020

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

}