package com.majesco.itaf.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Properties;
import java.util.Vector;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.core.config.Configurator;
import org.apache.logging.log4j.core.config.DefaultConfiguration;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import com.jcraft.jsch.Channel;
import com.jcraft.jsch.ChannelSftp;
import com.jcraft.jsch.ChannelSftp.LsEntry;
import com.jcraft.jsch.JSch;
import com.jcraft.jsch.JSchException;
import com.jcraft.jsch.Session;
import com.jcraft.jsch.SftpException;
import com.majesco.itaf.main.Config;

/**
 * This utility is to transfer output XML from server to local machine Version
 */
public class GenerateOutputXML {

	private final static Logger log = LogManager.getLogger(GenerateOutputXML.class.getName());
	private static String OutputCompUtility = null;
	public static HashMap<String, Object> LookupHashMap = new HashMap<String, Object>();
	public static String strHostName = null;
	public static String strUserName = null;
	public static String strPassword = null;
	public static String strCopyFrom = null;
	public static String strCopyTo = null;
	public static String strCopyToMainDestination = null;
	public static String OutputXmlFilePath = null;

	public static List<String> lstArrayList = new ArrayList<String>();

	public static void main(String[] args) throws IOException {
		InputStream input1 = null;
		Properties prop = new Properties();
		input1 = new FileInputStream("config.properties");
		prop.load(input1);
		String filePathFileName = prop.getProperty("filePathFileName");

		try {
			Configurator.initialize(new DefaultConfiguration());
			;
			log.info("########################Generate Output XML Version 1.0.0.1##################################");
			GetLookupValues(filePathFileName);
			ReadConfigFile();
			CopyOutputXMLFiles(strHostName, strUserName, strPassword, strCopyFrom, strCopyTo, strCopyToMainDestination);
			WriteToExcel();

		} catch (Exception e) {
			log.info("Exception in main method " + e.getMessage());
		}

	}

	/**
	 * This method is to write all file name into excel.
	 * 
	 * @throws IOException
	 */
	private static void WriteToExcel() throws IOException {
		String strFileName = null;
		InputStream input2 = null;
		Properties prop1 = new Properties();
		input2 = new FileInputStream("config.properties");
		prop1.load(input2);
		String OutputXmlFilePath = prop1.getProperty("OutputXmlFilePath");

		try {
			FileInputStream file = new FileInputStream(OutputXmlFilePath);
			try (XSSFWorkbook workbook = new XSSFWorkbook(file)) {
				XSSFSheet sheet = workbook.getSheetAt(0);
				@SuppressWarnings("unused")
				Cell cell = null;
				file.close();
				int nSize = lstArrayList.size();
				for (int i = 0; i < nSize; i++) {
					strFileName = lstArrayList.get(i).toString();
					Row row = sheet.createRow(i + 1);
					Cell cell1 = row.createCell(6);
					cell1.setCellValue(strFileName);
					Cell cell2 = row.createCell(0);
					cell2.setCellValue("OutputXML_" + i);
					Cell cell3 = row.createCell(1);
					cell3.setCellValue("Verify");
					Cell cell4 = row.createCell(2);
					cell4.setCellValue("Y");
					Cell cell5 = row.createCell(3);
					cell5.setCellValue("02_02_2018");
					// file.close();
				}
				try (FileOutputStream outputStream = new FileOutputStream(OutputXmlFilePath)) {
					workbook.write(outputStream);

					outputStream.close();

				}
			}
			log.info("Execution Completed............................");
			File webServiceUtilityFile = new File(OutputCompUtility);
			if (webServiceUtilityFile.exists()) {
				log.info(OutputCompUtility + " macro is running...");
				Jacob.main(OutputCompUtility, "BillXMLVerification");
				log.info(OutputCompUtility + "Macro execution completed.....");
			}
			System.exit(0);

		} catch (Exception e) {
			log.info("Exception while writing file name to excel " + e.getMessage());
			try {

			} catch (Exception e1) {
			}

		} finally {
			strFileName = null;
		}
	}

	/**
	 * This function is to read basic configuration
	 */
	private static void ReadConfigFile() {
		// TODO Auto-generated method stub
		@SuppressWarnings("unused")
		String projectPath = null;
		InputStream input = null;

		try {
			Properties prop = new Properties();
			input = new FileInputStream("config.properties");
			prop.load(input);
			strHostName = Config.flatFileHostName;
			strUserName = Config.flatFileUserName;
			strPassword = Config.flatFilePassword;
			if (strHostName == null) {
				strHostName = prop.getProperty("HostName");
				strUserName = prop.getProperty("UserName");
				strPassword = prop.getProperty("Password");
			}

			strCopyFrom = prop.getProperty("CopyFrom");
			strCopyTo = prop.getProperty("CopyTo");
			strCopyToMainDestination = prop.getProperty("CopyToMainDestination");
			OutputCompUtility = prop.getProperty("OutputCompUtility");

		} catch (Exception e) {
			log.info("Exception in class" + e.getClass() + " of method ReadConfigFile" + e.getMessage());
		} finally {
			projectPath = null;
		}
	}

	/**
	 * 
	 * @param filePathFileName
	 */
	private static void GetLookupValues(String filePathFileName) {
		XSSFRow rowActual1 = null;
		String projectPath = null;
		String configPath = null;
		InputStream myXls = null;
		try {
			log.info("Read values from Configuration file for lookup");
			projectPath = System.getProperty("user.dir");
			configPath = projectPath + "\\" + filePathFileName;
			myXls = new FileInputStream(configPath);
			try (XSSFWorkbook wBook = new XSSFWorkbook(myXls)) {
				XSSFSheet workSheet = wBook.getSheet("Config");
				DataFormatter format = new DataFormatter();
				int rowCount = workSheet.getLastRowNum() + 1;

				for (int rowIndex = 1; rowIndex < rowCount; rowIndex++) {
					rowActual1 = workSheet.getRow(rowIndex);
					String parameterName1 = format.formatCellValue(rowActual1.getCell(0));
					String value1 = format.formatCellValue(rowActual1.getCell(1));
					LookupHashMap.put(parameterName1, value1);
				}
			}

		} catch (Exception e) {
			System.out.println("Exception in GetLookupValues=" + e);
			log.info("Exception in GetLookupValues=" + e.getMessage());

		} finally {
			rowActual1 = null;
			projectPath = null;
			configPath = null;
			myXls = null;
		}
	}

	private static void CopyOutputXMLFiles(String strHostName, String strUserName, String strPwd, String strCopyFrom,
			String strDestination, String strMainDestination) {

		String strResponseFilename = null;
		String strFinalFileName = null;
		String strFileName = null;

		try {

			JSch jsch = new JSch();

			Session session = null;
			System.out.println("Trying to connect.....");
			session = jsch.getSession(strUserName, strHostName, 22);
			session.setConfig("StrictHostKeyChecking", "no");
			session.setPassword(strPwd);
			session.connect();
			Channel channel = session.openChannel("sftp");
			channel.connect();
			ChannelSftp sftpChannel = (ChannelSftp) channel;
			System.out.println("Done !!");
			sftpChannel.cd(strCopyFrom);
			@SuppressWarnings("rawtypes")
			Vector filelist = sftpChannel.ls(strCopyFrom);

			for (int i = 0; i < filelist.size(); i++) {
				// for (int i = 0; i < 5; i++) {
				LsEntry entry = (LsEntry) filelist.get(i);
				Boolean b = entry.getFilename().endsWith(".XML");
				if (b == true) {
					strFileName = entry.getFilename();
					strResponseFilename = strDestination + "/" + strFileName;
					sftpChannel.get(strFileName, strDestination);
					String strRenamedFilename = getUniqueFileName(strFileName, strResponseFilename);
					strFinalFileName = strFileName.substring(0, strFileName.indexOf("2") + 9) + strRenamedFilename
							+ "_Verify";
					WriteFileNameToExcel(strFinalFileName);
					strFinalFileName = strFinalFileName + "_Response.xml";
					strFinalFileName = strMainDestination + "/" + strFinalFileName;
					// System.out.println("strFinalFileName"+strFinalFileName);
					File objFile = null;
					File objNewFile = null;
					try {
						objFile = new File(strResponseFilename);
						objNewFile = new File(strFinalFileName);
						Thread.sleep(500);
						if (objNewFile.exists() && !objNewFile.isDirectory()) {
							strFinalFileName = strFinalFileName.replace("_Response.xml", "") + "_" + i
									+ "_Response.xml";
							objNewFile = new File(strFinalFileName);
							FileUtils.copyFile(objFile, objNewFile);
						} else {
							FileUtils.copyFile(objFile, objNewFile);
						}
					} catch (Exception e) {
						System.out.println("Error in file rename" + e);
					} finally {
						objFile = null;
						objNewFile = null;
						strResponseFilename = null;
						strFinalFileName = null;
						strFileName = null;
					}
				}
			}

		} catch (JSchException e) {
			e.printStackTrace();
		} catch (SftpException e) {
			e.printStackTrace();
		} catch (Exception e) {
			System.out.println("Exception in CopyOutputXMLFiles" + e);
		}

	}

	/**
	 * This method is to write file name into excel file for comparison
	 * 
	 * @param strFinalFileName
	 */
	private static void WriteFileNameToExcel(String strFinalFileName) {
		try {
			log.info("Inserting file name into list " + strFinalFileName);
			lstArrayList.add(strFinalFileName);
		} catch (Exception e) {
			log.info("Error while maintaining list of file name");
		}

	}

	private static String getUniqueFileName(String strFileName, String strFilePath) {
		// TODO Auto-generated method stub
		String strLookupFile = null;
		String strName = null;
		String strTagValue = "";

		try {
			strName = strFileName.substring(0, strFileName.indexOf("2") - 1);
			strLookupFile = LookupHashMap.get(strName).toString();
			strTagValue = getXMLTagValue(strLookupFile, strFilePath);
			strTagValue = strTagValue.substring(0, strTagValue.length() - 6);
			if (strName.contains("DIRECT_INVOICE")) {
				strTagValue = strTagValue + "_" + getXMLTagValue("EntityType", strFilePath);
			}

		} catch (Exception e) {
			System.out.println("Exception in getUniqueFileName=" + e);
		}
		return strTagValue;
	}

	@SuppressWarnings("finally")
	public static String getXMLTagValue(String xmlTagName, String strXMLFile) throws IOException, Exception {
		String tagValue = null;
		File fXmlFile = new File(strXMLFile);

		try {
			DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
			Document doc = dBuilder.parse(fXmlFile);
			doc.getDocumentElement().normalize();

			// First matching node
			Node firstNode = doc.getElementsByTagName(xmlTagName).item(0);
			tagValue = firstNode.getTextContent().toString();
			// return tagValue;
		} catch (IOException ioe) {

			throw new IOException(
					"Failed to get tag value " + xmlTagName + "  <-|-> LocalizeMessage " + ioe.getLocalizedMessage()
							+ " <-|-> Message" + ioe.getMessage() + " <-|-> Cause " + ioe.getCause());
		} catch (Exception e) {
			// throw new Exception("Error while XML tag verification: " +
			// e.getMessage());

			throw new Exception("Failed to read process list  <-|-> LocalizeMessage " + e.getLocalizedMessage()
					+ " <-|-> Message" + e.getMessage() + " <-|-> Cause " + e.getCause());
		} finally {
			return tagValue;
		}
	}
}
