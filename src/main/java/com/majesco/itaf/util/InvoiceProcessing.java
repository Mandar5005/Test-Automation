package com.majesco.itaf.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;
import com.jcraft.jsch.JSchException;
import com.jcraft.jsch.Session;
import com.jcraft.jsch.SftpException;

public class InvoiceProcessing {
	private static final Logger logger = LogManager.getLogger(InvoiceProcessing.class);
	public final String FILE_EXTENSION;
	public final String REMOTE_PATH;
	public final String ARCHIVE_PATH;
	public final String DIR_PATH;
	public final String CONFIG_EXCEL_FILE_PATH;
	public final String MAIN_CONTROLLER_PATH;
	public final String TEST_CASE_DATE;
	public final String EXCEL_TAB_NAME;
	public final String HDR_LOOKUP;
	public final boolean IS_ENV_WINDOWS;
	public final String REPORT_EXCEL_FILE;
	private static final String LINUX_ENV_TYPE = "Linux";
	private static final int TEST_CASE_ID_COL = 1; // MAIN CONTROLLER TEST CASE ID COLUMN
	private static final int NO_OF_COL_PARSE = 10; // MAIN CONTROLLER COLUMNS
	private static final String PDF_PROCESSING = "PDF";
	private static final String PDF_EXTENSION = ".pdf";
	private static final String XML_EXTENSION = ".xml";
	private static final String MAIN_CONTROLLER_SHEET = "MainControlSheet";
	public static final String DATE_FORMAT = "MM_dd_yyyy";
	private static final String HDR_OUTPUT_FORM = "OUTPUT_FORM";
	private static final String HDR_PDF_LOOKUP = "PDF_Lookup";
	private static final String HDR_XML_LOOKUP = "XML_Lookup";
	private Map<String, String> labelMapping = new HashMap<>();
	private SFTPUtil sftpUtil = null;
	private static final String[] HEADER_LABELS = { "TestCaseID", "TransactionType", "CycleDate", "WebServiceName",
			"REQUEST_XML" };
	private static final String REPORT_EXCEL_SHEET_NAME = "Report";
	private static final String VERIFY_LABEL = "Verify";

	public InvoiceProcessing(String remotePath, String archivePath, String dirPath, String configExcelFilePath,
			String excelTabName, String fileProcessing, String mainControllerPath, String testCaseDate, String envType,
			String sftpUid, String sftpPwd, String sftpHost, int sftpPort, String reportExcelFile) {
		this.REMOTE_PATH = remotePath;
		this.ARCHIVE_PATH = archivePath;
		this.DIR_PATH = dirPath;
		this.CONFIG_EXCEL_FILE_PATH = configExcelFilePath;
		this.EXCEL_TAB_NAME = excelTabName;
		this.MAIN_CONTROLLER_PATH = mainControllerPath;
		this.TEST_CASE_DATE = testCaseDate;
		this.IS_ENV_WINDOWS = !envType.equalsIgnoreCase(LINUX_ENV_TYPE);
		this.REPORT_EXCEL_FILE = reportExcelFile;

		if (!IS_ENV_WINDOWS) {
			sftpUtil = new SFTPUtil(sftpUid, sftpPwd, sftpHost, sftpPort);
		}

		if (fileProcessing.equalsIgnoreCase(PDF_PROCESSING)) {
			FILE_EXTENSION = PDF_EXTENSION;
			HDR_LOOKUP = HDR_PDF_LOOKUP;
		} else {
			FILE_EXTENSION = XML_EXTENSION;
			HDR_LOOKUP = HDR_XML_LOOKUP;
		}
	}

	public Map<String, String> getLabelMapping() {
		return labelMapping;
	}

	public void setLabelMapping(Map<String, String> labelMapping) {
		this.labelMapping = labelMapping;
	}

	public boolean copyFiles(String source, String destination) throws IOException {
		File[] files = getFiles(source);
		for (File file : files) {
			Path path = Files.copy(Paths.get(file.getPath()), Paths.get(destination + file.getName()),
					StandardCopyOption.REPLACE_EXISTING);
			logger.info("File Copied : " + path);
		}

		return true;
	}

	public boolean copySFTPFiles(String remoteSource, String localDestination) throws JSchException, SftpException {
		List<String> fileNameList = getSFTPFiles(remoteSource);
		Session session = null;

		try {
			session = sftpUtil.getSftpSession();
			for (String fileName : fileNameList) {
				logger.info("Remote : " + remoteSource + fileName);
				logger.info("Local : " + localDestination + fileName);
				sftpUtil.downloadSFTPFile(session, remoteSource + fileName, localDestination + fileName);
				logger.info("File Downloaded : " + fileName);
			}
		} finally {
			if (null != session && session.isConnected())
				session.disconnect();
		}

		return true;
	}

	public boolean moveFiles(String source, String destination) throws IOException {
		File[] files = getFiles(source);
		for (File file : files) {
			Path path = Files.copy(Paths.get(file.getPath()), Paths.get(destination + file.getName()),
					StandardCopyOption.REPLACE_EXISTING);
			boolean isDelete = file.delete();
			logger.info("File Moved : " + isDelete + " " + path);
		}

		return true;
	}

	public boolean moveSFTPFiles(String remoteOldPath, String remoteNewPath) throws JSchException, SftpException {
		List<String> fileNameList = getSFTPFiles(remoteOldPath);
		Session session = null;

		try {
			session = sftpUtil.getSftpSession();
			for (String fileName : fileNameList) {
				logger.info("Remote Old : " + remoteOldPath + fileName);
				logger.info("Remote Arch New : " + remoteNewPath + fileName);
				sftpUtil.moveSFTPFile(session, remoteOldPath, remoteNewPath, fileName);
				logger.info("File Moved : " + fileName);
			}
		} finally {
			if (null != session && session.isConnected())
				session.disconnect();
		}

		return true;
	}

	public boolean copyAndArchivedFiles(String remotePath, String destPath, String archivingPath) throws IOException {
		File remoteDir = new File(remotePath);

		if (remoteDir.exists() && remoteDir.isDirectory()) {
			File destDir = new File(destPath);
			if (destDir.exists() && destDir.isDirectory()) {
				copyFiles(remotePath, destPath);

				File archiveDir = new File(archivingPath);
				if (archiveDir.exists() && archiveDir.isDirectory()) {
					moveFiles(remotePath, archivingPath);
				} else {
					logger.error("Archive path does not exist : " + archivingPath);
					return false;
				}
			} else {
				logger.error("Destination path does not exist : " + destPath);
				return false;
			}
		} else {
			logger.error("Remote path does not exist: " + remotePath);
			return false;
		}

		return true;
	}

	public boolean copyAndArchivedSFTPFiles(String remoteSource, String localDestination, String remoteArchivingPath)
			throws JSchException, SftpException {
		boolean isCopied = copySFTPFiles(remoteSource, localDestination);
		boolean isCopiedAndArchived = false;

		if (isCopied) {
			isCopiedAndArchived = moveSFTPFiles(remoteSource, remoteArchivingPath);
		}

		return isCopiedAndArchived;
	}

	public List<String> getSFTPFiles(String directoryPath) throws JSchException, SftpException {
		Session session = null;

		try {
			session = sftpUtil.getSftpSession();
			return sftpUtil.getSFTPFiles(session, directoryPath, FILE_EXTENSION);
		} finally {
			if (null != session && session.isConnected())
				session.disconnect();
		}
	}

	public File[] getFiles(String directoryPath) {
		File dir = new File(directoryPath);
		if (!dir.exists() || !dir.isDirectory()) {
			logger.error("Directory for PDF files does not exist : " + directoryPath);
			return new File[0];
		}

		return dir.listFiles((d, s) -> s.toLowerCase().endsWith(FILE_EXTENSION));
	}

	public String getPDFLabelText(File file, String label) throws IOException {
		try (PDDocument document = PDDocument.load(file)) {
			document.getClass();
			if (!document.isEncrypted()) {
				PDFTextStripperByArea stripper = new PDFTextStripperByArea();
				stripper.setSortByPosition(true);
				PDFTextStripper tStripper = new PDFTextStripper();
				String pdfFileInText = tStripper.getText(document);
				String[] lines = pdfFileInText.split("\\r?\\n");// split by whitespace

				for (String line : lines) {
					if (line.contains(label)) {
						String[] text = line.split(label);
						if (text.length > 0) {
							return text[1].trim();
						}
					}
				}
			}
		}

		return null;
	}

	public String getXMLLabelText(File file, String label) throws IOException {
		try {
			DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
			Document doc = dBuilder.parse(file);
			doc.getDocumentElement().normalize();
			NodeList nodes = doc.getElementsByTagName(label);
			Element line = (Element) nodes.item(0);
			return line.getFirstChild().getTextContent();
		} catch (ParserConfigurationException pce) {
			logger.error("Parse exception : " + file.getName() + " " + pce.getMessage(), pce);
			return null;
		} catch (SAXException sax) {
			logger.error("SAX exception : " + file.getName() + " " + sax.getMessage(), sax);
			return null;
		}
	}

	public static String getDate(String cycleDateCellValue, String dateFormat) throws ParseException {
		String date;

		DateFormat formatter = new SimpleDateFormat("MM/dd/yyyy");
		String dateString = cycleDateCellValue.replaceAll("\"", "");
		Date cellCycleDate = formatter.parse(dateString);
		formatter = new SimpleDateFormat(dateFormat);
		date = formatter.format(cellCycleDate);
		logger.info("date : " + date);

		return date;
	}

	public void processFile(File file, String reportExcelPath, String dirPath, String extension,
			String mainControllerFilePath, String testCaseDate) throws IOException, InvalidFormatException {
		boolean isFileNameFound = false;
		for (Map.Entry<String, String> entry : getLabelMapping().entrySet()) {
			String key = entry.getKey();
			if (file.getName().contains(key)) {
				isFileNameFound = true;
				String label = entry.getValue();
				String labelText;

				if (HDR_LOOKUP.contains(PDF_PROCESSING))
					labelText = getPDFLabelText(file, label);
				else
					labelText = getXMLLabelText(file, label);

				if (StringUtils.isNotBlank(labelText)) {
					logger.info("File : " + file.getName() + " Label Text : " + labelText);

					String testCaseId = getTestCaseId(mainControllerFilePath, labelText);
					if (StringUtils.isNotBlank(testCaseId)) {
						StringBuilder newFileName = new StringBuilder();
						newFileName.append(key);
						newFileName.append("_");
						newFileName.append(testCaseId);
						newFileName.append("_");
						newFileName.append(testCaseDate);

						boolean isFileRename = renameFile(file, dirPath, extension, newFileName.toString());
						logger.info("isFileRename : " + isFileRename);

						if (isFileRename) {
							addFileNameInReportExcel(reportExcelPath, testCaseId, testCaseDate, newFileName.toString(),
									key);
						} else {
							for (int newFilecount = 1; newFilecount <= 100;) {
								String newFileName1 = dirPath + File.separator + newFileName + "_" + newFilecount
										+ "_Response" + extension;
								File fileExist = new File(newFileName1);
								if (fileExist.exists()) {
									newFilecount = newFilecount + 1;
								} else {
									// renaming new file
									file.renameTo(fileExist);
									newFileName.append("_");
									newFileName.append(newFilecount);
									newFileName.append("_");
									addFileNameInReportExcel(reportExcelPath, testCaseId, testCaseDate,
											newFileName.toString(), key);
									break;
								}
							}
						}
					} else {
						logger.error("Test case id does not exist for label : " + labelText);
					}
				} else {
					logger.error("Label " + label + " does not exist in file: " + file.getName());
				}
			}

			if (isFileNameFound)
				break;
		}

		if (!isFileNameFound) {
			logger.error("File name mapping does not exist : " + file.getName());
		}
	}

	public boolean renameFile(File file, String dirPath, String extension, String newFileName) {
		logger.info("File renaming done : " + dirPath + File.separator + newFileName + "_Response" + extension);
		return file.renameTo(new File(dirPath + File.separator + newFileName + "_Response" + extension));

	}

	public String getTestCaseId(String mainControllerFilePath, String labelToBeRetrieved) throws IOException {
		try (XSSFWorkbook excelBook = new XSSFWorkbook(new FileInputStream(new File(mainControllerFilePath)))) {
			XSSFSheet excelSheet = excelBook.getSheet(MAIN_CONTROLLER_SHEET);
			for (int row = 0; row < excelSheet.getLastRowNum() + 1; row++) {
				XSSFRow excelRow = excelSheet.getRow(row);
				for (int col = 0; col < NO_OF_COL_PARSE; col++) {

					// if (null != excelRow.getCell(col) && excelRow.getCell(col).getCellType() ==
					// CellType.STRING//Sel4
					if (null != excelRow.getCell(col) && excelRow.getCell(col).getCellType() == CellType.STRING
							&& StringUtils.isNotBlank(excelRow.getCell(col).getStringCellValue())
							&& excelRow.getCell(col).toString().equalsIgnoreCase(labelToBeRetrieved)) {
						return excelRow.getCell(TEST_CASE_ID_COL).toString();
					}
				}
			}
		}

		return null;
	}

	public void readConfigExcel(String configExcelFilePath, String tabName) throws IOException {
		try (XSSFWorkbook excelBook = new XSSFWorkbook(new FileInputStream(new File(configExcelFilePath)))) {
			XSSFSheet excelSheet = excelBook.getSheet(tabName);
			int hdrOutputFormCol = 0;
			int hdrPDFLookupCol = 0;
			for (int row = 0; row < excelSheet.getLastRowNum() + 1; row++) {
				if (row == 0) {// READ HEADERS
					hdrOutputFormCol = getLabelMappingIndex(excelSheet.getRow(row), HDR_OUTPUT_FORM);
					hdrPDFLookupCol = getLabelMappingIndex(excelSheet.getRow(row), HDR_LOOKUP);
				}

				XSSFRow excelRow = excelSheet.getRow(row);// Sel4
				/*
				 * if (excelRow.getCell(hdrOutputFormCol).getCellType() == CellType.STRING //
				 * FILE NAME && excelRow.getCell(hdrPDFLookupCol).getCellType() ==
				 * CellType.STRING // PDF LABEL &&
				 * StringUtils.isNotBlank(excelRow.getCell(hdrOutputFormCol).getStringCellValue(
				 * )) &&
				 * StringUtils.isNotBlank(excelRow.getCell(hdrPDFLookupCol).getStringCellValue()
				 * ))
				 */

				if (excelRow.getCell(hdrOutputFormCol).getCellType() == CellType.STRING // FILE NAME
						&& excelRow.getCell(hdrPDFLookupCol).getCellType() == CellType.STRING // PDF LABEL
						&& StringUtils.isNotBlank(excelRow.getCell(hdrOutputFormCol).getStringCellValue())
						&& StringUtils.isNotBlank(excelRow.getCell(hdrPDFLookupCol).getStringCellValue())) {
					String fileName = excelRow.getCell(hdrOutputFormCol).getStringCellValue();
					String pdfLabel = excelRow.getCell(hdrPDFLookupCol).getStringCellValue();

					getLabelMapping().put(fileName, pdfLabel);
				}
			}
		}
	}

	public void movedProcessedFiles(String dirPath, String testCaseDate) throws IOException {
		boolean dirExist;
		File dir = new File(dirPath + testCaseDate);
		logger.info("movedProcessedFiles : " + dir.getAbsolutePath());
		if (!dir.exists() || !dir.isDirectory()) {
			dirExist = dir.mkdir();
		} else
			dirExist = true;

		if (dirExist) {
			moveFiles(dirPath, dir.getAbsolutePath() + File.separator);
			logger.info("Files moved to respective date folder");
		}
	}

	public int getLabelMappingIndex(XSSFRow excelRow, String label) {
		for (int col = 0; col < excelRow.getLastCellNum(); col++) {
			if (null != excelRow.getCell(col) && StringUtils.isNotBlank(excelRow.getCell(col).getStringCellValue())
					&& excelRow.getCell(col).getStringCellValue().equals(label)) {
				return col;
			}
		}

		logger.error("Label does not exist : " + label);

		return -1;
	}

	public boolean isReportExcelFileExist(String filePath) throws IOException {
		File file = new File(filePath);

		if (!file.exists() || !file.isFile()) {
			createExcelHeaders(filePath);
			return true;
		}

		return file.exists();
	}

	public void addFileNameInReportExcel(String filePath, String testCaseId, String testCaseDate, String fileName,
			String key) throws InvalidFormatException, IOException {
		try (InputStream is = new FileInputStream(new File(filePath))) {
			try (Workbook workbook = WorkbookFactory.create(is)) {
				Sheet sheet = workbook.getSheet(InvoiceProcessing.REPORT_EXCEL_SHEET_NAME);
				int rowNum = sheet.getLastRowNum() + 1;
				Row row = sheet.createRow(rowNum);
				row.createCell(0).setCellValue(testCaseId);
				row.createCell(1).setCellValue(InvoiceProcessing.VERIFY_LABEL);
				row.createCell(2).setCellValue(testCaseDate);
				row.createCell(3).setCellValue(key);
				row.createCell(4).setCellValue(fileName);
				try (FileOutputStream fos = new FileOutputStream(filePath)) {
					workbook.write(fos);
				}
			}
		}
	}

	public void createExcelHeaders(String filePath) throws IOException {
		try (Workbook workbook = new XSSFWorkbook()) {
			Sheet sheet = workbook.createSheet(InvoiceProcessing.REPORT_EXCEL_SHEET_NAME);
			Row row = sheet.createRow(0);
			for (int i = 0; i < InvoiceProcessing.HEADER_LABELS.length; i++) {
				Cell header = row.createCell(i);
				header.setCellValue(InvoiceProcessing.HEADER_LABELS[i]);
			}
			sheet.autoSizeColumn(1);
			try (FileOutputStream fos = new FileOutputStream(filePath)) {
				workbook.write(fos);
			}
		}
	}

	public static void main(String[] args) throws Exception {

		InvoiceProcessing invoiceProcessing = new InvoiceProcessing(
				"D:\\Sukrut\\PDF_Read_Rename_Files_By_Sonali\\RemotePath\\", // "/data/IBM/Product-Group/Reports/",
				"D:\\Sukrut\\PDF_Read_Rename_Files_By_Sonali\\RemotePath\\ArchivePath\\", // "/data/IBM/Product-Group/Reports/Archive/",
				"D:\\Sukrut\\PDF_Read_Rename_Files_By_Sonali\\LocalPath\\",
				"D:\\Sukrut\\PDF_Read_Rename_Files_By_Sonali\\OutputConfiguration_new.xlsx", "Config", "PDF", // "XML",
				"D:\\Sukrut\\Sharing_files_billing_automation_suite\\iTAFCommon\\RegressionMaster\\CommonResources\\MainController.xlsx",
				"11/14/2019", "Windows", // "Linux",
				"webapp", "webapp", "172.17.204.9", 22, "D:\\Sukrut\\PDF_Read_Rename_Files_By_Sonali\\OutputXML.xlsx");

		invoiceProcessing.readConfigExcel(invoiceProcessing.CONFIG_EXCEL_FILE_PATH, invoiceProcessing.EXCEL_TAB_NAME);
		logger.info(invoiceProcessing.getLabelMapping());

		boolean isCopyAndArchived;

		if (invoiceProcessing.IS_ENV_WINDOWS) {
			isCopyAndArchived = invoiceProcessing.copyAndArchivedFiles(invoiceProcessing.REMOTE_PATH,
					invoiceProcessing.DIR_PATH, invoiceProcessing.ARCHIVE_PATH);
		} else {
			isCopyAndArchived = invoiceProcessing.copyAndArchivedSFTPFiles(invoiceProcessing.REMOTE_PATH,
					invoiceProcessing.DIR_PATH, invoiceProcessing.ARCHIVE_PATH);
		}

		logger.info("isCopyAndArchived : " + isCopyAndArchived);
		if (isCopyAndArchived) {
			boolean isFileExist = invoiceProcessing.isReportExcelFileExist(invoiceProcessing.REPORT_EXCEL_FILE);
			logger.info("Report excel file exist : " + isFileExist);

			File[] files = invoiceProcessing.getFiles(invoiceProcessing.DIR_PATH);
			if (null != files && files.length > 0) {
				for (File file : files) {
					logger.info("file.getName() : " + file.getName());
					invoiceProcessing.processFile(file, invoiceProcessing.REPORT_EXCEL_FILE, invoiceProcessing.DIR_PATH,
							invoiceProcessing.FILE_EXTENSION, invoiceProcessing.MAIN_CONTROLLER_PATH,
							getDate(invoiceProcessing.TEST_CASE_DATE, DATE_FORMAT));
				}

				invoiceProcessing.movedProcessedFiles(invoiceProcessing.DIR_PATH,
						getDate(invoiceProcessing.TEST_CASE_DATE, DATE_FORMAT));
			} else {
				logger.error("Files does not exist to process : " + invoiceProcessing.DIR_PATH);
			}
		}
	}

}
