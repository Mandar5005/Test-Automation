package com.majesco.itaf.util;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReconReport {

	private final static Logger log = LogManager.getLogger(ReconReport.class);
	private static final String RECON_REPORT_SHEET_NAME = "Reports Recon";
	private static final String RECON_REPORT_MISMATCH_TEXT = "MisMatch";

	private ReconReport() {
		throw new IllegalStateException("ReconReport class");
	}

	public static void main(String[] args) {
		log.info("inside main");
	}

	public static boolean generateReconReport(String cycleDateCellValue, final String JAR_FILE_PATH,
			final String JAR_FILE_NAME, final String REPORT_GENERATED_FILE_PATH, final String REPORT_ARCHIVAL_PATH,
			final String REPORT_DATE_FORMAT, final String REPORT_ARGS_1, final String REPORT_ARGS_2) {
		try {
			if (!isConfigParametersValid(JAR_FILE_PATH, JAR_FILE_NAME, REPORT_GENERATED_FILE_PATH, REPORT_ARCHIVAL_PATH,
					REPORT_DATE_FORMAT, REPORT_ARGS_1, REPORT_ARGS_2)) {
				log.error("some of the config parameters missing for Recon Report");
				return false;
			}

			File file = new File(JAR_FILE_PATH + File.separator + JAR_FILE_NAME);
			if (!file.exists() || !file.isFile()) {
				log.error("jar file does not exist for Recon_Report");
				return false;
			}

			File directory = new File(REPORT_GENERATED_FILE_PATH);
			if (!directory.exists() || !directory.isDirectory()) {
				log.error("Recon Report directory does not exist");
				return false;
			}

			boolean isArchived = archiveExistingFiles(REPORT_GENERATED_FILE_PATH, REPORT_ARCHIVAL_PATH);
			if (!isArchived) {
				log.error("Error in archiving existing Recon Reports");
				return false;
			}

			String yearMonth = getDate(cycleDateCellValue, REPORT_DATE_FORMAT);
			boolean isExecuted = executeJarFile(yearMonth, file, REPORT_ARGS_1, REPORT_ARGS_2);
			if (!isExecuted)
				return false;

			String[] files = directory.list((File dir, String name) -> {
				boolean flag = false;

				if (name.toLowerCase().startsWith(yearMonth + "_")
						&& (name.toLowerCase().endsWith(".xls") || name.toLowerCase().endsWith(".xlsx"))
						&& (name.length() == 25 || name.length() == 26))
					flag = true;

				return flag;
			});

			if (files.length > 1) {
				log.error("Multiple report files exist at : " + REPORT_GENERATED_FILE_PATH);
				return false;
			}

			return readExcel(REPORT_GENERATED_FILE_PATH + files[0]);
		} catch (Exception ex) {
			log.error("Error occurred in Recon Report : " + ex.getMessage(), ex);
		}

		return false;
	}

	private static String getDate(String cycleDateCellValue, final String REPORT_DATE_FORMAT) throws ParseException {
		String yearMonth;

		DateFormat formatter = new SimpleDateFormat("MM/dd/yyyy");
		String dateString = cycleDateCellValue.replaceAll("\"", "");
		Date cellCycleDate = formatter.parse(dateString);
		formatter = new SimpleDateFormat(REPORT_DATE_FORMAT);
		yearMonth = formatter.format(cellCycleDate);
		log.info("Recon_Report for date : " + yearMonth);

		return yearMonth;
	}

	private static boolean executeJarFile(String yearMonth, File file, final String REPORT_ARGS_1,
			final String REPORT_ARGS_2) throws IOException, InterruptedException {
		int completed;
		String output;
		String errorString;
		String absoluteFilePath = file.getAbsolutePath();

		log.info("Recon Jar absolutePath : " + absoluteFilePath);

		ProcessBuilder processBuilder = new ProcessBuilder("java", "-jar", absoluteFilePath, yearMonth, REPORT_ARGS_1,
				REPORT_ARGS_2);
		processBuilder.redirectErrorStream(true);

		Process process = processBuilder.start();

		BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));
		BufferedReader error = new BufferedReader(new InputStreamReader(process.getErrorStream()));

		while ((output = reader.readLine()) != null) {
			log.info(output);
		}

		while ((errorString = error.readLine()) != null) {
			log.error(errorString);
		}

		completed = process.waitFor();

		if (completed != 0) {
			log.error("Error in the execution of JAR file");
			return false;
		}

		return true;
	}

	private static boolean isConfigParametersValid(final String JAR_FILE_PATH, final String JAR_FILE_NAME,
			final String REPORT_GENERATED_FILE_PATH, final String REPORT_ARCHIVAL_PATH, final String REPORT_DATE_FORMAT,
			final String REPORT_ARGS_1, final String REPORT_ARGS_2) {
		boolean isValidate = false;

		if (StringUtils.isNotBlank(JAR_FILE_NAME) && StringUtils.isNotBlank(JAR_FILE_PATH)
				&& StringUtils.isNotBlank(REPORT_GENERATED_FILE_PATH) && StringUtils.isNotBlank(REPORT_DATE_FORMAT)
				&& StringUtils.isNotBlank(REPORT_ARGS_1) && StringUtils.isNotBlank(REPORT_ARGS_2)
				&& StringUtils.isNotBlank(REPORT_ARCHIVAL_PATH)) {
			isValidate = true;
		}

		return isValidate;
	}

	private static boolean readExcel(String file) throws IOException {
		boolean flag = false;

		try (XSSFWorkbook excelBook = new XSSFWorkbook(new FileInputStream(file))) {
			XSSFSheet excelSheet = excelBook.getSheet(RECON_REPORT_SHEET_NAME);
			log.info("No of Rows in generated excel sheet : " + excelSheet.getLastRowNum());
			for (int row = excelSheet.getFirstRowNum(); row <= excelSheet.getLastRowNum(); row++) {
				XSSFRow excelRow = excelSheet.getRow(row);
				flag = true;
				log.debug("Row : " + excelRow);

				if (null != excelRow && excelRow.getLastCellNum() > 5) {
					XSSFCell excelCell = excelRow.getCell(5);
					log.debug("Cell : " + excelCell.getStringCellValue());

					if (null != excelCell.getStringCellValue()
							&& excelCell.getStringCellValue().equalsIgnoreCase(RECON_REPORT_MISMATCH_TEXT)) {
						log.info("There is a 'mismatch' in recon report : " + file);
						return false;
					}
				}
			}

			if (flag)
				return true;
			else
				log.error("Record does not exist in the recon file");
		}

		return false;
	}

	private static boolean archiveExistingFiles(String existingFilePath, String archivingPath) throws IOException {
		File directory = new File(existingFilePath);

		String[] files = directory.list((File dir, String name) -> {
			boolean flag = false;
			if (name.toLowerCase().endsWith(".xls") || name.toLowerCase().endsWith(".xlsx")) {
				flag = true;
			}
			return flag;
		});

		File archiveDirectory = new File(archivingPath);
		if ((!archiveDirectory.exists()) || (!archiveDirectory.isDirectory())) {
			archiveDirectory.mkdir();
		}

		boolean isArchive = true;
		for (String fileName : files) {
			Path temp = Files.move(Paths.get(new File(existingFilePath + File.separator + fileName).toURI()),
					Paths.get(new File(archivingPath + File.separator + fileName).toURI()),
					java.nio.file.StandardCopyOption.REPLACE_EXISTING);

			if (null != temp) {
				log.info("Recon file archived : " + fileName);
			} else {
				log.error("Error in archiving recon file : " + fileName);
				isArchive = false;
			}
		}

		return isArchive;
	}
}
