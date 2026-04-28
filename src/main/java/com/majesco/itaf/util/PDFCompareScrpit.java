package com.majesco.itaf.util;

import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import com.majesco.compare.CompareUtil;
import com.majesco.itaf.main.Config;

public class PDFCompareScrpit {
	private final static Logger log = LogManager.getLogger(PDFCompareScrpit.class.getName());
	public static String businessDate = null;
	public static String expectedFile_Path = "";
	public static String actualFile_Path = "";
	public static String email = "";
	public static String defaultLocation1 = "";
	public static String defaultLocation2 = "";
	public static String expectedLocation = "";
	public static String actualLocation = "";
	public static String BCFolderValue = "";
	public static String FileNameDate = "";
	public static String agencyOrPolicy = "";
	public static String billingType = "";
	public static WebDriver driver;
	public static String pdfBackupLocation = "";
	public static String agentCodePDF = "";
	public static String[] pdfParams = new String[5];
	public static String[] pdfCompare = new String[2];
	WebElement tableElement = null;
	List<WebElement> rowElements = null;

	public static void fileRename(final String pdfCompareParam, final String TestCaseID)

	{
		email = Config.emailId;
		defaultLocation1 = Config.defaultLocation1;
		defaultLocation2 = Config.defaultLocation2;
		expectedLocation = Config.expectedLocation;
		pdfBackupLocation = Config.pdfBackupLocation;
		try {

			try {
				ResultSet rs = null;
				Connection conn = JDBCConnection.establishHTML5BillingDBConn();
				Statement st = conn.createStatement();
				rs = st.executeQuery("select * from business_day");
				while (rs.next()) {
					businessDate = rs.getString("business_date");
				}
				rs.close();
				st.close();
				JDBCConnection.closeConnection(conn);
			} catch (Exception e) {
				log.error("Failed to get Current Date from Database<-|-> LocalizeMessage " + e.getLocalizedMessage()
						+ " <-|-> Message " + e.getMessage() + " <-|-> Cause " + e.getCause());
				throw new Exception("Failed to get Current Date from Database<-|-> LocalizeMessage "
						+ e.getLocalizedMessage() + " <-|-> Message" + e.getMessage() + " <-|-> Cause " + e.getCause());
			}

			SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SS");
			Date newBDate = simpleDateFormat.parse(businessDate);
			DateFormat cycleDateFormat = new SimpleDateFormat("MM/dd/yyyy");
			DateFormat eVoltDateFormat = new SimpleDateFormat("yyyyMMdd");
			String cycleDateValue = cycleDateFormat.format(newBDate);
			Date Date = new Date();
			final String execDate = eVoltDateFormat.format(Date);

			pdfCompare = pdfCompareParam.split(",");

			agencyOrPolicy = pdfCompare[0];
			billingType = pdfCompare[1];

			log.info("New PDF compare date is :" + cycleDateValue);
			log.info("New PDF compare date is :" + execDate);

			if (!cycleDateValue.equalsIgnoreCase("")) {
				BCFolderValue = cycleDateValue.replace("/", "_");
				FileNameDate = execDate.replace("_", "");
				log.info("Beyond Compare  path corrected to : " + BCFolderValue);
				expectedFile_Path = System.getProperty("user.dir") + "\\" + Config.projectName + "\\" + expectedLocation
						+ "\\" + BCFolderValue + "_" + TestCaseID;
				pdfParams[0] = "pdf";
				pdfParams[1] = expectedFile_Path;
				actualFile_Path = System.getProperty("user.dir") + "\\" + Config.projectName + "\\" + actualLocation
						+ "\\" + BCFolderValue + "_" + TestCaseID;
				pdfParams[2] = actualFile_Path;
				pdfParams[3] = Config.pdfParamPath;
				pdfParams[4] = Config.pdfParamResultPath;
			} else {
				// Deadcode
				// expectedFile_Path = expectedFile_Path;
				// actualFile_Path = actualFile_Path;
			}

			File dest = new File(actualFile_Path);
			File backup = new File(pdfBackupLocation);
			try {

				if (billingType.equalsIgnoreCase("AGENCY")) {
					File file = new File(defaultLocation1);
					File[] files = file.listFiles(new FilenameFilter() {
						@Override
						public boolean accept(File dir, String name) {
							if (name.toUpperCase().startsWith(agencyOrPolicy + "_" + FileNameDate)) {
								return true;
							} else {
								return false;
							}
						}
					});
					for (File f : files) {
						log.info(f.getName());
						FileUtils.copyFileToDirectory(f, dest);
						FileUtils.copyFileToDirectory(f, backup);
						f.delete();
					}
				}
				// When the Single billing PDF is generated :Mukul 08/28/2018
				else if (billingType.equalsIgnoreCase("SINGLE")) {
					File file = new File(defaultLocation2);
					File[] files = file.listFiles(new FilenameFilter() {
						@Override
						public boolean accept(File dir, String name) {
							if (name.toUpperCase().startsWith(agencyOrPolicy + "_" + FileNameDate)) {
								return true;
							} else {
								return false;
							}
						}
					});

					for (File f : files) {
						log.info(f.getName());
						FileUtils.copyFileToDirectory(f, dest);
						FileUtils.copyFileToDirectory(f, backup);
						f.delete();
					}
				}
			} catch (IOException e) {
				e.printStackTrace();
			}

			File f_Expected = new File(expectedFile_Path);
			File f_Actual = new File(actualFile_Path);

			File[] fileList_Expected = f_Expected.listFiles();
			File[] fileList_Actual = f_Actual.listFiles();

			log.info("Count" + fileList_Expected.length);
			log.info("Count" + fileList_Actual.length);

			if (fileList_Expected.length == fileList_Actual.length)

			{
				for (int i = 0; i < fileList_Expected.length; i++) {

					String itrateFileName = fileList_Expected[i].getName();
					log.info(itrateFileName);
					File myFile = new File(actualFile_Path + "\\" + fileList_Actual[i].getName());
					myFile.renameTo(new File(actualFile_Path + "\\" + itrateFileName));
				}
				log.info("File renamed successfully");
			} else if (fileList_Expected.length != fileList_Actual.length) {

				for (File file : dest.listFiles())
					if (!file.isDirectory())
						file.delete();
				log.info("The count of Expected and actual PDF files is not same.");
			}
		} catch (Exception e) {

			log.error(e.toString());
			e.printStackTrace();
		}

	}

	public static boolean callBeyondCompare() {
		boolean exitVal = true;

		try {

			log.info("Start Time : " + System.currentTimeMillis());
			exitVal = CompareUtil.compareDirectories(pdfParams);
			log.info("End Time : " + System.currentTimeMillis());
			System.out.println("Exited with error code " + exitVal);
		} catch (Exception e) {
			log.error(e.toString());
			e.printStackTrace();
		}
		return exitVal;
	}

}
