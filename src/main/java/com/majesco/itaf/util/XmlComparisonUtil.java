package com.majesco.itaf.util;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.majesco.compare.CompareUtil;
import com.majesco.compare.xml.XMLResultBean;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.ITAFWebDriver;
import com.majesco.itaf.main.MainController;
import com.majesco.itaf.main.ObjectFactory;
import com.majesco.utils.UnzipFolder;

public class XmlComparisonUtil {
	private final static Logger log = LogManager.getLogger(XmlComparisonUtil.class.getName());
	public static String ActualXmlPath = null;
	public static String NodeLevelTagsToIgnore = null;
	public static String AttributeLevelTagsToIgnore = null;
	private static MainController controller = ObjectFactory.getMainController();
	public static ITAFWebDriver webDriverXML = ITAFWebDriver.getInstance();
	static {
		if (Config.xmlDefaultAttributeLevelTagsToIgnore == null || Config.xmlDefaultAttributeLevelTagsToIgnore == "") {
			AttributeLevelTagsToIgnore = "";
		} else {
			AttributeLevelTagsToIgnore = Config.xmlDefaultAttributeLevelTagsToIgnore.trim();
		}

		if (Config.xmlDefaultNodeLevelTagsToIgnore == null || Config.xmlDefaultNodeLevelTagsToIgnore == "") {
			NodeLevelTagsToIgnore = "";
		} else {
			NodeLevelTagsToIgnore = Config.xmlDefaultNodeLevelTagsToIgnore.trim();
		}

	}

	public static void setActualXmlPath(String zipFileLocation) {
		if (zipFileLocation.isEmpty()) {
			ActualXmlPath = "";
			System.out.println("Actual xml path not found during XML comparison");
		} else {
			List<File> fileList = new ArrayList<>();
			File file = new File(zipFileLocation);
			UnzipFolder.getXML(file, fileList);
			for (File f : fileList) {
				ActualXmlPath = f.getAbsolutePath();
			}
			log.info("Actual Xml Path for XML comparison is  '" + ActualXmlPath + "'");
		}
	}

	public static void unZipXmlFolder(String zipFileLocation) {
		if (zipFileLocation.isEmpty()) {
			System.out.println("Zipped XML Path not found");
		} else {
			List<File> fileList = new ArrayList<>();
			File file = new File(zipFileLocation);
			UnzipFolder.getAllZipFolders(file, fileList);
			UnzipFolder.unzip(fileList);
			fileList.clear();
		}
	}

	public static void setNodeLevelTagsToIgnore(String str) {
		if (str.isEmpty() && NodeLevelTagsToIgnore == "") {
			System.out.println("No input is provided to ignore Node level tags during XML comparison");
			log.info("No input is provided to ignore Node level tags during XML comparison");
		} else {
			if (!str.isEmpty() && !(NodeLevelTagsToIgnore == "")) {
				NodeLevelTagsToIgnore = NodeLevelTagsToIgnore + "," + str.trim();
			} else if (NodeLevelTagsToIgnore == "") {
				NodeLevelTagsToIgnore = str.trim();
			}

			System.out.println("String to ignore  Node level tags during XML comparison is set to '"
					+ NodeLevelTagsToIgnore + "'");
			log.info("String to ignore  Node level tags during XML comparison is set to '" + NodeLevelTagsToIgnore
					+ "'");
		}
	}

	public static void setAttributeLevelTagsToIgnore(String str) {
		if (str.isEmpty() && AttributeLevelTagsToIgnore == "") {
			System.out.println("No input is provided to ignore Attribute level tags during XML comparison");
			log.info("No input is provided to ignore Attribute level tags during XML comparison");
		} else {

			if (!str.isEmpty() && !(AttributeLevelTagsToIgnore == "")) {
				AttributeLevelTagsToIgnore = AttributeLevelTagsToIgnore + "," + str.trim();
			} else if (AttributeLevelTagsToIgnore == "") {
				AttributeLevelTagsToIgnore = str.trim();
			}

			System.out.println("String to ignore Attribute level tags during XML comparison is set to '"
					+ AttributeLevelTagsToIgnore + "'");
			log.info("String to ignore Attribute level tags during XML comparison is set to '"
					+ AttributeLevelTagsToIgnore + "'");
		}
	}

	public static String[] XMLCompare(String logicalName, String ctrlValue) throws Exception {

		// get path of files for comparison
		String expFilePath = Config.expXMLFilePath + "\\" + ctrlValue;
		// System.out.println("Expected XML Path " + expFilePath);
		String actualFilePath = Config.actualXMLDownloadPath + "\\" + ctrlValue;
		// System.out.println("Actual XML Path " + actualFilePath);
		@SuppressWarnings("unused")
		String resultFilePath = Config.xmlCompResultPath;

		// create TC_Trasnaction folder if not present
		String folderPath = Config.xmlCompResultPath + "\\" + controller.controllerTestCaseID + "\\"
				+ controller.controllerTransactionType;
		File createDir = new File(folderPath);
		if (!createDir.exists()) {
			createDir.mkdirs();
		}

		// create subfolder with timestamp inside folder created above
		String timeStampForFolder = new SimpleDateFormat("ddMMMMyyyy_HH_mm_ss").format(new Date());
		String subFolderPathForOutput = folderPath + "\\" + timeStampForFolder;
		File subFolder = new File(subFolderPathForOutput);
		if (!subFolder.exists()) {
			subFolder.mkdirs();
		}

		String[] args = new String[6];

		// xml type
		args[0] = "xml";
		// Expected file path
		args[1] = expFilePath;

		// Actual file path
		args[2] = actualFilePath;

		// Results folder path
		args[3] = subFolderPathForOutput;

		// Attribute level tags to ignore during comparison
		args[4] = AttributeLevelTagsToIgnore;

		// Node level tags to ignore during comparison
		args[5] = NodeLevelTagsToIgnore;

		// Check for presence of expected and actual files else throw error
		File expFile = new File(expFilePath);
		if (!expFile.exists()) {
			log.info("Expected XML file missing - " + expFilePath);
			if (subFolder.exists()) {
				subFolder.delete();
			}
			webDriverXML.getReport().setMessage("Expected XML File missing - " + expFilePath);
			webDriverXML.getReport().setStatus("FAIL");
			return (new String[0]);
			// throw new FileNotFoundException("Expected XML File missing - "+
			// expFilePath);
		}
		File actualFile = new File(actualFilePath);
		if (!actualFile.exists()) {
			log.info("Actual XML file missing - " + actualFilePath);
			if (subFolder.exists()) {
				subFolder.delete();
			}
			webDriverXML.getReport().setMessage("Actual XML File missing - " + actualFilePath);
			webDriverXML.getReport().setStatus("FAIL");
			return (new String[0]);
			// throw new
			// FileNotFoundException("Actual XML File missing - "+actualFilePath);
		}

		log.info("XML comparison started at : " + new SimpleDateFormat("ddMMMM HH:mm:ss").format(new Date()));
		log.info("XML comparison is in progess");

		XMLResultBean xmlResultBean = new XMLResultBean();
		try {
			xmlResultBean = (XMLResultBean) CompareUtil.compareFiles(args);
		} catch (Throwable e) {
			e.printStackTrace();
			log.error(e.getMessage(), e);

			webDriverXML.getReport().setMessage("Error during XML Comparison - " + e);
			webDriverXML.getReport().setStatus("FAIL");
			return (new String[0]);
			// throw new Exception("Error during XML Comparison - " +e);
		}
		log.info("XML comparison completed at: " + new SimpleDateFormat("ddMMMM HH:mm:ss").format(new Date()));

		String returnVal1 = "", returnVal2 = "";
		if (xmlResultBean.isError() && xmlResultBean.getDifferenceCount() > 0) {
			returnVal1 = "Mismatch found in " + xmlResultBean.getDifferenceCount() + " xml tag values.";
			returnVal2 = "\nResults saved at: " + subFolderPathForOutput;
		} else {
			returnVal1 = "XML Comparison is successful, NO mismatch found.";
			returnVal2 = returnVal1;
			if (subFolder.exists()) {
				String[] listFiles = subFolder.list();
				for (String s : listFiles) {
					File currentFile = new File(subFolder.getPath(), s);
					currentFile.delete();
				}
				subFolder.delete();
			}
		}

		String[] result = new String[3];
		result[0] = String.valueOf(xmlResultBean.isError());
		result[1] = returnVal2;
		result[2] = returnVal1;
		return result;
	}

}
