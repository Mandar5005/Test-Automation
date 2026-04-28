package com.majesco.itaf.main;

import java.io.IOException;
import javax.swing.JFrame;
//import org.apache.logging.log4j.LogManager;
//import org.apache.logging.log4j.Logger;
import com.majesco.itaf.vo.Reporter;


public abstract class ITAFWebDriver {
	
	//private final static Logger log = LogManager.getLogger(Automation.class.getName());

	Reporter report = new Reporter();
	JFrame frame = new JFrame("iTAF FRAMEWORK");
	String email = "";
	private static final String BILLING_APPLICATION_NAME = "Billing";
	private static final String CLAIMS_APPLICATION_NAME = "Claims";
	private static final String PAS_APPLICATION_NAME = "PAS";
	private static final String DM_APPLICATION_NAME = "DM";
	private static final String LNA_APPLICATION_NAME = "LNA";
	private static final String PLAT_APPLICATION_NAME = "DS";
	private static final String GB_APPLICATION_NAME = "GB";
	private static final String SUITE_APPLICATION_NAME = "Suite";
	private static final String CS_APPLICATION_NAME = "CS";
	private static String applicationName = null;
	private static String suitePrimaryApplicationName = null;
	//private final static Logger log = LogManager.getLogger(Config.class);
	
	public static ITAFWebDriver getInstance() {
		return ObjectFactory.getWebDriver();
	}
	public static void main(String args[]) throws Exception {
		
		String configFilePathFileName = args[0];
		Config.initializeConfigMap(configFilePathFileName);
		
		/*
		 * if (args.length > 1) { String configFilePathFileName = args[0]; String
		 * condition = args[1];
		 * 
		 * if ("Billing".equals(condition)) {
		 * log.info("initiating initializeConfigMap for Billing");
		 * Config.initializeConfigMap(configFilePathFileName); } else {
		 * log.info("initiating initializeConfigMapNew");
		 * Config.initializeConfigMapNew(configFilePathFileName); } } else { System.out.
		 * println("Insufficient arguments provided. Please provide both the config file path and the condition."
		 * ); }
		 */

		if (args.length > 1) {
			applicationName = args[1];
			suitePrimaryApplicationName = args[1];
		} else {
			applicationName = Config.applicationName;
			suitePrimaryApplicationName = Config.applicationName;
		}
		//log.info("Initializing Object...");
		ObjectFactory.initialize();
	}

	protected abstract void init() throws IOException;

	public void DataInput(String filePath, String testcaseID, String transactionType, String transactionCode,
			String operationType) throws Exception {
	}

	public void DataInput(String filePath, String testcaseID, String transactionType, String transactionCode,
			String operationType, String cycleDate) throws Exception {
	}

	public void DataInput(String structurePath, String filePath, String testcaseID, String transactionType,
			String transactionCode, String operationType, String cycleDate) throws Exception {
	}

	public Reporter getReport() {
		return report;
	}

	public void setReport(Reporter report) {
		this.report = report;
	}

	public JFrame getFrame() {
		return frame;
	}

	public String getEmail() {
		return email;
	}

	public static boolean isBillingApplication() {
		return BILLING_APPLICATION_NAME.equalsIgnoreCase(applicationName);
	}

	public static boolean isPASApplication() {
		return PAS_APPLICATION_NAME.equalsIgnoreCase(applicationName);
	}

	public static boolean isClaimsApplication() {
		return CLAIMS_APPLICATION_NAME.equalsIgnoreCase(applicationName);
	}

	public static boolean isDMApplication() {
		return DM_APPLICATION_NAME.equalsIgnoreCase(applicationName);
	}

	public static boolean isLNAApplication() {
		return LNA_APPLICATION_NAME.equalsIgnoreCase(applicationName);
	}

	public static boolean isSuiteApplication() {
		return SUITE_APPLICATION_NAME.equalsIgnoreCase(applicationName);
	}

	public static boolean isPLATApplication() {
		return PLAT_APPLICATION_NAME.equalsIgnoreCase(applicationName);
	}

	public static boolean isGBApplication() {
		return GB_APPLICATION_NAME.equalsIgnoreCase(applicationName);
	}

	public static boolean isCSApplication() {
		return CS_APPLICATION_NAME.equalsIgnoreCase(applicationName);
	}

	public static String getApplicationName() {
		return applicationName;
	}

	public static void setApplicationName(String applicationName) {
		ITAFWebDriver.applicationName = applicationName;
	}

	public static String getSuitePrimaryApplicationName() {
		return suitePrimaryApplicationName;
	}

}
