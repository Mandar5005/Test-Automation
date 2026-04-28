package com.majesco.itaf.vo;

import org.openqa.selenium.WebDriver;

public class Reporter {

	private String strIteration;
	private String strTestDescription;
	private String strGroupName;
	private String strTestcaseId;
	private String strTrasactionType;
	private String strVerificationType;
	private String strStatus;
	private String strStepName;
	private String strMessage;
	private String strFieldName;
	private String strExpectedValue;
	private String strActualValue;
	private String toDate;
	private String frmDate;
	private String strTransactioncode;
	private String strDirectoryPath;
	private String strOperationType;
	private String strInputPath;
	private String strScreenshot = "";
	private WebDriver driver;
	private String mainWindowHandle;
	private String strCycleDate = "";
	private Reporter report = null;
	private String passCount;
	private String failCount;
	private String columnName;
	private String silo;
	private String transactionVerificationType;

	/**
	 * @return the columnName
	 */
	public String getColumnName() {
		return columnName;
	}

	/**
	 * @param columnName the columnName to set
	 */
	public void setColumnName(String columnName) {
		this.columnName = columnName;
	}

	/**
	 * @return the passCount
	 */
	public String getPassCount() {
		return passCount;
	}

	/**
	 * @param passCount the passCount to set
	 */
	public void setPassCount(String passCount) {
		this.passCount = passCount;
	}

	/**
	 * @return the failCount
	 */
	public String getFailCount() {
		return failCount;
	}

	/**
	 * @param failCount the failCount to set
	 */
	public void setFailCount(String failCount) {
		this.failCount = failCount;
	}

	private String strQuoteId;

	public String getStrQuoteId() {

		return strQuoteId;
	}

	public void setStrQuoteId(String strQuoteId) {

		this.strQuoteId = strQuoteId;
	}

	public void setCycleDate(String strCycleDate) {
		this.strCycleDate = strCycleDate;
	}

	public String getCycleDate() {
		return this.strCycleDate;
	}

	public void setDriver(WebDriver driver) {
		this.driver = driver;
		this.mainWindowHandle = this.driver.getWindowHandle();
	}

	public WebDriver getDriver() {
		return driver;
	}

	public String getMainWindowHandle() {
		return this.mainWindowHandle;
	}

	public String getScreenShot() {
		return strScreenshot;
	}

	public void setScreenShot(String strScreenshot) {
		this.strScreenshot = strScreenshot;
	}

	public String getInputPath() {
		return strInputPath;
	}

	public void setInputPath(String strInputPath) {
		this.strInputPath = strInputPath;
	}

	public String getOperationType() {
		return strOperationType;
	}

	public void setOperationType(String strOperationType) {
		this.strOperationType = strOperationType;
	}

	public String getDirectoryPath() {
		return strDirectoryPath;
	}

	public void setDirectoryPath(String strDirectoryPath) {
		this.strDirectoryPath = strDirectoryPath;
	}

	public String getTransactioncode() {
		return strTransactioncode;
	}

	public void setTransactioncode(String strTransactioncode) {
		this.strTransactioncode = strTransactioncode;
	}

	public String getIteration() {
		return strIteration;
	}

	public void setIteration(String strIteration) {
		this.strIteration = strIteration;
	}

	public String getTestcaseId() {
		return strTestcaseId;
	}

	public void setTestcaseId(String strTestcaseId) {
		this.strTestcaseId = strTestcaseId;
	}

	public String getGroupName() {
		return strGroupName;
	}

	public void setGroupName(String strGroupName) {
		this.strGroupName = strGroupName;
	}

	public String getTrasactionType() {
		return strTrasactionType;
	}

	public void setTrasactionType(String strTrasactionType) {
		this.strTrasactionType = strTrasactionType;
	}

	public String getVerificationType() {
		return strVerificationType;
	}

	public void setVerificationType(String strVerifcationType) {
		this.strVerificationType = strVerifcationType;
	}

	public String getStatus() {
		return strStatus;
	}

	public void setStatus(String strStatus) {
		this.strStatus = strStatus;
	}

	public String getStepName() {
		return strStepName;
	}

	public void setStepName(String strStepName) {
		this.strStepName = strStepName;
	}

	public String getMessage() {
		return strMessage;
	}

	public void setMessage(String strMessage) {
		this.strMessage = strMessage;
	}

	public String getFieldName() {
		return strFieldName;
	}

	public void setFieldName(String strFieldName) {
		this.strFieldName = strFieldName;
	}

	public String getExpectedValue() {
		return strExpectedValue;
	}

	public void setExpectedValue(String strExpectedValue) {
		this.strExpectedValue = strExpectedValue;
	}

	public String getActualValue() {
		return strActualValue;
	}

	public void setActualValue(String strActualValue) {
		this.strActualValue = strActualValue;
	}

	public String getToDate() {
		return toDate;
	}

	public void setToDate(String toDate) {
		this.toDate = toDate;
	}

	public String getFromDate() {
		return frmDate;
	}

	public void setFromDate(String frmDate) {
		this.frmDate = frmDate;
	}

	public Reporter getReport() {
		return report;
	}

	public void setReport(Reporter report) {
		this.report = report;
	}

	public String getTestDescription() {
		return strTestDescription;
	}

	public void setTestDescription(String strTestDescription) {
		this.strTestDescription = strTestDescription;
	}

	public String getSilo() {
		return silo;
	}

	public void setSilo(String silo) {
		this.silo = silo;
	}

	public String getTransactionVerificationType() {
		return transactionVerificationType;
	}

	public void setTransactionVerificationType(String transactionVerificationType) {
		this.transactionVerificationType = transactionVerificationType;
	}
}
