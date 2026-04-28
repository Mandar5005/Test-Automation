package com.majesco.itaf.main;

import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import com.majesco.compare.CompareUtil;
import com.majesco.itaf.util.ExcelUtility;
import com.majesco.itaf.vo.Reporter;

public abstract class MainController {

	private final static Logger log = LogManager.getLogger(MainController.class);

	public boolean pauseExecution = false;
	public boolean stopExecution = false;
	public Cell controllerGroupName = null;
	public Cell controllerTestCaseID = null;
	public Cell controllerTransactionType = null;
	public String controllerQuoteId = null;
	public Cell controllerSC_No = null;
	public String testDescription = null;
	public String FailScreen = "";
	public String cycleDateCellValue = "";
	protected ResultSet rs = null;
	public String businessDate = null;
	public String businessDateValue = null;
	public String cycleDateValue = null;
	public String xmlbusinessdate = null;
	public boolean pExecution = false;
	protected static List<String> objFailedAccountList = new ArrayList<>();
	protected static List<String> objFailedPolicyList = new ArrayList<>();
	protected static List<String> objFailedBrokerList = new ArrayList<>();

	public abstract Reporter ControllerData(String FilePath) throws NullPointerException, Exception;

	public abstract boolean pauseFun(String message);

	public abstract void recoveryhandler();

	public abstract void batchRecoveryScenario(String batchNo);

	public void verifyCSVReport() {

		try {

			log.info("###################### Entered VerifyReport Condition ##################");

			String[] input = { Config.csvCompareReportFormat, Config.csvCompareReportExpectedPath,
					Config.csvCompareReportActualPath, Config.csvCompareReportConfigPath,
					Config.csvCompareReportResultPath };

			if (log.isDebugEnabled()) {
				log.debug("--------- Reports Path are : ------------");
				log.debug("REPORTFORMAT = " + Config.csvCompareReportFormat);
				log.debug("REPORTEXPECTEDPATH = " + Config.csvCompareReportExpectedPath);
				log.debug("REPORTACTUALPATH = " + Config.csvCompareReportActualPath);
				log.debug("REPORTCONFIGPATH = " + Config.csvCompareReportConfigPath);
				log.debug("REPORTRESULTPATH = " + Config.csvCompareReportResultPath);
			}

			String strErrorMessage = "";
			boolean bResult;
			String StrStatus = "PASS";

			try {
				Map<String, List<String>> objReportMap = new HashMap<>();

				objReportMap.put("SkipRowIdentifier1", objFailedAccountList);
				log.info("Setting the value of SkipRowIdentifier1 : " + objFailedAccountList);
				objReportMap.put("SkipRowIdentifier2", objFailedPolicyList);
				objReportMap.put("SkipRowIdentifier3", objFailedBrokerList);
				log.info("The value in the map is : " + objReportMap.toString());
				log.info("Input for CSV report comparison :" + input.toString());
				CompareUtil.populateDataValuesForSkipIdentifier(objReportMap);
				bResult = CompareUtil.compareDirectories(input);
				log.info("Report Execution Status :" + bResult);

			} catch (Exception e) {
				strErrorMessage = e.getMessage();
				System.out.println("Exception occurred");
				e.printStackTrace();
				bResult = false;
			} catch (Throwable th) {

				th.printStackTrace();

				bResult = false;
			}
			if (!bResult) {
				StrStatus = "FAIL";

				// recovery handler();
			}
			Reporter objReport = new Reporter();
			objReport.setMessage(strErrorMessage);
			System.out.println("Error msg = " + strErrorMessage.toString());
			objReport.setTrasactionType(controllerTransactionType.toString());
			objReport.setTestcaseId("Common");
			objReport.setCycleDate(cycleDateCellValue);
			objReport.setStatus(StrStatus);
			objReport.setFromDate("");
			ExcelUtility.writeReport(objReport);

			log.info("###################### End VerifyReport Condition ##################");

		} catch (Exception e) {
			log.error(e.getMessage(), e);
		}
	}

}
