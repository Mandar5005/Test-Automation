package com.majesco.itaf.util;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.MainController;
import com.majesco.itaf.main.ObjectFactory;
import com.majesco.itaf.vo.Reporter;

public class DB_Validation {

	public static MainController controller = ObjectFactory.getMainController();

	public static Connection conn = null;
	public static Statement stmt = null;
	public static ResultSet rs = null;
	public static String dateForFileName = (controller.cycleDateCellValue).replaceAll("/", "_");
	public static String filename2 = Config.getConfigValue("RESULT_FILEPATH") + "DBVerification//DBData_Result" + "_"
			+ controller.controllerTestCaseID.toString() + ".csv";
	public static File f2 = new File(filename2);
	public static char myChar = 34;
	public static PrintStream print = null;
	protected static List<String> status = new ArrayList<String>();
	protected static List<String> rowStatus = new ArrayList<String>();
	protected static List<String> actualValue = new ArrayList<String>();
	public static List<Integer> PassCount = new ArrayList<Integer>();
	public static List<Integer> FailCount = new ArrayList<Integer>();
	public static int firstRow = 1;
	public static int TotalpassCount;
	public static int TotalfailCount;
	protected static List<List<String>> actualRows = new ArrayList<List<String>>();
	public static int DResult = 1;
	public static Date frmDate = new Date();
	public static DateFormat dtFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");

	private static List<String> columns = new ArrayList<String>();
	private static List<String> columnsData = new ArrayList<String>();

	public static void DBValidation(String ctrlValue) throws Exception {

		String filename = Config.getConfigValue("ACTUALVALUESPATH").toString()
				+ controller.controllerTestCaseID.toString() + "_" + dateForFileName
				+ ".xls";

		File f = new File(filename);

		String filename1 = Config.getConfigValue("EXPECTEDVALUESPATH").toString()
				+ controller.controllerTestCaseID.toString() + "_" + dateForFileName/* controller.cycleDateCellValue */
				+ "_Expected" + ".xls";

		File f1 = new File(filename1);

		Class.forName("oracle.jdbc.driver.OracleDriver");
		DriverManager.registerDriver(new oracle.jdbc.driver.OracleDriver());
		conn = DriverManager.getConnection(
				"jdbc:oracle:thin:@" + Config.databaseHost + ":" + Config.databasePort + ":" + Config.databaseSID,
				Config.applicationdatabaseusername, Config.applicationdatabasepassword);
		stmt = conn.createStatement();

		ResultSet rs = stmt
				.executeQuery("SELECT * FROM nd_wang_recap_history where pol_num='" + ctrlValue + "' " + "order by 2");
		System.out.println(rs);

		if (f.exists() && !f.isDirectory()) {
			filename = Config.getConfigValue("ACTUALVALUESPATH").toString() + controller.controllerTestCaseID.toString()
			+ "_" + dateForFileName
			+ ".xls";
		} else {
			filename = Config.getConfigValue("ACTUALVALUESPATH").toString() + controller.controllerTestCaseID.toString()
			+ "_" + dateForFileName
			+ ".xls";
		}

		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("Actual");

		Row rowhead = sheet.createRow((short) 0);
		rowhead.createCell((short) 0).setCellValue("Policy_Number");
		rowhead.createCell((short) 1).setCellValue("Process_Date");
		rowhead.createCell((short) 2).setCellValue("Addl_Key");
		rowhead.createCell((short) 3).setCellValue("Agent_Num");
		rowhead.createCell((short) 4).setCellValue("Effective_Date");
		rowhead.createCell((short) 5).setCellValue("Exp_Date");
		rowhead.createCell((short) 6).setCellValue("Total_Prem");
		rowhead.createCell((short) 7).setCellValue("Change_Date");
		rowhead.createCell((short) 8).setCellValue("Group_Line");
		rowhead.createCell((short) 9).setCellValue("Bill_Type");
		rowhead.createCell((short) 10).setCellValue("Pay_Option");
		rowhead.createCell((short) 11).setCellValue("Ext_Int_SW");
		rowhead.createCell((short) 12).setCellValue("Written_Prem");
		rowhead.createCell((short) 13).setCellValue("State_CD");
		rowhead.createCell((short) 14).setCellValue("Company");
		rowhead.createCell((short) 15).setCellValue("Act_Code");
		rowhead.createCell((short) 16).setCellValue("Act_Date");
		rowhead.createCell((short) 17).setCellValue("Name");
		rowhead.createCell((short) 18).setCellValue("Int_Date");
		rowhead.createCell((short) 19).setCellValue("Cancel_Reas");
		rowhead.createCell((short) 20).setCellValue("Bill_Form1");
		rowhead.createCell((short) 21).setCellValue("Ret_Ck_Reas");
		rowhead.createCell((short) 22).setCellValue("Amount_Due");
		rowhead.createCell((short) 23).setCellValue("Cash_Amt");
		rowhead.createCell((short) 24).setCellValue("Service_Charge");
		rowhead.createCell((short) 25).setCellValue("Prev_Amt_due");
		rowhead.createCell((short) 26).setCellValue("Ret_Ck_Amt");
		rowhead.createCell((short) 27).setCellValue("Times_Billed");
		rowhead.createCell((short) 28).setCellValue("Tot_To_Bill");
		rowhead.createCell((short) 29).setCellValue("Inst_Svc_Chg");
		rowhead.createCell((short) 30).setCellValue("Inst_Amt");
		rowhead.createCell((short) 31).setCellValue("Terminal_User");
		rowhead.createCell((short) 32).setCellValue("Rep_Unpaid_Amt");
		rowhead.createCell((short) 33).setCellValue("Due_Date");
		rowhead.createCell((short) 34).setCellValue("Arrears_Amt");
		rowhead.createCell((short) 35).setCellValue("Arrears_Svc_Chg");
		rowhead.createCell((short) 36).setCellValue("Proc_Center");
		rowhead.createCell((short) 37).setCellValue("Earned_Prem_RLetter_Sw");
		rowhead.createCell((short) 38).setCellValue("Run_Date");
		rowhead.createCell((short) 39).setCellValue("Account_Number");
		rowhead.createCell((short) 40).setCellValue("MBS_Trans_Type");
		rowhead.createCell((short) 41).setCellValue("MBS_Bill_Send_date");
		rowhead.createCell((short) 42).setCellValue("MBS_Bill_Due_Date");
		rowhead.createCell((short) 43).setCellValue("MBS_Canc_Check_Date");
		rowhead.createCell((short) 44).setCellValue("MBS_Canc_Reason");
		rowhead.createCell((short) 45).setCellValue("MBS_Audit_Flag");
		rowhead.createCell((short) 46).setCellValue("MBS_Num_Demand_Notices");
		rowhead.createCell((short) 47).setCellValue("MBS_Num_Nsfs");
		rowhead.createCell((short) 48).setCellValue("MBS_Num_Ep_Invoice");
		rowhead.createCell((short) 49).setCellValue("MBS_Num_Noc");
		rowhead.createCell((short) 50).setCellValue("MBS_Policy_Status");
		rowhead.createCell((short) 51).setCellValue("Policy_Hold");
		rowhead.createCell((short) 52).setCellValue("Account_Hold");
		rowhead.createCell((short) 53).setCellValue("Noc_Hold");
		rowhead.createCell((short) 54).setCellValue("MBS_System_Activity_Num");
		rowhead.createCell((short) 55).setCellValue("MBS_Source_Trans_Seq");
		rowhead.createCell((short) 56).setCellValue("Output_Seq");
		rowhead.createCell((short) 57).setCellValue("Payment_Id");
		rowhead.createCell((short) 58).setCellValue("MBS_Trans_Reason");
		rowhead.createCell((short) 59).setCellValue("MBS_Payment_Method");
		rowhead.createCell((short) 60).setCellValue("MBS_Policy_Term_Id");
		rowhead.createCell((short) 61).setCellValue("MBS_Account_System_Code");
		rowhead.createCell((short) 62).setCellValue("Batch_No");
		rowhead.createCell((short) 63).setCellValue("Payment_Seq");
		rowhead.createCell((short) 64).setCellValue("Source_System_Process_Date");

		Map<String, Object[]> excel_data = new HashMap<String, Object[]>();
		int row_counter = 0;

		while (rs.next()) {
			row_counter = row_counter + 1;
			String pol_num = rs.getString("POL_NUM");
			String process_date = rs.getString("PROCESS_DATE");
			String addl_key = rs.getString("ADDL_KEY");
			String agent_num = rs.getString("AGENT_NUM");
			String eff_date = rs.getString("EFF_DATE");
			String exp_date = rs.getString("EXP_DATE");
			String total_prem = rs.getString("TOTAL_PREM");
			String change_date = rs.getString("CHANGE_DATE");
			String group_line = rs.getString("GROUP_LINE");
			String bill_type = rs.getString("BILL_TYPE");
			String pay_option = rs.getString("PAY_OPTION");
			String ext_int_sw = rs.getString("EXT_INT_SW");
			String written_prem = rs.getString("WRITTEN_PREM");
			String state_cd = rs.getString("STATE_CD");
			String company = rs.getString("COMPANY");
			String act_code = rs.getString("ACT_CODE");
			String act_date = rs.getString("ACT_DATE");
			String name = rs.getString("NAME");
			String int_date = rs.getString("INT_DATE");
			String cancel_reas = rs.getString("CANCEL_REAS");
			String bill_form1 = rs.getString("BILL_FORM1");
			String ret_ck_reas = rs.getString("RET_CK_REAS");
			String amount_due = rs.getString("AMOUNT_DUE");
			String cash_amt = rs.getString("CASH_AMT");
			String service_charge = rs.getString("SERVICE_CHARGE");
			String prev_amt_due = rs.getString("PREV_AMT_DUE");
			String ret_ck_amt = rs.getString("RET_CK_AMT");
			String times_billed = rs.getString("TIMES_BILLED");
			String tot_to_bill = rs.getString("TOT_TO_BILL");
			String inst_svc_chg = rs.getString("INST_SVC_CHG");
			String inst_amt = rs.getString("INST_AMT");
			String terminal_user = rs.getString("TERMINAL_USER");
			String rep_unpaid_amt = rs.getString("REP_UNPAID_AMT");
			String due_date = rs.getString("DUE_DATE");
			String arrears_amt = rs.getString("ARREARS_AMT");
			String arrears_svc_chg = rs.getString("ARREARS_SVC_CHG");
			String proc_center = rs.getString("PROC_CENTER");
			String earned_prem_rletter_sw = rs.getString("EARNED_PREM_RLETTER_SW");
			String run_date = rs.getString("RUN_DATE");
			String account_number = rs.getString("ACCOUNT_NUMBER");
			String mbs_trans_type = rs.getString("MBS_TRANS_TYPE");
			String mbs_bill_send_date = rs.getString("MBS_BILL_SEND_DATE");
			String mbs_bill_due_date = rs.getString("MBS_BILL_DUE_DATE");
			String mbs_canc_check_date = rs.getString("MBS_CANC_CHECK_DATE");
			String mbs_canc_reason = rs.getString("MBS_CANC_REASON");
			String mbs_audit_flag = rs.getString("MBS_AUDIT_FLAG");
			String mbs_num_demand_notices = rs.getString("MBS_NUM_DEMAND_NOTICES");
			String mbs_num_nsfs = rs.getString("MBS_NUM_NSFS");
			String mbs_num_ep_invoice = rs.getString("MBS_NUM_EP_INVOICE");
			String mbs_num_noc = rs.getString("MBS_NUM_NOC");
			String mbs_policy_status = rs.getString("MBS_POLICY_STATUS");
			String policy_hold = rs.getString("POLICY_HOLD");
			String account_hold = rs.getString("ACCOUNT_HOLD");
			String noc_hold = rs.getString("NOC_HOLD");
			String mbs_system_activity_no = rs.getString("MBS_SYSTEM_ACTIVITY_NO");
			String mbs_source_transaction_seq = rs.getString("MBS_SOURCE_TRANSACTION_SEQ");
			String outputseq = rs.getString("OUTPUTSEQ");
			String payment_id = rs.getString("PAYMENT_ID");
			String mbs_transaction_reason = rs.getString("MBS_TRANSACTION_REASON");
			String mbs_payment_method = rs.getString("MBS_PAYMENT_METHOD");
			String mbs_policy_term_id = rs.getString("MBS_POLICY_TERM_ID");
			String mbs_account_system_code = rs.getString("MBS_ACCOUNT_SYSTEM_CODE");
			String batch_no = rs.getString("BATCH_NO");
			String payment_seq = rs.getString("PAYMENT_SEQ");
			String Source_System_Process_Date = rs.getString("SOURCE_SYSTEM_PROCESS_DATE");

			excel_data.put(Integer.toString(row_counter),
					new Object[] { pol_num, process_date, addl_key, agent_num, eff_date, exp_date, total_prem,
							change_date, group_line, bill_type, pay_option, ext_int_sw, written_prem, state_cd, company,
							act_code, act_date, name, int_date, cancel_reas, bill_form1, ret_ck_reas, amount_due,
							cash_amt, service_charge, prev_amt_due, ret_ck_amt, times_billed, tot_to_bill, inst_svc_chg,
							inst_amt, terminal_user, rep_unpaid_amt, due_date, arrears_amt, arrears_svc_chg,
							proc_center, earned_prem_rletter_sw, run_date, account_number, mbs_trans_type,
							mbs_bill_send_date, mbs_bill_due_date, mbs_canc_check_date, mbs_canc_reason, mbs_audit_flag,
							mbs_num_demand_notices, mbs_num_nsfs, mbs_num_ep_invoice, mbs_num_noc, mbs_policy_status,
							policy_hold, account_hold, noc_hold, mbs_system_activity_no, mbs_source_transaction_seq,
							outputseq, payment_id, mbs_transaction_reason, mbs_payment_method, mbs_policy_term_id,
							mbs_account_system_code, batch_no, payment_seq, Source_System_Process_Date });
		}

		rs.close();
		stmt.close();
		conn.close();

		Set<String> keyset = excel_data.keySet();
		int rownum = 1;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = excel_data.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				cell.setCellValue((String) obj);
			}
		}

		if (f.exists() && !f.isDirectory()) {
			// FileOutputStream fileOut = new FileOutputStream(f1);
			FileOutputStream fileOut = new FileOutputStream(f);
			wb.write(fileOut);
			fileOut.close();
			System.out.println("Your excel file_1 has been generated!");
		} else {

			FileOutputStream fileOut = new FileOutputStream(f);
			wb.write(fileOut);
			fileOut.close();
			System.out.println("Your excel file_0 has been generated!");
		}


		// Close the Workbook object after writing data and saving the Excel file //Sel4
		wb.close();
		if (f.exists() && f1.exists()) {

			@SuppressWarnings("unused")
			DataFormatter df = new DataFormatter();
			try {
				FileInputStream file1 = new FileInputStream(f1);
				FileInputStream file2 = new FileInputStream(f);
				Workbook wb1 = WorkbookFactory.create(file1);// expected file

				Sheet sh1 = wb1.getSheet("Actual");
				if (sh1 != null) {

					Workbook wb2 = WorkbookFactory.create(file2);// actual file
					Sheet sh2 = wb2.getSheet("Actual");
					Sheet expectedSheet = sh1;
					Sheet actualSheet = sh2;
					@SuppressWarnings("unused")
					Reporter report = new Reporter();
					report = CompareExcel(actualSheet, expectedSheet, columns, columnsData);

				} else {
					Exception exception = new Exception();
					System.out.println("Expected Sheet is not present " + exception);
				}

			}

			catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

	}

	public static Reporter CompareExcel(Sheet actualSheet, Sheet expectedSheet, List<String> columns,
			List<String> columnsData) throws IOException {

		@SuppressWarnings("unused")
		boolean isrowFound = false;
		int expSheetRowCount = expectedSheet.getPhysicalNumberOfRows(); // getRowCount(expectedSheet);
		Reporter report = new Reporter();
		report.setReport(report);
		int passCount = 0;
		int failCount = 0;
		int colCount = 0;

		for (int rowIndex = firstRow; rowIndex < expSheetRowCount; rowIndex++) {
			passCount = 0;
			failCount = 0;

			Row actualRow = actualSheet.getRow(rowIndex);
			Row expectedRow = expectedSheet.getRow(rowIndex);
			System.out.println(actualRow.getCell(0).toString());
			System.out.println(expectedRow.getCell(0).toString());

			actualValue = new ArrayList<String>();
			if (actualRow == null || expectedRow == null) {
				break;
			}

			colCount = expectedRow.getLastCellNum();
			for (int columnIndex = 0; columnIndex <= colCount; columnIndex++) {
				Cell actualCell = actualRow.getCell(columnIndex);
				DataFormatter fmt = new DataFormatter();
				Cell expectedCell = expectedRow.getCell(columnIndex);

				if (actualCell != null || expectedCell != null) {
					String expectedValue = fmt.formatCellValue(expectedCell);

					if (!actualCell.toString().equalsIgnoreCase(expectedValue)) {
						if (actualCell.getColumnIndex() == 0 || actualCell.getColumnIndex() == 3
								|| actualCell.getColumnIndex() == 39) {
							passCount += 1;
							TotalpassCount += 1;
							report.setStatus("PASS");
							report.setActualValue(actualCell.toString());
							System.out.println(actualCell.toString());
						} else {
							report.setStatus("FAIL");
							failCount += 1;
							TotalfailCount += 1;
							System.out.println(TotalfailCount);

							report.setActualValue(DB_Validation.myChar + "," + DB_Validation.myChar + "$$$ "
									+ expectedValue + DB_Validation.myChar + "," + DB_Validation.myChar + "$$$ "
									+ actualCell.toString() + DB_Validation.myChar + "," + DB_Validation.myChar);

						}
					} else {
						passCount += 1;
						TotalpassCount += 1;
						report.setStatus("PASS");
						report.setActualValue(actualCell.toString());
						System.out.println(actualCell.toString());

					}
					status.add(report.getStatus());
					actualValue.add(report.getActualValue());
				}

			}
			if (status.contains("FAIL")) {
				report.setStatus("FAIL");
			} else {
				report.setStatus("PASS");
			}

			status.clear();
			System.out.println("---" + rowIndex + "----");
			rowStatus.add(report.getStatus());

			PassCount.add(passCount);
			FailCount.add(failCount);
			actualRows.add(actualValue);
			report.setReport(report);
		}

		if (rowStatus.contains("FAIL")) {
			report.setStatus("FAIL");
		}
		WriteToDetailResults(columns, actualRows, passCount, failCount, expSheetRowCount, colCount, report, rowStatus);
		PassCount.clear();
		FailCount.clear();
		return report;
	}

	public static void WriteToDetailResults(List<String> columns, List<List<String>> columnsData, int passCount,
			int failCount, int rowCount, int colCount, Reporter report, List<String> status) throws IOException {
		try {

			report = report.getReport();
			report.getFromDate();
			report.setTestcaseId(controller.controllerTestCaseID.toString());
			report.setTrasactionType(controller.controllerTransactionType.toString());
			report.getStatus();
			if (DResult == 1) {
				f2.delete();
				DResult++;
			}
			if (f2.exists() == false) {
				System.out.println("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$");
				print = new PrintStream(f2);
			}

			print = new PrintStream(new FileOutputStream(f2, true));
			int usedRows = count(f2);
			if (usedRows == 0) {

				print.println(
						"Iteration,TestCaseID,TransactionType,CurrentDate,RowType,Status,PassCount,FailCount,Pass/FailData");
			}
			usedRows = count(f2);

			print.print(DB_Validation.myChar + "" + DB_Validation.myChar + "," + DB_Validation.myChar
					+ report.getTestcaseId() + DB_Validation.myChar + "," + DB_Validation.myChar
					+ report.getTrasactionType() + DB_Validation.myChar + "," + DB_Validation.myChar
					+ report.getFromDate() + DB_Validation.myChar + "," + DB_Validation.myChar + "Header"
					+ DB_Validation.myChar + "," + DB_Validation.myChar + report.getStatus() + DB_Validation.myChar
					+ "," + DB_Validation.myChar + "" + DB_Validation.myChar + "," + DB_Validation.myChar + ""
					+ DB_Validation.myChar);
			String r = report.getStatus();

			System.out.println(r);
			int counter = 0;
			while (columns.isEmpty() == false) {
				if (counter != columns.size()) {
					print.print("," + DB_Validation.myChar + columns.get(counter) + DB_Validation.myChar);
					counter++;
				} else {
					break;
				}
			}
			print.println();
			rowCount = actualRows.size();

			for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
				print.print(DB_Validation.myChar + "" + DB_Validation.myChar + "," + DB_Validation.myChar
						+ report.getTestcaseId() + DB_Validation.myChar + "," + DB_Validation.myChar
						+ report.getTrasactionType() + DB_Validation.myChar + "," + DB_Validation.myChar
						+ report.getFromDate() + DB_Validation.myChar + "," + DB_Validation.myChar + "Data"
						+ DB_Validation.myChar + "," + DB_Validation.myChar + rowStatus.get(rowIndex).toString()
						+ DB_Validation.myChar + "," + DB_Validation.myChar + PassCount.get(rowIndex)
						+ DB_Validation.myChar + "," + DB_Validation.myChar + FailCount.get(rowIndex)
						+ DB_Validation.myChar + "," + DB_Validation.myChar + "" + DB_Validation.myChar);

				counter = 0;
				while (!actualRows.isEmpty()) {
					if (counter != actualRows.get(rowIndex).size()) {
						System.out.print(actualRows.get(rowIndex).get(counter));
						print.print(
								DB_Validation.myChar + actualRows.get(rowIndex).get(counter) + DB_Validation.myChar);
						counter++;
					} else {
						break;
					}
				}
				print.println();

			}
		} catch (Exception e) {
			System.out.println(e.getMessage());

		} finally {
			 if (print != null) {
		            print.close();
		        }
			actualRows.clear();
			rowStatus.clear();
			columns.clear();
			columnsData.clear();

		}

	}

	public static int count(File filename) throws IOException {
		InputStream is = new BufferedInputStream(new FileInputStream(filename));
		try {
			byte[] c = new byte[1024];
			int count = 0;
			int readChars = 0;
			boolean empty = true;
			while ((readChars = is.read(c)) != -1) {
				empty = false;
				for (int i = 0; i < readChars; ++i) {
					if (c[i] == '\n')
						++count;
				}
			}
			return (count == 0 && !empty) ? 1 : count;
		} finally {
			is.close();
		}
	}

}