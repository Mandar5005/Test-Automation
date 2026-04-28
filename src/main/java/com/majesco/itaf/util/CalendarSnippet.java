package com.majesco.itaf.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.text.DateFormatSymbols;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class CalendarSnippet {

	private final static Logger log = LogManager.getLogger(CalendarSnippet.class.getName());
	private static final String TASKLIST = "tasklist";
	private static String KILL = "";
	public static String getMonthForInt(int monthInt) {
		String monthName = "invalid";
		monthInt = monthInt - 1;
		DateFormatSymbols dfs = new DateFormatSymbols();
		String[] months = dfs.getMonths();
		if (monthInt >= 0 && monthInt <= 11) {
			monthName = months[monthInt];
		}
		return monthName;
	}
	public static String getShortMonthForInt(int monthInt) {
		String monthName = "invalid";
		monthInt = monthInt - 1;
		DateFormatSymbols dfs = new DateFormatSymbols();
		String[] months = dfs.getShortMonths();
		if (monthInt >= 0 && monthInt <= 11) {
			monthName = months[monthInt];
		}
		return monthName;
	}

	public static String getMonthForString(String monthName) throws ParseException {
		Calendar cal = Calendar.getInstance();
		cal.setTime(new SimpleDateFormat("MMM").parse("december"));
		NumberFormat nf = NumberFormat.getInstance();
		nf.setMinimumIntegerDigits(2);
		String monthInt = nf.format(cal.get(Calendar.MONTH) + 1);

		return monthInt;
	}


	public static boolean isProcessRunning(String serviceName) throws Exception {

		Process p = Runtime.getRuntime().exec(TASKLIST);
		BufferedReader reader = null;
		try {
			reader = new BufferedReader(new InputStreamReader(p.getInputStream()));
		} catch (Exception e) {
			log.error("Failed to read process list <-|-> LocalizeMessage " + e.getLocalizedMessage() + " <-|-> Message "
					+ e.getMessage() + " <-|-> Cause " + e.getCause(), e);
			throw new Exception("Failed to read process list  <-|-> LocalizeMessage " + e.getLocalizedMessage()
			+ " <-|-> Message" + e.getMessage() + " <-|-> Cause " + e.getCause());
		}
		String line;
		while ((line = reader.readLine()) != null) {
			// log.info(line); // commnented for log clearing
			if (line.contains(serviceName)) {
				return true;
			}
		}
		return false;
	}
	


	
	public static void killProcess(String serviceName) throws Exception {
		KILL = "\\System32\\taskkill /F /IM ";
		KILL = System.getenv("SystemRoot") + KILL;
		try {
			Runtime.getRuntime().exec(KILL + serviceName);
		} catch (Exception e) {
			log.error("Failed to kill process <-|-> LocalizeMessage " + e.getLocalizedMessage() + " <-|-> Message "
					+ e.getMessage() + " <-|-> Cause " + e.getCause(), e);
			throw new Exception("Failed to kill process <-|-> LocalizeMessage " + e.getLocalizedMessage()
					+ " <-|-> Message " + e.getMessage() + " <-|-> Cause " + e.getCause());
		}
	}
	 

}
