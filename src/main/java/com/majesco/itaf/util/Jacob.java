package com.majesco.itaf.util;

import java.io.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.LibraryLoader;
import com.jacob.com.Variant;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.ITAFWebDriver;

public class Jacob {
	private final static Logger log = LogManager.getLogger(Jacob.class.getName());
	private static boolean isDMApplication = ITAFWebDriver.isDMApplication();

	public static void main(String filepath, String macroname) {
		try {
			// Load the DLL file dynamically from the classpath
			InputStream inputStream = Jacob.class.getResourceAsStream("/jacob/" + getJacobDllFilename());

			if (inputStream == null) {
				String libFile;
				libFile = System.getProperty("user.dir") + "\\libs\\Jacob\\jacob-1.18-x64.dll";
				
				File dllFile = new File(libFile);

				if (!dllFile.exists()) {
					throw new FileNotFoundException("Could not find jacob DLL file in classpath or custom path.");
				}
				// Use the DLL from custom path
				System.setProperty(LibraryLoader.JACOB_DLL_PATH, dllFile.getAbsolutePath());

				LibraryLoader.loadJacobLibrary();
				@SuppressWarnings({ "unused" })
				ActiveXComponent compCLSID = new ActiveXComponent("clsid:{00024500-0000-0000-C000-000000000046}");

				log.info("The Library been loaded, and an activeX component been created");
				File file = new File(filepath);
				String macroName = macroname;
				callExcelMacro(file, macroName);
				// temporaryDll.deleteOnExit();
				// inputStream.close();
			} else {

				// Create temporary file for the DLL
				File temporaryDll = File.createTempFile("jacob", ".dll");
				FileOutputStream outputStream = new FileOutputStream(temporaryDll);
				byte[] buffer = new byte[8192];
				int bytesRead;
				while ((bytesRead = inputStream.read(buffer)) != -1) {
					outputStream.write(buffer, 0, bytesRead);
				}
				outputStream.close();

				// Set system property for Jacob DLL path
				System.setProperty(LibraryLoader.JACOB_DLL_PATH, temporaryDll.getAbsolutePath());
				LibraryLoader.loadJacobLibrary();
				@SuppressWarnings({ "unused" })
				ActiveXComponent compCLSID = new ActiveXComponent("clsid:{00024500-0000-0000-C000-000000000046}");
				log.info("The Library been loaded, and an activeX component been created");
				File file = new File(filepath);
				String macroName = macroname;
				callExcelMacro(file, macroName);
				temporaryDll.deleteOnExit();
				inputStream.close();
			}

		} catch (Exception e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
		}
	}

	private static String getJacobDllFilename() {
		String architecture = System.getProperty("os.arch");
		return "jacob-1.18-" + (architecture.contains("64") ? "x64" : "x86") + ".dll";
	}

	public static void callExcelMacro(File file, String macroName) {
		ComThread.InitSTA(true);
		final ActiveXComponent excel = new ActiveXComponent("Excel.Application");
		try {
			excel.setProperty("EnableEvents", new Variant(false));
			Dispatch workbooks = excel.getProperty("Workbooks").toDispatch();
			Dispatch workBook = Dispatch.call(workbooks, "Open", file.getAbsolutePath()).toDispatch();

			// Calls the macro
			Variant V1 = new Variant(file.getName() + macroName);
			String strFileName = file.getName();

			if ((isDMApplication)
					&& (!strFileName.equalsIgnoreCase("Ind_XML_COMPARISON_Utility_OpenAPIService_Linear.xlsm"))
					&& (!Config.executionStatusReportUtility.contains(strFileName))) {
				Variant result1 = Dispatch.call(excel, "Run", V1);
				log.info(result1.getString());
				log.info(strFileName + " macro executed successfully");
			} else {
				@SuppressWarnings("unused")
				Variant result = Dispatch.call(excel, "Run", macroName);
			}
			// Saves and closes
			Dispatch.call(workBook, "Save");
			com.jacob.com.Variant f = new com.jacob.com.Variant(true);
			Dispatch.call(workBook, "Close", f);

		} catch (Exception e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
		} finally {
			excel.invoke("Quit", new Variant[0]);
			ComThread.Release();
		}
	}

	// Other methods of your class...
}
