package com.majesco.itaf.util;

import java.io.IOException;
import java.util.Scanner;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.unidex.xflat.flat2xml;

public class ReportToFlatConversion {

	private final static Logger log = LogManager.getLogger(ReportToFlatConversion.class.getName());

	public void reportFileToXml(Scanner sc) {

		System.out.println("Enter the location where your report file can be found: ");
		String targetLocation = sc.nextLine();
		System.out.println("Enter the name of the report file with Extension: ");
		String fileName = sc.nextLine();
		System.out.println("Enter the xfl to be executed with Extension: ");
		String xflFile = sc.nextLine();
		ReportToFlatConversion rfc1 = new ReportToFlatConversion();
		rfc1.convertReportToXml(targetLocation, fileName, xflFile);
	}

	// .csv to xml converter
	public void convertReportToXml(String copyTo, String fileName, String xflFile) {
		SystemExitControl.forbidSystemExitCall();
		try {
			SystemExitControl.forbidSystemExitCall();

			try {
				/*
				 * flat2xml flatToXml = new flat2xml(); String[] conData = { copyTo + "/" +
				 * xflFile, copyTo + "/" + fileName, copyTo + "/" + fileName.replaceAll(".csv",
				 * ".xml") }; flatToXml.main(conData);
				 */
				// Sel4

				// Call the main method statically
				flat2xml.main(new String[] { copyTo + "/" + xflFile, copyTo + "/" + fileName,
						copyTo + "/" + fileName.replaceAll(".csv", ".xml") });

			} catch (SystemExitControl.ExitTrappedException e) {
				log.error(e.getMessage(), e);
				System.out.println("Forbidding call to System.exit");
			}
			SystemExitControl.enableSystemExitCall();
		} catch (IOException e) {
			log.error(e.getMessage(), e);
		}
	}

}
