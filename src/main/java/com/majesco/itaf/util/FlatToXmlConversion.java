package com.majesco.itaf.util;

import java.io.IOException;
import java.security.Permission;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.unidex.xflat.flat2xml;

public class FlatToXmlConversion {

	private final static Logger log = LogManager.getLogger(FlatToXmlConversion.class.getName());

	// Flat to xml converter
	public void convertFlatToXml(String copyTo, String fileName, String extn, String xflFile) {
		SystemExitControl.forbidSystemExitCall();
		try {
			SystemExitControl.forbidSystemExitCall();

			try {
				//flat2xml flatToXml = new flat2xml();//Sel4
				// String[] conData = {copyTo+"/"+xflFile,
				// copyTo+"/"+fileName,copyTo+"/"+fileName.replaceAll(extn,".xml")};
				String[] conData = { xflFile, copyTo + "/" + fileName,
						copyTo + "/" + fileName.replaceAll(extn, ".xml") }; // Meghna--To
				flat2xml.main(conData);
			} catch (SystemExitControl.ExitTrappedException e) {
				log.error(e.getMessage(), e);
				System.out.println("Forbidding call to System.exit");
			}
			SystemExitControl.enableSystemExitCall();
		} catch (IOException e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
		}
	}

}

class SystemExitControl {

	public static class ExitTrappedException extends SecurityException {

		/**
		 * 
		 */
		private static final long serialVersionUID = 8359259423865549922L;
	}

	@SuppressWarnings("removal")
	public static void forbidSystemExitCall() {
		@SuppressWarnings({ })
		final SecurityManager securityManager = new SecurityManager() {
			@Override
			public void checkPermission(Permission permission) {
				if (permission.getName().contains("exitVM")) {
					throw new ExitTrappedException();
				}
			}
		};
		System.setSecurityManager(securityManager);
	}

	@SuppressWarnings("removal")
	public static void enableSystemExitCall() {
		System.setSecurityManager(null);
	}
}