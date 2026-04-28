package com.majesco.itaf.util;

import java.io.IOException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import com.unidex.xflat.xml2flat;

public class XmlToFlatConversion {
	private final static Logger log = LogManager.getLogger(XmlToFlatConversion.class.getName());
	public static void convertXmlToFlat(String copyFromLocal, String copyToRemote, String xmlFileName, String extn,
			String xflFile) {
		SystemExitControl.forbidSystemExitCall();
		try {
			SystemExitControl.forbidSystemExitCall();

			try {
				//xml2flat xmlToFlat = new xml2flat();//Sel4
				// String[] conData = {copyFromLocal+"/"+xflFile,
				// copyFromLocal+"/"+xmlFileName,copyFromLocal+"/"+xmlFileName.replaceAll(".xml",extn)};
				
				/*
				 * String[] conData = { xflFile, copyFromLocal + "/" + xmlFileName,
				 * copyFromLocal + "/" + xmlFileName.replaceAll(".xml", extn) };
				 * xmlToFlat.main(conData);
				 *///Sel4
				
				xml2flat.main(new String[] { xflFile, copyFromLocal + "/" + xmlFileName,
	                    copyFromLocal + "/" + xmlFileName.replaceAll(".xml", extn) });
				
			} catch (SystemExitControl.ExitTrappedException e) {
				System.out.println("Forbidding call to System.exit");
			}
			SystemExitControl.enableSystemExitCall();
		} catch (IOException e) {
			log.error(e.getMessage(), e);
			e.printStackTrace();
		}
	}

}
