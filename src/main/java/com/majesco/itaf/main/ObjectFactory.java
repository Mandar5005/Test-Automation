package com.majesco.itaf.main;

import java.io.IOException;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;


public class ObjectFactory {

	private static MainController mainController;
	private static ITAFWebDriver itafWebDriver;
	private final static Logger log = LogManager.getLogger(MainControllerClaims.class.getName());

	public static void initialize() throws IOException {
		
		if (ITAFWebDriver.isBillingApplication()) {
			itafWebDriver = new ITAFWebDriverBilling();
			mainController = new MainControllerBilling();
		
		} else if (ITAFWebDriver.isClaimsApplication()) {
			itafWebDriver = new ITAFWebDriverClaims();
			mainController = new MainControllerClaims();
		
		} else if (ITAFWebDriver.isPASApplication()) {
			itafWebDriver = new ITAFWebDriverPAS();
			mainController = new MainControllerPAS();
		
		} else if (ITAFWebDriver.isDMApplication()) {
			itafWebDriver = new ITAFWebDriverDM();
			mainController = new MainControllerDM();
		
		} else if (ITAFWebDriver.isLNAApplication()) {
			itafWebDriver = new ITAFWebDriverLNA();
			mainController = new MainControllerLNA();
		
		} else if (ITAFWebDriver.isPLATApplication()) {
			itafWebDriver = new ITAFWebDriverDS();
			mainController = new MainControllerDS();
		
		} else if (ITAFWebDriver.isSuiteApplication()) {
			itafWebDriver = new ITAFWebDriverSuite();
			mainController = new MainControllerSuite();
		
		} else if (ITAFWebDriver.isCSApplication()) {
			itafWebDriver = new ITAFWebDriverCS();
			mainController = new MainControllerCS();
		}
		itafWebDriver.init();
		log.info("itafWebDriver.init()initiated...");
	}

	public static MainController getMainController() {
		return mainController;
	}

	public static ITAFWebDriver getWebDriver() {
		return itafWebDriver;
	}

	public static void setWebDriver() {
		if (ITAFWebDriver.isPASApplication()) {
			itafWebDriver = new ITAFWebDriverPAS();
		} else if (ITAFWebDriver.isDMApplication()) {
			itafWebDriver = new ITAFWebDriverDM();
		} else if (ITAFWebDriver.isPLATApplication()) {
			itafWebDriver = new ITAFWebDriverDS();
		} else if (ITAFWebDriver.isCSApplication()) {
			itafWebDriver = new ITAFWebDriverCS();
		}

	}
}
