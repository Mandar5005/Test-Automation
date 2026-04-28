package com.majesco.itaf.util;

import java.io.File;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.ie.InternetExplorerOptions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.safari.SafariDriver;
import org.openqa.selenium.safari.SafariOptions;
import com.majesco.itaf.main.Config;


public class WebDriverFactory {

	private final static Logger log = LogManager.getLogger(WebDriverFactory.class.getName());
	private static WebDriverFactory instance;
	private static final String CHROME_DRIVER_LOCATION = "webdriver.chrome.driver";
	private static final String CHROME_DRIVER_NOT_DEFINED = "notDefined";
	private final Map<String, WebDriverCreator> drivers = new HashMap<String, WebDriverCreator>();


	public static WebDriverFactory getInstance() {
		if (instance == null) {
			instance = new WebDriverFactory();
		}
		return instance;
	}

	private boolean isChromeDriverInstalled() {
		return System.getProperty(CHROME_DRIVER_LOCATION) != CHROME_DRIVER_NOT_DEFINED;
	}

	@SuppressWarnings("unused")
	private void initializeWebDriverCreators() {
		drivers.put("geckodriver", new GeckoDriverCreator());
		drivers.put("msie", new IEWebDriverCreator());
		drivers.put("chrome", new ChromeWebDriverCreator());
		drivers.put("edge", new EdgeWebDriverCreator());
		drivers.put("safari", new SafariWebDriverCreator());
		drivers.put("chromium", new ChromiumWebDriverCreator());
		drivers.put("saucelab", new SauceLabWebDriverCreator());
		drivers.put("ghost", new GhostWebDriverCreator());
		drivers.put("grid", new GridWebDriverCreator());
		drivers.put("remotegrid", new RemoteGridWebDriverCreator());
	}
	public WebDriver createDriver(String driverName) {
	    if (drivers.containsKey(driverName)) {
	        return drivers.get(driverName).create();
	    } else {
	        FirefoxOptions options = new FirefoxOptions();
	        WebDriver fireFoxDriver = new FirefoxDriver(options);
	        return fireFoxDriver;
	    }
	}
	
	private interface WebDriverCreator {
		WebDriver create();
	}

	private class IEWebDriverCreator implements WebDriverCreator {
	    public WebDriver create() {
	        InternetExplorerOptions ieOptions = new InternetExplorerOptions();
	        ieOptions.setCapability("unexpectedAlertBehaviour", "ignore");
	        ieOptions.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS, true);
	        ieOptions.setCapability("ignoreZoomSetting", true); // set zoom level to 100%
	        //ieOptions.setJavascriptEnabled(true);//removed in Sel4
	        ieOptions.setCapability("browserstack.ie.enablePopups", "true");
	        ieOptions.setCapability("acceptSslCerts", "true");
	        //ieOptions.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
	        ieOptions.setCapability(CapabilityType.ACCEPT_INSECURE_CERTS, true);
	        File file = new File(System.getProperty("user.dir") + "/libs/IEDriverServer.exe");
	        System.setProperty("webdriver.ie.driver", file.getAbsolutePath());
	        return new InternetExplorerDriver(ieOptions);
	    }
	}
	

	// Initialize Chrome Driver
	private class ChromeWebDriverCreator implements WebDriverCreator {
		@Override
		public WebDriver create() {
			
			ChromeOptions options = new ChromeOptions();
			HashMap<String, Object> chromeOptionsMap = new HashMap<String, Object>();
			chromeOptionsMap.put("plugins.plugins_disabled", new String[] { "Chrome PDF Viewer" });
			chromeOptionsMap.put("plugins.always_open_pdf_externally", true);
			chromeOptionsMap.put("credentials_enable_service", false);
			chromeOptionsMap.put("profile.password_manager_enabled", false);
			chromeOptionsMap.put("autofill.profile_enabled", false);//For claims address popup issue 18/04/2023

			// To disable file download pop-up
			chromeOptionsMap.put("download.default_directory", Config.runtimeFileDownloadFolder);
			chromeOptionsMap.put("profile.default_content_settings.popups", 0);
			chromeOptionsMap.put("safebrowsing.enabled", "true");
			options.setExperimentalOption("prefs", chromeOptionsMap);
			// options.addArguments("--disable-infobars");// This call is extra in
			// PAS
			//options.AddUserProfilePreference("autofill.profile_enabled", false);
			options.addArguments("--ignore-certificate-errors");
			options.addArguments("--disable-popup-blocking");
			options.addArguments("--disable-translate");
			options.addArguments("--start-maximized");

			if (Config.headlessBrowser != null && Config.headlessBrowser.equalsIgnoreCase("TRUE")) {
				options.addArguments("window-size=1366,768");
				//options.addArguments("headless");
				options.addArguments("--headless=new"); // ref: https://github.com/SeleniumHQ/selenium/issues/11637
			}
			options.setExperimentalOption("excludeSwitches", Collections.singletonList("enable-automation"));
			options.setExperimentalOption("useAutomationExtension", false);
			return new ChromeDriver(options);
		}
	}

	// Initialize Ms Edge Driver
	private class EdgeWebDriverCreator implements WebDriverCreator {
		@Override
		public WebDriver create() {

			File file = new File(System.getProperty("user.dir") + "/libs/msedgedriver.exe");
			System.setProperty("webdriver.edge.driver", file.getAbsolutePath());
			//Auto-upgrade chromedriver at runtime through Driver Manager
			try {
				//WebDriverManager.edgedriver().setup();
			} catch (Exception wdmexc){
				System.out.println("Edgedriver has not been downloaded online... " + wdmexc.getMessage());
				log.info("Edgedriver has not been downloaded online... " + wdmexc.getMessage());				
				if (!isChromeDriverInstalled()) {
					throw new RuntimeException("Edge driver is not installed");
				}
			}
			//return new ChromeDriver(options);
			return new EdgeDriver();
		}
	}

	// Initialize Safari Driver
	private class SafariWebDriverCreator implements WebDriverCreator {
		@Override
		public WebDriver create() {
			SafariOptions safariOptions = new SafariOptions();
			return new SafariDriver(safariOptions);
		}
	}

	// Initialize Chromium Driver
	private class ChromiumWebDriverCreator implements WebDriverCreator {
		@Override
		public WebDriver create() {
			if (!isChromeDriverInstalled()) {
				throw new RuntimeException("Chrome/Chromium driver not installed");
			}
			ChromeOptions options = new ChromeOptions();
			options.setBinary(new File("/usr/bin/chromium-browser"));
			options.addArguments("--ignore-certificate-errors");
			options.addArguments("--disable-popup-blocking");
			options.addArguments("--disable-translate");
			options.addArguments("--start-maximized");
			return new ChromeDriver(options);
		}
	}
	private class GridWebDriverCreator implements WebDriverCreator {
		@Override
		public WebDriver create() {
			
			
			FirefoxOptions options = new FirefoxOptions();
			FirefoxProfile profile = new FirefoxProfile();
		    options.setProfile(profile);
			WebDriver driver = null;
			try {
				driver = new RemoteWebDriver(new URL("http://localhost:4444/wd/hub"), options);
			} catch (MalformedURLException e) {
				log.error(e.getMessage(), e);
			}
			return driver;
		}
	}

	private class RemoteGridWebDriverCreator implements WebDriverCreator {

		@Override
		public WebDriver create() {
			FirefoxOptions options = new FirefoxOptions();
			FirefoxProfile profile = new FirefoxProfile();
			options.setProfile(profile);
			WebDriver driver = null;
			try {
				driver = new RemoteWebDriver(new URL("http://10.48.11.144:4444/wd/hub"), options);//selenium4
			} catch (MalformedURLException e) {
				log.error(e.getMessage(), e);
			}
			return driver;
		}

	}
	
	private class SauceLabWebDriverCreator implements WebDriverCreator {
		@Override
		public WebDriver create() {
			FirefoxOptions options = new FirefoxOptions();
			options.setCapability("version", "16");//Selenium4
			options.setCapability("platform", org.openqa.selenium.Platform.XP);//Selenium4
			options.setCapability("name", "TalentLink on Sauce");//Selenium4
			try {
				URL sauceURL = new URL(
						"http://michaldec:48beee6d-1691-457d-9ac1-af71347faf9d@ondemand.saucelabs.com:80/wd/hub");
				return new RemoteWebDriver(sauceURL, options);//Selenium4
			} catch (MalformedURLException e) {
				log.error(e.getMessage(), e);
				throw new RuntimeException(e);
			}
		}
	}
	
	private class GhostWebDriverCreator implements WebDriverCreator {
		public WebDriver create() {
			String ghostDriverUrl = System.getProperty("ghost.url", "http://localhost:4444");
			ChromeOptions options = new ChromeOptions();
			try {
				return new RemoteWebDriver(new URL(ghostDriverUrl), options);
			} catch (MalformedURLException e) {
				log.error(e.getMessage(), e);
				throw new RuntimeException(e);
			}
		}
	}
	
	private class GeckoDriverCreator implements WebDriverCreator {
		@Override
		public WebDriver create() {
			System.setProperty("webdriver.gecko.driver", "D:/geckodriver1/geckodriver.exe");
			FirefoxOptions options = new FirefoxOptions();
			options.setCapability("marionette", true);
			return new FirefoxDriver(options);
		}
	}

}
