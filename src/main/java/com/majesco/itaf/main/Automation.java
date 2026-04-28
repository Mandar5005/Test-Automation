package com.majesco.itaf.main;

import java.io.File;
//import java.io.FileWriter;
import java.io.IOException;
import java.net.InetAddress;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.UnknownHostException;
//import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Duration;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
//import java.util.logging.Level;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.ie.InternetExplorerOptions;
//import org.openqa.selenium.logging.LogEntries;
//import org.openqa.selenium.logging.LogEntry;
//import org.openqa.selenium.logging.LogType;
//import org.openqa.selenium.logging.LoggingPreferences;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.safari.SafariDriver;
import org.openqa.selenium.safari.SafariOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Wait;
import com.majesco.itaf.util.CalendarSnippet;

//import org.openqa.selenium.remote.RemoteWebDriver;
//import java.net.URL;
import java.awt.Dimension;
import java.awt.Toolkit;

public class Automation {

	private final static Logger log = LogManager.getLogger(Automation.class.getName());
	public static WebDriver driver;
	public static String MainWebDriverHandle = null;

	public static enum browserTypeEnum {
		InternetExplorer, FireFox, Chrome, Safari, GeckoFireFox, MsEdge
	}

	private static String browser = null;
	public static browserTypeEnum browserType = null;
	private static ITAFWebDriver itafWebDriver = null;

	public static void setMainWebDriverHandle(String handle) {
		MainWebDriverHandle = handle;
	}

	public static String getMainWebDriverHandle() {
		return MainWebDriverHandle;
	}

	@SuppressWarnings("deprecation")
	public static void setUp() throws NullPointerException, Exception {

		int maxRetries = 3; // Maximum number of retries
		int attempt = 0;
		boolean setupSuccessful = false;

		while (attempt < maxRetries && !setupSuccessful) {
			attempt++;
			log.info("Setup attempt " + attempt + " of " + maxRetries);

			try {

				itafWebDriver = ITAFWebDriver.getInstance();
				// log.info( "iTAFWebDriver instance is:" + itafWebDriver);

				Date initialDate = new Date();
				String strInitialDate = Config.dtFormat.format(initialDate);
				itafWebDriver.getReport().setFromDate(strInitialDate);

				if (!ITAFWebDriver.isPASApplication()) {
					if (CalendarSnippet.isProcessRunning("iexplore.exe")) {
						CalendarSnippet.killProcess("iexplore.exe");
					}
				}
				if (CalendarSnippet.isProcessRunning("IEDriverServer.exe")) {
					CalendarSnippet.killProcess("IEDriverServer.exe");
				}

				// Commenting this to stop killing chrome driver of PAS
				// execution//
				if ((ITAFWebDriver.isSuiteApplication() || ITAFWebDriver.isPASApplication()
						|| ITAFWebDriver.isClaimsApplication() || ITAFWebDriver.isDMApplication()
						|| ITAFWebDriver.isLNAApplication() || ITAFWebDriver.isBillingApplication()
						|| ITAFWebDriver.isPLATApplication() || ITAFWebDriver.isCSApplication())
						&& Config.browserType.equalsIgnoreCase("Chrome")) {

					if (CalendarSnippet.isProcessRunning("chromedriver.exe")) {

						WebDriver driverInstance = itafWebDriver.getReport().getDriver();

						if (driverInstance != null) {
							log.info("Closing chromedriver.exe instance:" + "(" + driverInstance + ")" + "...");
							// driverInstance.close();
							driverInstance.quit(); // Quit the driver instance
							log.info("Closed successfully.");
						} // CalendarSnippet.killProcess("chromedriver.exe");

						// Below given code to be commented for parallel execution
						if (CalendarSnippet.isProcessRunning("chrome.exe")) {
							log.info("Closing all chrome browser instances...");
							CalendarSnippet.killProcess("chrome.exe");
							log.info("All chrome browser instances closed successfully.");
						} else {
							log.info("No chrome browser instance found...");
						}

					} else if (!ITAFWebDriver.isPASApplication()) {// Sel4 - 15/5/24
						if (CalendarSnippet.isProcessRunning("chrome.exe")) {
							log.info("Closing all chrome browser instances...");
							CalendarSnippet.killProcess("chrome.exe");
							log.info("All chrome browser instances closed successfully.");
						}
					}

				}

				if ((ITAFWebDriver.isSuiteApplication() || ITAFWebDriver.isPASApplication()
						|| ITAFWebDriver.isClaimsApplication() || ITAFWebDriver.isDMApplication()
						|| ITAFWebDriver.isLNAApplication() || ITAFWebDriver.isBillingApplication()
						|| ITAFWebDriver.isPLATApplication() || ITAFWebDriver.isCSApplication())
						&& Config.browserType.equalsIgnoreCase("MsEdge")) {

					if (CalendarSnippet.isProcessRunning("msedgedriver.exe")) {

						WebDriver driverInstance = itafWebDriver.getReport().getDriver();

						if (driverInstance != null) {
							log.info("Closing msedgedriver.exe instance:" + "(" + driverInstance + ")" + "...");
							// driverInstance.close();
							driverInstance.quit(); // Quit the driver instance
							log.info("Closed successfully.");
						} // CalendarSnippet.killProcess("chromedriver.exe");

						// Below given code to be commented for parallel execution
						if (CalendarSnippet.isProcessRunning("msedge.exe")) {
							log.info("Closing all msedge browser instances...");
							CalendarSnippet.killProcess("msedge.exe");
							log.info("All msedge browser instances closed successfully.");
						} else {
							log.info("No msedge browser instance found...");
						}

					} else if (!ITAFWebDriver.isPASApplication()) {// Sel4 - 15/5/24
						if (CalendarSnippet.isProcessRunning("msedge.exe")) {
							log.info("Closing all msedge browser instances...");
							CalendarSnippet.killProcess("msedge.exe");
							log.info("All msedge instances closed successfully.");
						}
					}

				}

				if (!ITAFWebDriver.isPASApplication()) {
					if (CalendarSnippet.isProcessRunning("firefox.exe")) {
						CalendarSnippet.killProcess("firefox.exe");
					}
				}

				try {

					browser = Config.browserType;
					browserType = browserTypeEnum.valueOf(browser);
					Object baseURL = Config.baseURL;

					URL gridUrlCh = null;

					switch (browserType) {

					case InternetExplorer:
						InternetExplorerOptions ieopt = new InternetExplorerOptions();
						ieopt.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS,
								true);
						driver = new InternetExplorerDriver(ieopt);
						itafWebDriver.getReport().setDriver(driver);
						driver.manage().deleteAllCookies();
						driver.manage().window().maximize();
						driver.manage().timeouts()
								.pageLoadTimeout(Duration.ofSeconds(Integer.parseInt(Config.timeOut)));
						;
						driver.get(baseURL.toString());
						driver.getWindowHandles();
						Automation.setMainWebDriverHandle(driver.getWindowHandle());

						try {
							if (driver.findElements(By.id("overridelink")).size() > 0) {
								Thread.sleep(5000);
								driver.navigate().to("javascript:document.getElementById('overridelink').click()");
								Thread.sleep(1000);
								driver.navigate().to("javascript:document.getElementById('overridelink').click()");
							}
						} catch (Exception e) {
							log.info("Error occured while handling override link: " + e.getMessage());
						}
						break;

					case MsEdge:

						if (Config.seleniumGrid != null && Config.seleniumGrid.equalsIgnoreCase("TRUE")) {
							try {
								// Get the local host's IP address
								InetAddress localHost = InetAddress.getLocalHost();
								String ipAddress = localHost.getHostAddress();

								// Define the protocol and port
								String protocol = "http";
								int port = 4444;

								// Construct the URL
								String gridUrlString = String.format("%s://%s:%d", protocol, ipAddress, port);
								gridUrlCh = new URL(gridUrlString);

								log.info("Selenium Grid URL: " + gridUrlCh);

							} catch (UnknownHostException e) {
								log.info("Failed to get the local host IP address: " + e.getMessage());
							} catch (MalformedURLException e) {
								log.info("Failed to construct the URL: " + e.getMessage());
							}
						} else {
							log.info("Selenium Grid not set");
						}

						EdgeOptions opt = new EdgeOptions();
						opt.addArguments("--guest");

						if (gridUrlCh != null) {
							driver = new RemoteWebDriver(gridUrlCh, opt); // Use the constructed grid URL
						} else {
							driver = new EdgeDriver(opt); // Fallback in case URL creation failed
						}
						// URL gridUrl = new URL("http://192.168.1.43:4444");//Sel4 - Testing
						// driver = new EdgeDriver(opt);
						// driver = new RemoteWebDriver(gridUrl, opt);//Sel4 - Testing

						log.info(" Webdriver instance is: " + itafWebDriver.getReport());
						log.info(" MsEdgeDriver instance is:" + driver);
						itafWebDriver.getReport().setDriver(driver);
						driver.manage().deleteAllCookies();
						driver.manage().window().maximize();
						driver.get(baseURL.toString());
						driver.getWindowHandles();
						Automation.setMainWebDriverHandle(driver.getWindowHandle());
						break;

					case FireFox:
						if (Config.seleniumGrid != null && Config.seleniumGrid.equalsIgnoreCase("TRUE")) {
							try {
								// Get the local host's IP address
								InetAddress localHost = InetAddress.getLocalHost();
								String ipAddress = localHost.getHostAddress();

								// Define the protocol and port
								String protocol = "http";
								int port = 4444;

								// Construct the URL
								String gridUrlString = String.format("%s://%s:%d", protocol, ipAddress, port);
								gridUrlCh = new URL(gridUrlString);

								log.info("Selenium Grid URL: " + gridUrlCh);

							} catch (UnknownHostException e) {
								log.info("Failed to get the local host IP address: " + e.getMessage());
							} catch (MalformedURLException e) {
								log.info("Failed to construct the URL: " + e.getMessage());
							}
						} else {
							log.info("Selenium Grid not set");
						}

						FirefoxOptions fox = new FirefoxOptions();
						fox.setBrowserVersion("stable");

						if (gridUrlCh != null) {
							driver = new RemoteWebDriver(gridUrlCh, fox); // Use the constructed grid URL
						} else {
							driver = new FirefoxDriver(fox); // Fallback in case URL creation failed
						}

						// driver = new FirefoxDriver(fox);
						itafWebDriver.getReport().setDriver(driver);
						driver.manage().deleteAllCookies();
						driver.manage().window().maximize();
						driver.navigate().to(baseURL.toString());
						driver.getWindowHandles();
						Automation.setMainWebDriverHandle(driver.getWindowHandle());
						break;

					case Chrome:

						if (Config.seleniumGrid != null && Config.seleniumGrid.equalsIgnoreCase("TRUE")) {
							try {
								// Get the local host's IP address
								InetAddress localHost = InetAddress.getLocalHost();
								String ipAddress = localHost.getHostAddress();

								// Define the protocol and port
								String protocol = "http";
								int port = 4444;

								// Construct the URL
								String gridUrlString = String.format("%s://%s:%d", protocol, ipAddress, port);
								gridUrlCh = new URL(gridUrlString);

								log.info("Selenium Grid URL: " + gridUrlCh);

							} catch (UnknownHostException e) {
								log.info("Failed to get the local host IP address: " + e.getMessage());
							} catch (MalformedURLException e) {
								log.info("Failed to construct the URL: " + e.getMessage());
							}
						} else {
							log.info("Selenium Grid not set");
						}

						// Enable Chrome performance logging
						// LoggingPreferences logPrefs = new LoggingPreferences();
						// logPrefs.enable(LogType.PERFORMANCE, Level.ALL); // Enable Performance
						// Logging
						


						ChromeOptions options = new ChromeOptions();
						// options.setCapability("goog:loggingPrefs", logPrefs); // Add logging
						// preferences
						
						//04/06/2025 - For billing - "cannot create temp dir for user data dir" issue.
						// Safe temp directory for Chrome profile to avoid VDI issues
						
						try {
						    File profileDir = new File(System.getProperty("java.io.tmpdir"), "chrome-profile-" + System.currentTimeMillis());
						    if (profileDir.mkdirs()) {
						        log.info("Temporary Chrome profile directory created at: " + profileDir.getAbsolutePath());
						        options.addArguments("--user-data-dir=" + profileDir.getAbsolutePath());
						    } else {
						        log.info("Failed to create temporary Chrome profile directory. Falling back to default.");
						    }
						} catch (Exception e) {
						    log.error("Exception while creating temp Chrome user data dir: " + e.getMessage(), e);
						}

						// ChromeOptions options = new ChromeOptions();
						HashMap<String, Object> chromeOptionsMap = new HashMap<String, Object>();
						chromeOptionsMap.put("plugins.plugins_disabled", new String[] { "Chrome PDF Viewer" });
						chromeOptionsMap.put("plugins.always_open_pdf_externally", true);
						chromeOptionsMap.put("credentials_enable_service", false);
						chromeOptionsMap.put("profile.password_manager_enabled", false);
						chromeOptionsMap.put("autofill.profile_enabled", false); // For claims address popup issue
																					// 18/04/2023
						// To disable file download pop-up
						chromeOptionsMap.put("download.default_directory", Config.runtimeFileDownloadFolder);
						chromeOptionsMap.put("profile.default_content_settings.popups", 0);
						chromeOptionsMap.put("safebrowsing.enabled", "true");
						// chromeOptionsMap.put("profile.default_content_setting_values.zoom_level",
						// 0.9);
						options.setExperimentalOption("prefs", chromeOptionsMap);
						options.addArguments("--ignore-certificate-errors");
						options.addArguments("--disable-popup-blocking");
						options.addArguments("--disable-translate");
						options.addArguments("--start-maximized");
						options.setExperimentalOption("detach", false);//
						// options.addArguments("--start-fullscreen");

						if ((Config.incognitoMode != null) && Config.incognitoMode.equalsIgnoreCase("TRUE")) {
							options.addArguments("--incognito"); // For PAS

						}

						if (Config.headlessBrowser != null && Config.headlessBrowser.equalsIgnoreCase("TRUE")) {
							options.addArguments("--headless=new"); // ref:
																	// https://github.com/SeleniumHQ/selenium/issues/11637
							Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
							int width = (int) screenSize.getWidth();
							int height = (int) screenSize.getHeight();
							String sizeArg = "window-size=" + width + "*" + height;
							System.out.println("Screen size: " + sizeArg);

							options.addArguments(sizeArg); //
							options.addArguments("--force-device-scale-factor=1"); // Use a 1x scaling factor
							options.addArguments("--disable-gpu"); // Recommended for headless mode
							options.addArguments("--no-sandbox"); // Bypass OS security model
							options.addArguments("--disable-dev-shm-usage"); // Overcome limited resource problems
						}
						options.setExperimentalOption("excludeSwitches",
								Collections.singletonList("enable-automation"));
						options.setExperimentalOption("useAutomationExtension", false);

						if (gridUrlCh != null) {
							driver = new RemoteWebDriver(gridUrlCh, options); // Use the constructed grid URL
						} else {
							driver = new ChromeDriver(options); // Fallback in case URL creation failed
						}

						itafWebDriver.getReport().setDriver(driver);
						driver.manage().deleteAllCookies();
						driver.manage().window().maximize();
						driver.get(baseURL.toString());
						driver.getWindowHandles();
						Automation.setMainWebDriverHandle(driver.getWindowHandle());

						// Retrieve performance logs after the driver is initialized
						/*
						 * LogEntries logs = driver.manage().logs().get(LogType.PERFORMANCE); try
						 * (FileWriter writer = new FileWriter("performanceLogs.txt", true)) { // Append
						 * mode for (LogEntry entry : logs) { String logMessage = "Performance Log: " +
						 * entry.getMessage(); // log.info(logMessage); // Print to console (optional)
						 * writer.write(logMessage + "\n"); // Write to file } //
						 * log.info("Performance logs written to performanceLogs.txt"); } catch
						 * (IOException e) { log.error("Failed to write performance logs: " +
						 * e.getMessage()); }
						 */
						break;

					case Safari:
						SafariOptions saf = new SafariOptions();
						driver = new SafariDriver(saf);
						itafWebDriver.getReport().setDriver(driver);
						driver.manage().window().maximize();
						driver.get(baseURL.toString());
						break;

					case GeckoFireFox:
						FirefoxOptions options1 = new FirefoxOptions();
						options1.setCapability("marionette", true);
						driver = new FirefoxDriver(options1);
						driver.navigate().to(baseURL.toString());
						break;

					}

				} catch (NullPointerException npe) {
					log.error("Failed create DriverInstance <-|-> LocalizeMessage " + npe.getLocalizedMessage()
							+ " <-|-> Message " + npe.getMessage() + " <-|-> Cause " + npe.getCause(), npe);
					getController().pauseFun("Failed create DriverInstance in Automation.setUp <-|-> LocalizeMessage "
							+ npe.getLocalizedMessage() + " <-|-> Message " + npe.getMessage() + " <-|-> Cause "
							+ npe.getCause());
				}

				// If no exception occurred, setup was successful
				setupSuccessful = true;
				log.info("Browser invoked successfully");

			} catch (Exception e) {
				Object baseURL = Config.baseURL;
				String errorMessage = e.getMessage();
				String firstLine = errorMessage.split("\\r?\\n")[0];

				log.error("Failed connecting to - " + baseURL + " - " + firstLine);
				log.error("Setup failed on attempt " + attempt + ": " + e.getMessage());
				if (attempt == maxRetries) {
					log.error("Max retries reached. Terminating execution...");
					System.exit(1); // Terminate if max retries are exhausted
					getController().pauseFun("Error from Automation.Setup " + firstLine);
				} else {
					log.info("Retrying setup...");
				}
				// log.error("Terminating Execution...");
				// System.exit(1); // You can use any status code you prefer
				// getController().pauseFun("Error from Automation.Setup " + firstLine);

			}

		}

	}

	private static MainController getController() {
		MainController controller = ObjectFactory.getMainController();
		return controller;
	}

	// For Refreshing login page if the page is not loaded//
	public static void refreshPage(String controlName) {
		Wait<WebDriver> waitLogin;
		@SuppressWarnings("unused")
		WebElement loginPage;
		Boolean loginLoaded = false;
		int loginCnt = 0;

		waitLogin = new FluentWait<>(driver).withTimeout(Duration.ofSeconds(4)).pollingEvery(Duration.ofSeconds(4))
				.ignoring(NoSuchElementException.class);

		while ((!loginLoaded) && (loginCnt < 5)) {// Sel4-25/7/24
			try {
				loginPage = waitLogin.until(ExpectedConditions.elementToBeClickable(By.xpath(controlName)));
				loginLoaded = true;
			} catch (Exception e) {
				driver.navigate().refresh();
				loginCnt = loginCnt + 1;
			}
		}

		if (loginLoaded) {
			log.info("Login Page loaded");
		}

	}

	/** : Get the temp directory path **/

	public static void clearTempDirectory() {
		try {
			log.info("clearTempDirectory Initiated...");
			String tempDirPath = System.getProperty("java.io.tmpdir");
			Path tempDir = Paths.get(tempDirPath);
			log.info("TempDir Path is: {}", tempDir);
			Path parentDir = tempDir.getParent();
			log.info("ParentDir Path is: {}", parentDir);

			// Check if the parent directory ends with "Temp", if not, append it
			if (parentDir != null && Files.exists(parentDir)) {
				String parentDirString = parentDir.toString();
				if (!parentDirString.endsWith("Temp")) {
					parentDir = parentDir.resolve("Temp");
					log.info("ParentDir Path is: {}", parentDir);
				}
			}

			if (Files.exists(parentDir) && Files.isDirectory(parentDir)) {
				log.info("Now deleting the files and folders in temp directory...");

				try (Stream<Path> paths = Files.walk(parentDir)) {
					List<Path> sortedPaths = paths
							.sorted(Comparator.comparing(Path::getFileName, Comparator.reverseOrder()))
							.collect(Collectors.toList());

					for (Path file : sortedPaths) {
						try {
							Files.delete(file);
							System.out.println("Deleted: " + file);
						} catch (IOException e) {
							// Attempt force delete if normal deletion fails
							try {
								FileUtils.forceDelete(file.toFile());
								System.out.println("Force deleted: " + file);
							} catch (IOException ex) {
								// System.out.println("Failed to delete file or directory: "+ file);
								System.out
										.println("\u001B[31mFailed to delete file or directory: " + file + "\u001B[0m");
							}
						}
					}
				}

				log.info("Temporary files cleared successfully.");
			} else {
				log.info("Temporary directory does not exist or is not a directory.");
			}
		} catch (IOException e) {
			log.error("Error occurred while clearing temp directory: {}", e.getMessage());
		}
	}

}
