package com.majesco.itaf.recovery;

import java.awt.image.BufferedImage;
import java.io.File;
import java.time.Duration;
import java.util.Date;
import java.util.Set;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.By;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import com.majesco.itaf.main.Automation;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.ITAFWebDriver;
import com.majesco.itaf.main.MainController;
import com.majesco.itaf.main.ObjectFactory;
import com.majesco.itaf.main.TransactionMapping;
import com.majesco.itaf.main.WebHelper;

public class StartRecoveryClaims {
	private final static Logger log = LogManager.getLogger(StartRecoveryClaims.class);
	private static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	private static MainController controller = ObjectFactory.getMainController();
	public static void initiateRecovery(String msg) throws InterruptedException {
		if(StringUtils.equalsIgnoreCase(Config.recovery_scenario,"true"))
		{
			log.error("Recovery Initiated to handle popups or screen messages....");
			log.error(msg);
			String AlertMsg = null;
			String webPagebodyText = null;
			Set<String> allWindowHandles = null;
			StartRecoveryClaims.takeScreenShot();
			WebDriver driver = Automation.driver;
			String mainWindowHandle = driver.getWindowHandle();
			String modalMsgText = isModalPopupPresent();
			if (isAlertPresent()) {
				AlertMsg = Automation.driver.switchTo().alert().getText().toString();
			} else if (modalMsgText != null
					&& (modalMsgText.toUpperCase().contains("SYSTEM ERROR") | modalMsgText.toUpperCase().contains("YOUR SEARCH RESULTED IN")
							| modalMsgText.toUpperCase().contains("NO DATA") | modalMsgText.toUpperCase().contains("ENTER VALID DATA"))) {
				AlertMsg = modalMsgText;
			}
	
			if (AlertMsg != null
					&& (AlertMsg.toUpperCase().contains("SYSTEM ERROR") | AlertMsg.toUpperCase().contains("YOUR SEARCH RESULTED IN")
							| AlertMsg.toUpperCase().contains("NO DATA") | modalMsgText.toUpperCase().contains("ENTER VALID DATA"))) {
				if (modalMsgText != null && AlertMsg.toUpperCase().contains("SYSTEM ERROR")) 
			
				{
					System.out.println("System error Handling_Claims");
					try {
						WebElement ErrMsg;
						ErrMsg = Automation.driver.findElement(By.xpath("//button[@class='btn btn-popup' and @name='OK']"));
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();", ErrMsg);
						ErrMsg.click();
					} catch (Exception e) {
						log.error(e.getMessage(), e);
					}
					// MainController.recoveryhandler();
				}
	
				if (modalMsgText != null && AlertMsg.toUpperCase().contains("NO DATA"))
				{
					System.out.println("System error/Blank Data Handling_Claims");
				}
	
				else if (modalMsgText != null) // 1
				{
					// Automation.driver.findElement(By.name("Ok")).click();
					System.out.println("inside 1");
					try {
						WebElement ErrMsg;
						ErrMsg = Automation.driver.findElement(By.xpath("//button[@class='btn btn-popup' and @name='OK']"));
						// ErrMsg = WebHelper.getElementByType("XPath",
						// ".//*[@class='btn btn-popup' and @name='Ok']","","","");
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();", ErrMsg);
						ErrMsg.click();
					} catch (Exception e) {
						log.error(e.getMessage(), e);
					}
	
					controller.recoveryhandler(); // for above 3 conditions recovery
					// scenario should get triggered.
				} 
				
			} else if (isAlertPresent()) {
				Automation.driver.switchTo().alert().accept();
			} else {
	
				try {
					// allWindowHandles =
					// webDriver.getReport().getDriver().getWindowHandles();
					allWindowHandles = driver.getWindowHandles();
	
					for (String currentWindowHandle : allWindowHandles) {
						if (!currentWindowHandle.equals(mainWindowHandle)) {
							driver.switchTo().window(currentWindowHandle);
							webPagebodyText = driver.findElement(By.tagName("body")).getText();
							System.out.println(webPagebodyText.toString());
							if (webPagebodyText.toString().toUpperCase().contains("SYSTEM ERROR OCCURERED")) {
								controller.recoveryhandler();
								try {
									Automation.setUp();
									TransactionMapping.TransactionInputData("Login"); // Default
									// Login
									// START
								} catch (Exception e) {
									log.error(e.getMessage(), e);
									System.out.println("Failed to login into application");
								}
								// Code to initiate Scenario block
							} else if (webPagebodyText.toString().toUpperCase().contains("SESSION TIMED OUT.")) {
								System.out.println("Session time out and login in again");
								try {
									Automation.setUp();
									TransactionMapping.TransactionInputData("Login"); // Default
									// Login
									// START
								} catch (Exception e) {
									log.error(e.getMessage(), e);
									System.out.println("Failed to login into application");
								}
								// Code to login into application again
							}
			
							driver.close();
						}
					}
					driver.switchTo().window(mainWindowHandle);
				} catch (WebDriverException e) {
					log.error("Exception occured while closing non-main windows  <-|-> LocalizeMessage " + e.getLocalizedMessage() + " <-|-> Message "
							+ e.getMessage() + " <-|-> Cause " + e.getCause(), e);
				}
			}
	}
	}

	// isAlertPresent() function check whether this pop is application pop-up or
	// not
	public static boolean isAlertPresent() {
		try {
			Automation.driver.switchTo().alert();
			return true;
		} catch (NoAlertPresentException e) {
			log.error(e.getMessage());
			return false;
		}
	}

	// isModalPresent() function check whether this modal class present or not
	public static String isModalPopupPresent() {
		try {
			WebElement modalMsgElement = Automation.driver.findElement(By.className("modal-message"));
			return modalMsgElement.getText();
		} catch (Exception e) {
			log.error(e.getMessage());
			return "NULL";
		}
	}

	public static void takeScreenShot() throws InterruptedException {

		@SuppressWarnings("unused")
		String cdate = null;

		@SuppressWarnings("unused")
		BufferedImage image = null;
		if (StringUtils.isNotBlank(webDriver.getReport().getFromDate()))
			cdate = webDriver.getReport().getFromDate().replaceAll("[-/: ]", "");
		else
			webDriver.getReport().setFromDate(Config.dtFormat.format(new Date()));
		String cfileName = "Claims_Pre_Recovery_SS_" + webDriver.getReport().getTestcaseId() + "_"
				+ webDriver.getReport().getTrasactionType();
		String clocation = Config.resultFilePath + "\\ScreenShots\\" + cfileName + "_" + WebHelper.screenshotnum
				+ ".png";
		WebHelper.screenshotnum = WebHelper.screenshotnum + 1;
		File scrFile;
		try {
			scrFile = ((TakesScreenshot) WebHelper.currentdriver).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(scrFile, new File(clocation));

		} catch (Exception e) {
			log.error("Exception thrown while taking screenshot in pop-recovery  <-|-> LocalizeMessage "
					+ e.getLocalizedMessage() + " <-|-> Message " + e.getMessage() + " <-|-> Cause " + e.getCause(), e);
		} finally {

			Automation.driver.navigate().refresh();
			Thread.sleep(3000);

			String claimsloader = "//div[@class='overlay']/div[@class='logo-wrapper']/div";
			//WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, 90);
			WebDriverWait WaitForPageLoad = new WebDriverWait(Automation.driver, Duration.ofSeconds(5));//Sel4
			WaitForPageLoad.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath(claimsloader)));

			try {
				Thread.sleep(2000);

				//WebDriverWait WaitForPageLoadWhtsNew = new WebDriverWait(Automation.driver, 15);
				WebDriverWait WaitForPageLoadWhtsNew = new WebDriverWait(Automation.driver, Duration.ofSeconds(5));//Sel4

				WebElement WhatsNewClose = WaitForPageLoadWhtsNew.until(ExpectedConditions
						.elementToBeClickable(By.xpath("//div[@class='modal-dialog']//button[@class='close']/span")));
				WhatsNewClose.click();
			} catch (Exception ex1) {
				log.info("No What's New Pop up found");
			}

			Thread.sleep(2000);
			WebElement webElementLogoutArrow = WaitForPageLoad.until(
					ExpectedConditions.elementToBeClickable(By.xpath("//div[@class='loginDetailSecDownArrow']")));
			Thread.sleep(1000);
			webElementLogoutArrow.click();
			WebElement webElementLogoutB = WaitForPageLoad
					.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[@id='logout']")));
			webElementLogoutB.click();
			WebElement webElementLogin = WaitForPageLoad
					.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Log in')]")));
			webElementLogin.click();
		}

	}

}
