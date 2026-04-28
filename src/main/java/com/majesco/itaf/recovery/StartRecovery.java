package com.majesco.itaf.recovery;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.Set;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import com.majesco.itaf.main.Automation;
import com.majesco.itaf.main.Config;
import com.majesco.itaf.main.ITAFWebDriver;
import com.majesco.itaf.main.MainController;
import com.majesco.itaf.main.ObjectFactory;
import com.majesco.itaf.main.TransactionMapping;
import com.majesco.itaf.main.WebHelper;
import com.majesco.itaf.main.WebHelperUtil;

public class StartRecovery {
	private final static Logger log = LogManager.getLogger(StartRecovery.class.getName());
	private static ITAFWebDriver webDriver = ITAFWebDriver.getInstance();
	private static MainController controller = ObjectFactory.getMainController();
	public static void initiateRecovery() {
		if(!ITAFWebDriver.isSuiteApplication())
		{
			String AlertMsg = null;
			String webPagebodyText = null;
			Set<String> allWindowHandles = null;
			StartRecovery.takeScreenShot();
			String mainWindowHandle = webDriver.getReport().getMainWindowHandle();
			WebDriver driver = webDriver.getReport().getDriver();
			String modalMsgText = isModalPopupPresent();
			if (isAlertPresent()) {
				AlertMsg = Automation.driver.switchTo().alert().getText().toString();
			}

			else if (modalMsgText != null)

			{
				AlertMsg = modalMsgText;
			}

			if (AlertMsg == null) {
				try {
					WebElement Popup = Automation.driver.findElement(By.xpath(".//*[@class='close']"));
					Thread.sleep(1000);
					WebHelperUtil.saveScreenShot();
					System.out.println("Popup:" + Popup);
					((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();", Popup);
					@SuppressWarnings("unused")
					WebElement Popupbutton = Automation.driver.findElement(By.xpath(".//*[@class='close']"));
					Popup.click();

				} catch (Exception e) {
					WebHelperUtil.saveScreenShot();
					//log.error(e.getMessage(), e);
					//System.out.println(e);.//Sel4
					//log.error(e);
				}
			}

			if (AlertMsg != null)

			{
				if (modalMsgText != null) {

					try {
						WebElement ErrMsg;
						ErrMsg = Automation.driver.findElement(
								By.xpath(".//*[@name='Ok' or @name='Yes' or @name='btnOk' or @class='btn btn-popup']"));
						Thread.sleep(1000);
						WebHelperUtil.saveScreenShot();
						log.info("Error message:" + modalMsgText);
						((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();", ErrMsg);
						@SuppressWarnings("unused")
						WebElement ErrMsgbutton = Automation.driver.findElement(
								By.xpath(".//*[@name='Ok' or @name='btnOk' or @name='Yes' or @class='btn btn-popup']"));
						ErrMsg.click();
						Thread.sleep(1000);
						modalMsgText = isModalPopupPresent();

						if ((modalMsgText != null
								&& (modalMsgText.toUpperCase().contains("SUBMISSION COMPLETED SUCCESSFULLY")))) {
							try {
								WebElement ErrMsg1;
								ErrMsg1 = Automation.driver.findElement(By.xpath(".//*[@name='Ok']"));
								Thread.sleep(1000);
								WebHelperUtil.saveScreenShot();
								log.info("Error message:" + modalMsgText);// Mandar
								((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
										ErrMsg1);
								@SuppressWarnings("unused")
								WebElement ErrMs1gbutton = Automation.driver
										.findElement(By.xpath(".//*[@name='Ok' or @name='btnOk']"));
								ErrMsg.click();
							} catch (Exception e) {
								//log.error(e.getMessage(), e);
								log.error(e);
							}
						}

						else {
							WebElement BackGroundScreen;
							BackGroundScreen = Automation.driver.findElement(By.xpath(".//*[@class='close']"));
							if (BackGroundScreen != null) {
								System.out.println("BackGroundScreen:" + BackGroundScreen);
								((JavascriptExecutor) Automation.driver).executeScript("arguments[0].scrollIntoView();",
										BackGroundScreen);
								@SuppressWarnings("unused")
								WebElement BackGroundScreenbutton = Automation.driver
										.findElement(By.xpath(".//*[@class='close']"));
								BackGroundScreen.click();
							}

						}

					} catch (Exception e) {
						//log.error(e.getMessage(), e);
						//String errorMessage = e.getMessage();
						String errorMessage = e.getMessage();
						String firstLine = errorMessage.split("\\R")[0];
						log.error("Error: " + firstLine);
					}
				} else {
					Automation.driver.switchTo().alert().accept();
					controller.recoveryhandler(); // for above 3 conditions recovery
					// scenario should get triggered.
					try {
						Automation.setUp();
						WebHelperUtil.saveScreenShot();
						WebHelper.valuesHeader.clear();
						TransactionMapping.TransactionInputData("Login"); 
					} catch (Exception e) {
						//log.error(e.getMessage(), e);
						log.error(e);
						System.out.println("Failed to login into application");
					}
				}
			} else if (isAlertPresent()) {
				Automation.driver.switchTo().alert().accept();
			} else {

				try {
					allWindowHandles = webDriver.getReport().getDriver().getWindowHandles();

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
									//log.error(e.getMessage(), e);
									log.error(e);
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
									//log.error(e.getMessage(), e);
									log.error(e);
									System.out.println("Failed to login into application");
								}
								// Code to login into application again
							}
							driver.close();
						}
					}
					driver.switchTo().window(mainWindowHandle);
				} catch (WebDriverException e) {
					log.error("Exception occured while closing non-main windows  <-|-> LocalizeMessage "
							+ e.getLocalizedMessage() + " <-|-> Message " + e.getMessage() + " <-|-> Cause " + e.getCause(),
							e);
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
			//log.error(e);
			String errorMessage = e.getMessage();
			String firstLine = errorMessage.split("\\r?\\n")[0];
			//log.error("Error :" + firstLine);
			log.error("\u001B[31mError: "+ "No Alert Msg Found " + firstLine + "\u001B[0m");
			return false;
		}
	}

	// isModalPresent() function check whether this modal class present or not
	public static String isModalPopupPresent() {
		try {
			WebElement modalMsgElement = Automation.driver.findElement(By.className("modal-message"));
			return modalMsgElement.getText();
		} catch (Exception e) {
			//log.error(e.getMessage(), e);
			//log.error(e);
			String errorMessage = e.getMessage();
			String firstLine = errorMessage.split("\\r?\\n")[0];
			//log.error("Error :" + firstLine);
			log.error("\u001B[31mError: "+ "No popup message found" + firstLine + "\u001B[0m");
			return null;
		}
	}

	public static void takeScreenShot() {
		String cdate = null;
		if (StringUtils.isNotBlank(webDriver.getReport().getFromDate()))
			cdate = webDriver.getReport().getFromDate().replaceAll("[-/: ]", "");
		else
			webDriver.getReport().setFromDate(Config.dtFormat.format(new Date()));
		String cfileName = "Captured_InRecovery_" + webDriver.getReport().getTestcaseId() + "_"
				+ webDriver.getReport().getTrasactionType() + "_" + cdate;
		String clocation = Config.resultFilePath + "\\ScreenShots\\" + cfileName + "_" + WebHelper.screenshotnum
				+ ".png";
		WebHelper.screenshotnum = WebHelper.screenshotnum + 1;
		try {
			
			//TakesScreenshot scrShot = ((TakesScreenshot)webDriver);
			TakesScreenshot scrShot = ((TakesScreenshot)Automation.driver);//SEl4
			File SrcFile = scrShot.getScreenshotAs(OutputType.FILE);
			
			 File DestFile = new File(clocation);
	            try {
	                FileUtils.copyFile(SrcFile, DestFile);
	            } catch (IOException e) {
	                e.printStackTrace();
	            }
			
			//image = new Robot().createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
			//ImageIO.write(image, "png", new File(clocation));
		} catch (Exception e) {
			log.error("Exception thrown while taking screenshot in pop-recovery  <-|-> LocalizeMessage "
					+ e.getLocalizedMessage() + " <-|-> Message " + e.getMessage() + " <-|-> Cause " + e.getCause(), e);
		}

	}

}
