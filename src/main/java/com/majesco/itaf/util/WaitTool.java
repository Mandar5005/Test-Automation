package com.majesco.itaf.util;

import java.time.Duration;
import java.util.List;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class WaitTool {

	private final static Logger log = LogManager.getLogger(WaitTool.class.getName());
	public static final int DEFAULT_WAIT_4_ELEMENT = 2;
	public static final int DEFAULT_WAIT_4_PAGE = 2;

	public static WebElement waitForElement(WebDriver driver, final By by, int timeOutInSeconds) {
		WebElement element;
		try {
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(0));
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(timeOutInSeconds));
			element = wait.until(ExpectedConditions.visibilityOfElementLocated(by));
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(DEFAULT_WAIT_4_PAGE));
			return element;
		} catch (Exception e) {
			log.error(e.getMessage(), e);
		}
		return null;
	}

	public static WebElement waitForElementByWebElement(WebDriver driver, WebElement element, int timeOutInSeconds) {

		try {
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(0));
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(timeOutInSeconds));
			element = wait.until(elementToBeClickable(element, ""));
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(DEFAULT_WAIT_4_PAGE));
			return element;
		} catch (Exception e) {
			log.error(e.getMessage(), e);
		}
		return null;
	}

	public static WebElement waitForElementPresent(WebDriver driver, final By by, int timeOutInSeconds) {
		WebElement element;
		try {
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(0));
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(timeOutInSeconds));
			element = wait.until(ExpectedConditions.presenceOfElementLocated(by));
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(DEFAULT_WAIT_4_PAGE));
			return element;
		} catch (Exception e) {
			log.error(e.getMessage(), e);
		}
		return null;
	}

	public static List<WebElement> waitForListElementsPresent(WebDriver driver, final By by, int timeOutInSeconds) {
		List<WebElement> elements;
		try {
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(0));
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(timeOutInSeconds));
			wait.until((new ExpectedCondition<Boolean>() {
				@Override
				public Boolean apply(WebDriver driverObject) {
					return areElementsPresent(driverObject, by);
				}
			}));

			elements = driver.findElements(by);
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(DEFAULT_WAIT_4_PAGE));
			return elements;
		} catch (Exception e) {
			log.error(e.getMessage(), e);
		}
		return null;
	}

	public static WebElement waitForElementRefresh(WebDriver driver, final By by, int timeOutInSeconds) {
		WebElement element;
		try {
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(0));
			new WebDriverWait(driver, Duration.ofSeconds(timeOutInSeconds)).until((WebDriver driverObject) -> {
				driverObject.navigate().refresh(); // Refresh the page
				return isElementPresentAndDisplay(driverObject, by);
			});
			element = driver.findElement(by);
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(DEFAULT_WAIT_4_PAGE));
			return element;
		} catch (Exception e) {
			log.error(e.getMessage(), e);
		} finally {
			// Ensure to reset implicitlyWait in case of an exception //Added finally block
			// for Sel 4
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(DEFAULT_WAIT_4_PAGE));
		}
		return null;
	}

	public static boolean waitForTextPresent(WebDriver driver, final By by, final String text, int timeOutInSeconds) {
		boolean isPresent = false;
		try {

			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(0));
			new WebDriverWait(driver, Duration.ofSeconds(timeOutInSeconds)).until((WebDriver driverObject) -> {
				return isTextPresent(driverObject, by, text);
			});

			isPresent = isTextPresent(driver, by, text);
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(DEFAULT_WAIT_4_PAGE));
			return isPresent;
		} catch (Exception e) {
			log.error(e.getMessage(), e);
		} finally {
			// Ensure to reset implicitlyWait in case of an exception/Added finally block
			// for Sel 4
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(DEFAULT_WAIT_4_PAGE));
		}
		return false;
	}

	public static boolean waitForJavaScriptCondition(WebDriver driver, final String javaScript, int timeOutInSeconds) {
		boolean jscondition = false;
		try {
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(0));
			new WebDriverWait(driver, Duration.ofSeconds(timeOutInSeconds)).until((WebDriver driverObject) -> {
				return (Boolean) ((JavascriptExecutor) driverObject).executeScript(javaScript);
			});
			jscondition = (Boolean) ((JavascriptExecutor) driver).executeScript(javaScript);
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(DEFAULT_WAIT_4_PAGE));
			return jscondition;
		} catch (Exception e) {
			log.error(e.getMessage(), e);
		} finally {
			// Ensure to reset implicitlyWait in case of an exception// Added finally block
			// for sel4
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(DEFAULT_WAIT_4_PAGE));
		}
		return false;
	}

	public static boolean waitForJQueryProcessing(WebDriver driver, int timeOutInSeconds) {
		boolean jQcondition = false;
		try {
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(0));
			new WebDriverWait(driver, Duration.ofSeconds(timeOutInSeconds)).until((WebDriver driverObject) -> {
				return (Boolean) ((JavascriptExecutor) driverObject).executeScript("return jQuery.active == 0");
			});
			jQcondition = (Boolean) ((JavascriptExecutor) driver).executeScript("return jQuery.active == 0");
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(DEFAULT_WAIT_4_PAGE));
			return jQcondition;
		} catch (Exception e) {
			log.error(e.getMessage(), e);
		} finally {
			// Ensure to reset implicitlyWait in case of an exception//Added finally block -
			// sel4
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(DEFAULT_WAIT_4_PAGE));
		}
		return jQcondition;
	}

	public static ExpectedCondition<WebElement> elementToBeClickable(final WebElement webElement,
			final String ControlId) {
		try {
			return new ExpectedCondition<WebElement>() {

				public ExpectedCondition<WebElement> visibilityOfElementLocated = visibilityOfElementLocated(webElement,
						ControlId);

				@Override
				public WebElement apply(WebDriver driver) {
					WebElement element = visibilityOfElementLocated.apply(driver);
					try {
						if (element != null && element.isEnabled()) {
							return element;
						} else {
							return null;
						}
					} catch (StaleElementReferenceException e) {
						log.error(e.getMessage(), e);
						return null;
					}
				}

				@Override
				public String toString() {
					return "element to be clickable: " + ControlId;
				}
			};
		} catch (Exception e) {
			log.error(e.getMessage(), e);
		}
		return null;
	}

	public static ExpectedCondition<WebElement> visibilityOfElementLocated(final WebElement webElement,
			final String ControlId) {
		return new ExpectedCondition<WebElement>() {
			@Override
			public WebElement apply(WebDriver driver) {
				try {
					return elementIfVisible(webElement);
				} catch (StaleElementReferenceException e) {
					log.error(e.getMessage(), e);
					return null;
				}
			}

			@Override
			public String toString() {
				return "visibility of element located by " + ControlId;
			}
		};
	}

	private static WebElement elementIfVisible(WebElement element) {
		return element.isDisplayed() ? element : null;
	}

	public static void nullifyImplicitWait(WebDriver driver) {
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(0));

	}

	public static void setImplicitWait(WebDriver driver, int waitTime_InSeconds) {
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(waitTime_InSeconds));
	}

	public static void resetImplicitWait(WebDriver driver) {
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(0));
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(DEFAULT_WAIT_4_PAGE));
	}

	public static void resetImplicitWait(WebDriver driver, int newWaittime_InSeconds) {
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(0));
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(newWaittime_InSeconds));
	}

	private static boolean isTextPresent(WebDriver driver, By by, String text) {
		try {
			return driver.findElement(by).getText().contains(text);
		} catch (NullPointerException e) {
			log.error(e.getMessage(), e);
			return false;
		}
	}

	private static boolean areElementsPresent(WebDriver driver, By by) {
		try {
			driver.findElements(by);
			return true;
		} catch (NoSuchElementException e) {
			log.error(e.getMessage(), e);
			return false;
		}
	}

	private static boolean isElementPresentAndDisplay(WebDriver driver, By by) {
		try {
			return driver.findElement(by).isDisplayed();
		} catch (NoSuchElementException e) {
			log.error(e.getMessage(), e);
			return false;
		}
	}

}