package com.majesco.itaf.util;

import org.apache.logging.log4j.Level;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import com.majesco.itaf.main.Automation;

public class CommonExpectedConditions {

	private final static Logger log = LogManager.getLogger(ExpectedConditions.class.getName());

	public static ExpectedCondition<Boolean> valueToBePresentInElementsAttribute(final By locator,
			final String attribute, final String value) {

		return new ExpectedCondition<Boolean>() {
			public Boolean apply(WebDriver from) {
				try {
					String elementsAttributeValue = findElement(locator, from).getAttribute(attribute);
					return elementsAttributeValue.contains(value);
				} catch (StaleElementReferenceException e) {
					log.error(e.getMessage(), e);
					return null;
				}
			}

			@Override
			public String toString() {
				return String.format("value ('%s') to be present in element found by %s", value, locator);
			}
		};
	}

	public static ExpectedCondition<Boolean> valueToBePresentInElementsAttribute(final WebElement element,
			final String attribute, final String value) {

		return new ExpectedCondition<Boolean>() {
			public Boolean apply(WebDriver from) {
				try {
					String elementsAttributeValue = element.getAttribute(attribute);
					return elementsAttributeValue.contains(value);
				} catch (StaleElementReferenceException e) {
					log.error(e.getMessage(), e);
					return null;
				}
			}

			@Override
			public String toString() {
				return String.format("value ('%s') to be present in element found by %s", value, element.getTagName());
			}
		};
	}

	public static ExpectedCondition<Boolean> valueToNotBePresentInElementsAttribute(final By locator,
			final String attribute, final String value) {

		return new ExpectedCondition<Boolean>() {
			public Boolean apply(WebDriver from) {
				try {
					String elementsAttributeValue = findElement(locator, from).getAttribute(attribute);
					return !elementsAttributeValue.contains(value);
				} catch (StaleElementReferenceException e) {
					log.error(e.getMessage(), e);
					return null;
				}
			}

			@Override
			public String toString() {
				return String.format("value ('%s') to be present in element found by %s", value, locator);
			}
		};
	}

	public static ExpectedCondition<Boolean> classRemovedFromElement(final WebElement element, final String className) {
		return new ExpectedCondition<Boolean>() {
			public Boolean apply(WebDriver from) {
				String classString = element.getAttribute("class");
				boolean hasClass = classString.matches("^" + className + "$")
						|| classString.matches("\\s" + className + "$") || classString.matches("^" + className + "\\s")
						|| classString.matches("\\s" + className + "\\s");
				return !hasClass;
			}
		};
	}

	public static ExpectedCondition<Boolean> givenStringtoBePresentInURL(final String givenString) {

		// driver = DriverProvider.getWebDriver();
		return new ExpectedCondition<Boolean>() {
			public Boolean apply(WebDriver d) {
				Boolean contains;
				contains = Automation.driver.getCurrentUrl().contains(givenString);
				return contains;
			}
		};
	}

	public static ExpectedCondition<WebElement> elementToBeClickable(final WebElement GivenElement) {
		return new ExpectedCondition<WebElement>() {

			public ExpectedCondition<WebElement> visibilityOfElement = ExpectedConditions.visibilityOf(GivenElement);

			public WebElement apply(WebDriver driver) {
				WebElement element = visibilityOfElement.apply(driver);
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
				return "element to be clickable: " + GivenElement.getTagName();
			}
		};
	}

	public static ExpectedCondition<WebElement> elementNotToBeClickable(final WebElement GivenElement) {
		return new ExpectedCondition<WebElement>() {

			public ExpectedCondition<WebElement> visibilityOfElement = ExpectedConditions.visibilityOf(GivenElement);

			public WebElement apply(WebDriver driver) {
				WebElement element = visibilityOfElement.apply(driver);
				try {
					if (element != null && !element.isEnabled()) {
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
				return "element to be clickable: " + GivenElement.getTagName();
			}
		};
	}

	/**
	 * An expectation for checking if the given text is present in the specified
	 * element.
	 */
	public static ExpectedCondition<Boolean> textToBePresentInElement(final WebElement GivenElement,
			final String text) {

		return new ExpectedCondition<Boolean>() {
			public Boolean apply(WebDriver from) {
				try {
					String elementText = GivenElement.getText();
					return elementText.contains(text);
				} catch (StaleElementReferenceException e) {
					log.error(e.getMessage(), e);
					return null;
				}
			}

			@Override
			public String toString() {
				return String.format("text ('%s') to be present in element %s", text, GivenElement.getTagName());
			}
		};
	}

	private static WebElement findElement(By by, WebDriver driver) {
		try {
			return driver.findElement(by);
		} catch (NoSuchElementException e) {
			throw e;
		} catch (WebDriverException e) {
			log.log(Level.WARN, String.format("WebDriverException thrown by findElement(%s)", by), e);// Level1 updated
																										// - log4J
																										// vulnerabilities
																										// - 27/01/2023
			throw e;
		}
	}

	private static boolean findInvisibleElement(By by, WebDriver driver) {
		try {
			driver.findElement(by);
			return false;
		} catch (NoSuchElementException e) {
			log.error(e.getMessage(), e);
			return true;
		}
	}

	public static ExpectedCondition<Boolean> invisibilityOfElementLocated(final By locator) {
		return new ExpectedCondition<Boolean>() {
			public Boolean apply(WebDriver driver) {
				try {
					return (findInvisibleElement(locator, driver));
				} catch (StaleElementReferenceException e) {
					log.error(e.getMessage(), e);
					return true;
				}
			}

			@Override
			public String toString() {
				return "element to no longer be visible: " + locator;
			}
		};
	}

}