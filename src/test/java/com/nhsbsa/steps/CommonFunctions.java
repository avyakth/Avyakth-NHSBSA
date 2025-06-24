package com.nhsbsa.steps;

import static io.restassured.RestAssured.given;
import org.json.JSONObject;
import com.aventstack.extentreports.markuputils.CodeLanguage;
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.jayway.jsonpath.JsonPath;
import com.nhsbsa.util.ExcelUtilResponse;
import com.nhsbsa.util.ExcelWriterResponseUpdation;
import com.nhsbsa.util.ExtentReportListener;
import com.nhsbsa.util.ITestListenerImpl;
import com.nhsbsa.util.MyLogger;
import com.nhsbsa.util.Reportgenerator;
import com.nhsbsa.util.ResponseToRequest;

import io.restassured.response.Response;
import io.restassured.response.ValidatableResponse;
import io.restassured.specification.RequestSpecification;
import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Properties;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.TimeoutException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import org.apache.commons.codec.binary.Base64;
import org.apache.commons.io.comparator.LastModifiedFileComparator;
import org.apache.commons.io.filefilter.WildcardFileFilter;
import org.apache.commons.validator.GenericValidator;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Capabilities;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.Point;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.UnexpectedAlertBehaviour;
import org.openqa.selenium.UnhandledAlertException;
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
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Optional;
import org.testng.annotations.Parameters;
import com.aventstack.extentreports.ExtentTest;
import com.opencsv.CSVReader;
import io.cucumber.datatable.DataTable;
import io.github.bonigarcia.wdm.WebDriverManager;

import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.openqa.selenium.support.Color;

public class CommonFunctions extends ITestListenerImpl {
	RequestSpecification REQ_SPEC;
	Response RESP;
	ValidatableResponse VALIDATABLE_RESP;
	String baseUri;
	String resourcePath;
	String accessToken;
	ExcelUtilResponse dd = new ExcelUtilResponse();
	// protected static WebDriver browser;
	// CreatingUser cu = new CreatingUser();
	static ObjectMapper mapper = new ObjectMapper();
	File jsonFile;
	String[] requiredValue = new String[10];
	// protected static WebDriver browser;
	protected Properties prop = new Properties();
	protected FileInputStream fis;
	protected static ThreadLocal<Reportgenerator> rg = new ThreadLocal<Reportgenerator>();
	// Declare ThreadLocal Driver (ThreadLocalMap) for ThreadSafe Tests
	protected static ThreadLocal<RemoteWebDriver> remoteDriver = new ThreadLocal<RemoteWebDriver>();
	protected static ThreadLocal<WebDriver> webDriver = new ThreadLocal<WebDriver>();
	public static ThreadLocal<String> browserName = new ThreadLocal<String>();
	// window switching variables
	public String childWinAdd;
	public String parWinAdd;
	public String elementTextValue;
	public String takescreenshot = "yes";
	public String totalSubscriptionCountHeader = null;
	public String totalSubscriptionCountPagination = null;
	public String timeBeforeSubscription = null;
	public String dateBeforeSubCreation = null;
	public String timeAtSubscriptionCreation = null;
	public String validValue = null;
	public String explorerPageTransactionCount = null;
	public String subscriptionsPaginationCount = null;
	public String apiValidationReportPath;
	public String simbaUrl;
	public String testCaseQA3 = "TC_65";
	public String testCaseQA4 = "APITC38,APITC41";
	public int size;
	// public String validContractorSubscriber = null;
	int counter = 0;
	// int envLoadCounter = 0;
	String parentWindow = null, childWindow = null;
	protected static int i;

	@BeforeClass

	/**
	 * Initializes the browser and WebDriver instance before tests begin. Supports
	 * Chrome, Firefox, Microsoft Edge, or API test runs. Handles both grid and
	 * local driver configurations, including headless, profiles, and mobile
	 * emulation.
	 *
	 * @param browser Name of the browser to launch (e.g., "chrome", "firefox")
	 * @param x       (Optional) X position for window
	 * @param y       (Optional) Y position for window
	 * @param port    (Optional) Remote debugging port for existing Chrome instances
	 * @throws Exception if browser setup fails
	 */
	@Parameters({ "browser", "positionx", "positiony", "port" })
	public void setupBrowser(String browser, @Optional("0") Integer x, @Optional("0") Integer y,
			@Optional("127.0.0.1:9222") String port) throws Exception {
		try {
			DesiredCapabilities capabilities = new DesiredCapabilities();
			browserName.set(browser);
			capabilities.setCapability("browserName", browser);
			System.out.println("====Browser is " + browser.toUpperCase() + " ====");
			fis = new FileInputStream(System.getProperty("user.dir") + File.separator + "Properties" + File.separator
					+ "DefaultConfig.properties");
			prop.load(fis);
			if (browser.equalsIgnoreCase("chrome")) {
				ChromeOptions options = new ChromeOptions();
				options.addArguments("--disable-dev-shm-usage"); // overcome
				// limited
				// resource
				// problems
				options.addArguments("--disable-notifications");
				options.addArguments("--disable-gpu");
				options.addArguments("--disable-extensions");
				options.addArguments("--no-sandbox");
				options.addArguments("--disable-dev-shm-usage");
				options.addArguments("--remote-allow-origins=*");

				// options.addArguments("user-data-dir=C:\\Users\\" +
				// System.getProperty("user.name")
				// +"\\AppData\\Local\\Google\\Chrome\\User
				// Data\\");
				// options.addArguments("--profile-directory=Person 1");
				options.setExperimentalOption("useAutomationExtension", false);
				HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
				chromePrefs.put("profile.default_content_setting_values.automatic_downloads", 1);
				chromePrefs.put("download.default_directory",
						"C:Users" + File.separator + System.getProperty("user.name") + File.separator + "Downloads"
								+ File.separator + "ChromeDownloads");
				options.setExperimentalOption("prefs", chromePrefs);
				if (prop.getProperty("Headless").equalsIgnoreCase("Y")) {
					options.addArguments("--window-size=1400,600");
					options.addArguments("--headless");
				}
				if (prop.getProperty("Grid").equalsIgnoreCase("Y")) {
					capabilities.merge(options);
					remoteDriver.set(new RemoteWebDriver(new URL(prop.getProperty("GridUrl")), capabilities));
				} else {
					if (prop.getProperty("CICD").equalsIgnoreCase("Y")) {
						// options.addArguments("start-maximized");
						// System.setProperty("webdriver.chrome.driver",prop.getProperty("CICDChromeDriverPath"));
						WebDriverManager.chromedriver().setup();
					} else {
						if (prop.getProperty("ChromeDriverPath").equalsIgnoreCase("WebDriverManager")) {
							System.out.println("Chrome Browser launched through webdriver manager");
							WebDriverManager.chromedriver().setup();
						} else {
							// System.out.printf("Chrome Driver Path is ",System.getProperty("user.dir") +
							// File.separator + prop.getProperty("ChromeDriverPath"));
							System.setProperty("webdriver.chrome.driver",
									System.getProperty("user.dir") + File.separator + File.separator + "Drivers"
											+ File.separator + "chromedriver" + File.separator
											+ prop.getProperty("ChromeDriverPath"));
						}
					}
				
					if (prop.getProperty("ChromeProfile").equalsIgnoreCase("Y")) {
						options.addArguments("chrome.switches", "--disable-extensions");
						options.addArguments("user-data-dir=C:" + File.separator + "Users" + File.separator
								+ System.getProperty("user.name") + File.separator + "AppData" + File.separator
								+ "Local" + File.separator + "Google" + File.separator + "Chrome" + File.separator
								+ "User Data" + i++);
						System.out.println("===============Setting ChromeProfile " + i + "===============");
					}
					if (prop.getProperty("isChromeExisting").equalsIgnoreCase("Y")) {
						options.setExperimentalOption("debuggerAddress", port);
						System.out.println("Started existing Chrome");
					}
					if (prop.getProperty("ChromeIncognito").equalsIgnoreCase("Y")) {
						options.addArguments("--incognito");
						options.addArguments("--window-size=1400,600");
					}
					if (prop.getProperty("MobileEmulation").equalsIgnoreCase("Y")) {
						Map<String, String> mobileEmulation = new HashMap<>();
						mobileEmulation.put("deviceName", prop.getProperty("DeviceName"));
						options.setExperimentalOption("mobileEmulation", mobileEmulation);
					}
					options.addArguments("disable-infobars");
					options.setAcceptInsecureCerts(true);
					options.setUnhandledPromptBehaviour(UnexpectedAlertBehaviour.ACCEPT);
					webDriver.set(new ChromeDriver(options));
					// }
					// getDriver().manage().window().setPosition(new Point(x, y));
					 getDriver().manage().window().maximize();

					getDriver().manage().window().setSize(new Dimension(1850, 950));
					getDriver().manage().timeouts().implicitlyWait(
							Duration.ofSeconds(Long.parseLong(getDefaultProp().getProperty("ImplicitWait"))));
				}
			} else if (browser.equalsIgnoreCase("firefox")) {
				// —— new Firefox branch ——
				FirefoxOptions options = new FirefoxOptions();
				options.setAcceptInsecureCerts(true);
				options.setCapability("moz:webdriverClick", true);

				// headless?
				if ("Y".equalsIgnoreCase(prop.getProperty("Headless"))) {
					options.addArguments("--width=1400");
					options.addArguments("--height=600");
					options.addArguments("--headless");
				}

				if ("Y".equalsIgnoreCase(prop.getProperty("Grid"))) {
					capabilities.merge(options);
					remoteDriver.set(new RemoteWebDriver(new URL(prop.getProperty("GridUrl")), capabilities));
				} else {
					// local driver
					if ("Y".equalsIgnoreCase(prop.getProperty("CICD"))) {
						WebDriverManager.firefoxdriver().setup();
					} else {
						if ("WebDriverManager".equalsIgnoreCase(prop.getProperty("FirefoxDriverPath"))) {
							WebDriverManager.firefoxdriver().setup();
						} else {
							System.setProperty("webdriver.gecko.driver",
									System.getProperty("user.dir") + File.separator + "Drivers" + File.separator
											+ "geckodriver" + File.separator + prop.getProperty("FirefoxDriverPath"));
						}
					}
					webDriver.set(new FirefoxDriver(options));
					getDriver().manage().timeouts().implicitlyWait(
							Duration.ofSeconds(Long.parseLong(getDefaultProp().getProperty("ImplicitWait"))));
				}
				// optional: position or size
				getDriver().manage().window().setPosition(new Point(x, y));
				getDriver().manage().window().maximize();

			} else if (browser.equalsIgnoreCase("MicrosoftEdge")) {
				Runtime.getRuntime().exec("taskkill /f /im msedge.exe");
				EdgeOptions options = new EdgeOptions();
				options.setBinary("C:" + File.separator + "Program Files (x86)" + File.separator + "Microsoft"
						+ File.separator + "Edge" + File.separator + "Application" + File.separator + "msedge.exe");
				if (prop.getProperty("EdgeIncognito").equalsIgnoreCase("Y")) {
					options.addArguments("-inprivate");
				} else {
					options.addArguments(
							"user-data-dir=C:" + File.separator + "Users" + File.separator
									+ System.getProperty("user.name") + File.separator + "AppData" + File.separator
									+ "Local" + File.separator + "Microsoft" + File.separator + "Edge" + File.separator
									+ "User Data" + File.separator,
							"profile-directory=Default", "--disable-extensions");
				}
				if (prop.getProperty("EdgeProfile").equalsIgnoreCase("Y")) {
					options.addArguments(
							"user-data-dir=C:" + File.separator + "Users" + File.separator
									+ System.getProperty("user.name") + File.separator + "AppData" + File.separator
									+ "Local" + File.separator + "Microsoft" + File.separator + "Edge" + File.separator
									+ "User Data" + File.separator + i++,
							"profile-directory=Default", "--disable-extensions");
					System.out.println("===============Setting EdgeProfile " + i + "===============");
				}
				// options.addArguments("--remote-debugging-port=9222");
				options.addArguments("--disable-dev-shm-usage"); // overcome
				options.addArguments("--disable-notifications");
				options.addArguments("--disable-gpu");
				options.addArguments("--disable-extensions");
				options.addArguments("--no-sandbox");
				options.addArguments("--disable-dev-shm-usage");
				options.addArguments("--remote-allow-origins=*");
				options.setUnhandledPromptBehaviour(UnexpectedAlertBehaviour.ACCEPT);
				// limited
				// resource
				// problems
				options.addArguments("--no-sandbox"); // Bypass OS security
				// model
				// HashMap<String, Object> edgePrefs = new HashMap<String,
				// Object>();
				// edgePrefs.put("download.default_directory",
				// "C:\\Users\\" + System.getProperty("user.name") +
				// \\Downloads\\EdgeDownloads);
				// options.setExperimentalOption("prefs", edgePrefs);
				if (prop.getProperty("Headless").equalsIgnoreCase("Y")) {
					options.addArguments("--window-size=1400,600");
					options.addArguments("--headless");
				}
				capabilities.merge(options);
				if (prop.getProperty("Grid").equalsIgnoreCase("Y")) {
					remoteDriver.set(new RemoteWebDriver(new URL(prop.getProperty("GridUrl")), capabilities));
				} else {
					if (prop.getProperty("EdgeDriverPath").equalsIgnoreCase("WebDriverManager")) {
						System.out.println("Edge Browser launched through webdriver manager");
						WebDriverManager.edgedriver().setup();
					} else {
						System.setProperty("webdriver.edge.driver",
								System.getProperty("user.dir") + File.separator + "Drivers" + File.separator
										+ "edgedriver" + File.separator + prop.getProperty("EdgeDriverPath"));
					}
					webDriver.set(new EdgeDriver(options));
					getDriver().manage().timeouts().implicitlyWait(
							Duration.ofSeconds(Long.parseLong(getDefaultProp().getProperty("ImplicitWait"))));

					// getDriver().manage().window().setPosition(new Point(x, y));
					// getDriver().manage().window().setSize(new Dimension(700,1000));
				}
			} else if (browser.equalsIgnoreCase("API")) {
				System.out.println("----------------API TEST-------------");
				try {
					copyApiTemplate();
				} catch (Throwable e) {
					e.printStackTrace();
				}
			} else {
				System.out.println("----------------Invalid Browser-------------");
			}
			// getDriver().manage().window().maximize();
			// getDriver().manage().window().setSize(new Dimension(1500, 1000));
			// getDriver().manage().timeouts().implicitlyWait(Duration.ofSeconds(15));
		} catch (Exception e) {
			System.out.println("In start Browser Exception");
			e.printStackTrace();
		}
	}

	/**
	 * Gets the current browser name used by this thread.
	 * 
	 * @return The browser name.
	 */
	public String getBrowserName() {
		return browserName.get();
	}

	/**
	 * Returns the current WebDriver instance for this thread. Supports switching
	 * between remote and local WebDrivers.
	 * 
	 * @return WebDriver instance or null if not initialized.
	 */
	public WebDriver getDriver() {
		try {
			// Get driver from ThreadLocalMap
			Properties properties = new Properties();
			FileInputStream fis = new FileInputStream(System.getProperty("user.dir") + File.separator + "Properties"
					+ File.separator + "DefaultConfig.properties");
			properties.load(fis);
			if (properties.getProperty("Grid").equalsIgnoreCase("Y")) {
				return remoteDriver.get();
			} else {
				return webDriver.get();
			}
		} // return driver;
		catch (Exception e) {
			System.out.println("In GetDriver catch");
			e.printStackTrace();
			return null;
		}
	}

	/**
	 * Quits the browser after all tests are finished, unless configuration says not
	 * to. Does not close browser if running API tests.
	 * 
	 * @throws Exception if quitting fails
	 */
	@AfterClass
	public void endBrowser() throws Exception {
		if (prop.getProperty("DontCloseBrowser").equalsIgnoreCase("N")) {
			if (!getBrowserName().equalsIgnoreCase("API"))
				getDriver().quit();
		}
	}

	/**
	 * Loads the Object Repository properties from the file system.
	 * 
	 * @return Properties object for OR file.
	 */
	public Properties getORProp() {
		try {
			fis = new FileInputStream(
					System.getProperty("user.dir") + File.separator + "Properties" + File.separator + "OR.properties");
			prop.load(fis);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return prop;
	}

	/**
	 * Loads the Default Configuration properties.
	 * 
	 * @return Properties object for DefaultConfig.properties file.
	 */
	public Properties getDefaultProp() {
		try {
			fis = new FileInputStream(System.getProperty("user.dir") + File.separator + "Properties" + File.separator
					+ "DefaultConfig.properties");
			prop.load(fis);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return prop;
	}

	/**
	 * Decrypts a base64-encoded string.
	 * 
	 * @param encryptedtext The base64-encoded text.
	 * @return The decoded string.
	 */
	public static String decryption(String encryptedtext) {
		byte[] decodedBytes = Base64.decodeBase64(encryptedtext);
		// MyLogger.info("decodedBytes " + new String(decodedBytes));
		return new String(decodedBytes);
	}

	/**
	 * Waits for the web page to fully load (readyState=complete).
	 */
	public void waitForPageToLoad() {
		WebDriverWait wait = new WebDriverWait(getDriver(), Duration.ofSeconds(30));
		wait.until(new ExpectedCondition<Boolean>() {
			public Boolean apply(WebDriver wdriver) {
				return ((JavascriptExecutor) getDriver()).executeScript("return document.readyState")
						.equals("complete");
			}
		});
	}

	/**
	 * Scrolls the page to bring the element matching the xpath into view.
	 * 
	 * @param xpath Xpath of the element to scroll to.
	 * @throws Exception if element not found.
	 */
	public void scrollElement(String xpath) throws Exception {
		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("arguments[0].scrollIntoView({block:'center',inline:'nearest'});",
				getDriver().findElement(By.xpath(xpath)));

	}

	/**
	 * Highlights a UI element by adding a yellow border. Handles stale element
	 * exceptions and logs events.
	 * 
	 * @param xpath The xpath of the element to highlight.
	 */
	public void highlightElement(String xpath) {
		try {
			JavascriptExecutor js = (JavascriptExecutor) getDriver();
			js.executeScript("arguments[0].style.border='2px solid yellow';", getDriver().findElement(By.xpath(xpath)));

		} catch (StaleElementReferenceException e) {
			e.printStackTrace();
			MyLogger.info("Scroll highlight on Stale element" + e.getMessage());
			JavascriptExecutor js = (JavascriptExecutor) getDriver();
			js.executeScript("arguments[0].style.border='2px solid yellow';", getDriver().findElement(By.xpath(xpath)));

			MyLogger.info("StaleElemnt Exception Handled" + e.getMessage());
			throw e;
		} catch (Exception e) {
			e.printStackTrace();
			MyLogger.info("Exception Occurred" + e.getMessage());
			throw e;
		}
	}

	/**
	 * Scrolls to and highlights a specific WebElement.
	 * 
	 * @param element The WebElement to scroll and highlight.
	 * @throws Throwable if JavaScript fails.
	 */
	public void scrollHighLight(WebElement element) throws Throwable {
		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("arguments[0].scrollIntoView({block:'center',inline:'nearest'});", element);
		js.executeScript("arguments[0].style.border='2px solid yellow';", element);
	}

	/**
	 * Takes a screenshot and logs the provided message in ExtentReports. Skips
	 * screenshot if FailedScreenshots is disabled.
	 * 
	 * @param logMessage     Message to log.
	 * @param screenshotName Name for the screenshot.
	 * @param logInfo        ExtentTest logger.
	 * @throws Throwable if screenshot fails.
	 */
	public void addScreenshotAndLog(String logMessage, String screenshotName, ExtentTest logInfo) throws Throwable {
		try {
			if (!logMessage.equals(""))
				MyLogger.info(logMessage);
			if (getDefaultProp().getProperty("FailedScreenshots").equalsIgnoreCase("N")) {
				if (!screenshotName.equals(""))
					logInfo.addScreenCaptureFromPath(captureScreenShot(getDriver()), screenshotName);
			}
		} catch (Exception e) {
			MyLogger.info("Exception Occured " + e.getMessage());
			throw e;
		}
	}

	/**
	 * Takes a screenshot and logs the provided message in ExtentReports. Skips
	 * screenshot if FailedScreenshots is disabled.
	 * 
	 * @param logMessage     Message to log.
	 * @param screenshotName Name for the screenshot.
	 * @param logInfo        ExtentTest logger.
	 * @throws Throwable if screenshot fails.
	 */
	public void addScreenshotAndLog(String screenshotSize, String logMessage, String screenshotName, ExtentTest logInfo)
			throws Throwable {
		try {
			if (!logMessage.equals(""))
				MyLogger.info(logMessage);
			if (getDefaultProp().getProperty("FailedScreenshots").equalsIgnoreCase("N")) {
				if (!screenshotName.equals("")) {
					if (screenshotSize.equalsIgnoreCase("Full"))
						logInfo.addScreenCaptureFromPath(captureScreenShotFullImg(getDriver()), screenshotName);
				}
			}
		} catch (Exception e) {
			MyLogger.info("Exception Occured " + e.getMessage());
			throw e;
		}
	}

	/**
	 * Checks if an element is present in the DOM and visible, based on the key in
	 * the object repository. Scrolls and highlights if found.
	 * 
	 * @param ele     Key for xpath in OR properties.
	 * @param logInfo ExtentTest log.
	 * @return true if present and visible; false otherwise.
	 * @throws Throwable if not found.
	 */
	public boolean isElementPresent(String ele, ExtentTest logInfo) throws Throwable {
		String xpath = prop.getProperty(ele);
		Exception e1 = null;
		boolean found = false;
		final long startTime = System.currentTimeMillis();
		try {
			waitForPageToLoad();
			WebDriverWait wait = new WebDriverWait(getDriver(),
					Duration.ofSeconds(Long.parseLong(getDefaultProp().getProperty("ExplicitWait"))));
			WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(xpath)));
			if (!(element == null)) {
				found = true;
			} else {
				found = false;
			}
		} catch (Exception e) {
			long endTime = System.currentTimeMillis();
			long totalTime = endTime - startTime;
			MyLogger.info("Element not found after waiting for:" + totalTime + " ms : " + xpath);
			e.printStackTrace();
			e1 = e;
		}
		if (found) {
			scrollElement(xpath);
			highlightElement(xpath);
			// logInfo.addScreenCaptureFromBase64String((captureScreenShotBase64(getDriver())),
			// ele);
			if (takescreenshot == "yes") {
				//// addScreenshotAndLog("", ele + " is found", logInfo);
			} else {
				takescreenshot = "yes";
			}
		} else {
			long endTime = System.currentTimeMillis();
			long totalTime = endTime - startTime;
			// // addScreenshotAndLog("Exception : Element not found after waiting for:" +
			// totalTime + " ms : " + xpath,
			// ele + " is not found", logInfo);
			throw e1;
		}
		return found;
	}

	/**
	 * Checks if an element is present in the DOM and visible, based on the key in
	 * the object repository. Scrolls and highlights if found.
	 * 
	 * @param ele     Key for xpath in OR properties.
	 * @param logInfo ExtentTest log.
	 * @return true if present and visible; false otherwise.
	 * @throws Throwable if not found.
	 */

	public boolean isElementPresent(int waitTime, String ele, ExtentTest logInfo) throws Throwable {
		String xpath = prop.getProperty(ele);
		Exception e1 = null;
		boolean found = false;
		final long startTime = System.currentTimeMillis();
		try {
			waitForPageToLoad();
			WebDriverWait wait = new WebDriverWait(getDriver(), Duration.ofSeconds(waitTime));
			WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(xpath)));
			if (!(element == null)) {
				found = true;
			} else {
				found = false;
			}
		} catch (Exception e) {
			long endTime = System.currentTimeMillis();
			long totalTime = endTime - startTime;
			MyLogger.info("Element not found after waiting for:" + totalTime + " ms : " + xpath);
			e.printStackTrace();
			e1 = e;
		}
		if (found) {
			scrollElement(xpath);
			highlightElement(xpath);
			// logInfo.addScreenCaptureFromBase64String((captureScreenShotBase64(getDriver())),
			// ele);
			if (takescreenshot == "yes") {
				// // addScreenshotAndLog("", ele + " is found", logInfo);
			} else {
				takescreenshot = "yes";
			}
		} else {
			long endTime = System.currentTimeMillis();
			long totalTime = endTime - startTime;
			// // addScreenshotAndLog("Exception : Element not found after waiting for:" +
			// totalTime + " ms : " + xpath,
			// ele + " is not found", logInfo);
			throw e1;
		}
		return found;
	}

	/**
	 * Checks whether a dynamic element is present on the UI by its property key and XPATH.
	 * Waits for page load, highlights element, scrolls into view, and logs result.
	 * Throws exception if element not found within explicit wait.
	 *
	 * @param ele    Key for the element in property file.
	 * @param logInfo ExtentTest logger for reporting.
	 * @return true if element is present, false otherwise.
	 * @throws Exception if element not found.
	 */
	public boolean isDynamicElementPresent(String ele, ExtentTest logInfo) throws Exception {
		String xpath = prop.getProperty(ele);
		Exception e1 = null;
		boolean found = false;
		final long startTime = System.currentTimeMillis();
		try {
			waitForPageToLoad();
			WebDriverWait wait = new WebDriverWait(getDriver(),
					Duration.ofSeconds(Long.parseLong(getDefaultProp().getProperty("ExplicitWait"))));
			WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(xpath)));
			if (!(element == null)) {
				found = true;
			} else {
				found = false;
			}
		} catch (Exception e) {
			long endTime = System.currentTimeMillis();
			long totalTime = endTime - startTime;
			MyLogger.info("Element not found after waiting for:" + totalTime + " ms : " + xpath);
			e1 = e;
		}
		if (found) {
			scrollElement(xpath);
			highlightElement(xpath);
			// logInfo.addScreenCaptureFromPath(captureScreenShot(getDriver()),xpath);
		} else {
			long endTime = System.currentTimeMillis();
			long totalTime = endTime - startTime;
			MyLogger.info("Exception : Element not found after waiting for:" + totalTime + " ms : " + xpath);
			throw e1;
		}
		return found;
	}
	
	/**
	 * Checks whether a dynamic element is present by its XPATH (no scrolling).
	 * Waits for page load and highlights element if found.
	 *
	 * @param xpath   XPATH locator string.
	 * @param logInfo ExtentTest logger for reporting.
	 * @return true if element is present, false otherwise.
	 * @throws Exception if element not found.
	 */
	public boolean isDynamicElementPresentWithoutScroll(String xpath, ExtentTest logInfo) throws Exception {
		Exception e1 = null;
		boolean found = false;
		final long startTime = System.currentTimeMillis();
		try {
			waitForPageToLoad();
			WebDriverWait wait = new WebDriverWait(getDriver(),
					Duration.ofSeconds(Long.parseLong(getDefaultProp().getProperty("ExplicitWait"))));
			WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(xpath)));
			if (!(element == null)) {
				found = true;
			} else {
				found = false;
			}
		} catch (Exception e) {
			long endTime = System.currentTimeMillis();
			long totalTime = endTime - startTime;
			MyLogger.info("Element not found after waiting for:" + totalTime + " ms : " + xpath);
			e.printStackTrace();
			e1 = e;
		}
		if (found) {
			// scrollElement(xpath);
			highlightElement(xpath);
			// logInfo.addScreenCaptureFromPath(captureScreenShot(getDriver()),xpath);
		} else {
			long endTime = System.currentTimeMillis();
			long totalTime = endTime - startTime;
			MyLogger.info("Exception : Element not found after waiting for:" + totalTime + " ms : " + xpath);

			throw e1;
		}
		return found;
	}
	
	/**
	 * Verifies if an element is visible on the page, handles StaleElement and TimeoutException.
	 * Logs and retries once on failure.
	 *
	 * @param xpath   XPATH locator string.
	 * @param logInfo ExtentTest logger.
	 * @return true if element is visible, false otherwise.
	 * @throws Throwable for unrecoverable exceptions.
	 */
	public boolean verifyView(String xpath, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		try {
			if (isElementPresent(xpath, logInfo)) {
				MyLogger.info("Verified view for element " + xpath);
				result = true;
			}
			return result;
		} catch (StaleElementReferenceException e) {
			e.printStackTrace();
			MyLogger.info("Verify view on Stale element " + e.getMessage());
			try {
				if (isElementPresent(xpath, logInfo)) {
					result = true;
				}
			} catch (Exception e1) {
				throw e1;
			}
			MyLogger.info("StaleElement Exception Handled " + e.getMessage());
		} catch (TimeoutException e) {
			e.printStackTrace();
			MyLogger.info("Verify view on Timeout element " + e.getMessage());
			try {
				if (isElementPresent(xpath, logInfo)) {
					result = true;
				}
			} catch (Exception e1) {
				MyLogger.info("reached verify view timeout exception2");
				throw e1;
			}
			MyLogger.info("TimeoutException Handled" + e.getMessage());
		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		}
		return result;
	}
	
	/**
	 * Overloaded version of verifyView to allow custom wait time.
	 *
	 * @param seconds Wait time in seconds.
	 * @param xpath   XPATH locator string.
	 * @param logInfo ExtentTest logger.
	 * @return true if element is visible, false otherwise.
	 * @throws Throwable for unrecoverable exceptions.
	 */
	public boolean verifyView(int seconds, String xpath, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		try {
			if (isElementPresent(seconds, xpath, logInfo)) {
				MyLogger.info("Verified view for element " + xpath);
				result = true;
			}
			return result;
		} catch (StaleElementReferenceException e) {
			e.printStackTrace();
			MyLogger.info("Verify view on Stale element " + e.getMessage());
			try {
				if (isElementPresent(seconds, xpath, logInfo)) {
					result = true;
				}
			} catch (Exception e1) {
				throw e1;
			}
			MyLogger.info("StaleElement Exception Handled " + e.getMessage());
		} catch (TimeoutException e) {
			e.printStackTrace();
			MyLogger.info("Verify view on Timeout element " + e.getMessage());
			try {
				if (isElementPresent(seconds, xpath, logInfo)) {
					result = true;
				}
			} catch (Exception e1) {
				MyLogger.info("reached verify view timeout exception2");
				throw e1;
			}
			MyLogger.info("TimeoutException Handled" + e.getMessage());
		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		}
		return result;
	}
	
	/**
	 * Verifies a dynamic element is present and visible, with retry on Stale/Timeout exceptions.
	 *
	 * @param name    Name for logging.
	 * @param xpath   XPATH locator.
	 * @param logInfo ExtentTest logger.
	 * @return true if element found, otherwise throws exception.
	 * @throws Throwable on failure.
	 */
	public boolean verifyDynamicView(String name, String xpath, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		try {
			if (isDynamicElementPresent(xpath, logInfo)) {
				MyLogger.info("Verified view for element " + name);
				result = true;
			}
			return result;
		} catch (StaleElementReferenceException e) {
			e.printStackTrace();
			MyLogger.info("Verify view on Stale element " + e.getMessage());
			if (isDynamicElementPresent(xpath, logInfo)) {
				result = true;
			}
			MyLogger.info("StaleElement Exception Handled " + e.getMessage());
			throw e;
		} catch (TimeoutException e) {
			e.printStackTrace();
			MyLogger.info("Verify view on Timeout element " + e.getMessage());
			if (isDynamicElementPresent(xpath, logInfo)) {
				result = true;
			}
			MyLogger.info("TimeoutException Handled" + e.getMessage());
			throw e;
		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		}
	}
	
	/**
	 * Verifies that an element is not displayed on the page.
	 *
	 * @param xpath   XPATH locator string.
	 * @param logInfo ExtentTest logger.
	 * @return true if element is NOT visible, false otherwise.
	 * @throws Throwable on error.
	 */
	public boolean verifyNoView(String xpath, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		try {
			WebDriverWait wait = new WebDriverWait(getDriver(), Duration.ofSeconds(3));
			WebElement ele = getDriver().findElement(By.xpath(prop.getProperty(xpath)));
			if (ele.isDisplayed()) {
				// // addScreenshotAndLog("", "Failed:" + xpath + " is still found", logInfo);
				return result;
			} else {
				result = true;
				// // addScreenshotAndLog("", "Verified " + xpath + " is not displayed",
				// logInfo);
			}
		} catch (Exception e) {
			result = true;
			// // addScreenshotAndLog("", "Verified " + xpath + " is not displayed",
			// logInfo);
		}
		return result;
	}
	
	/**
	 * Verifies that a dynamic element (by name) is not displayed.
	 *
	 * @param eleName Name for reporting.
	 * @param xpath   XPATH locator string.
	 * @param logInfo ExtentTest logger.
	 * @return true if element is NOT visible, false otherwise.
	 * @throws Throwable on error.
	 */
	public boolean verifDynamicNoView(String eleName, String xpath, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		try {
			WebElement ele = getDriver().findElement(By.xpath(xpath));
			if (ele.isDisplayed()) {
				// // addScreenshotAndLog("", "Failed:" + eleName + " is still found", logInfo);
				return result;
			} else {
				// // addScreenshotAndLog("", "Verified " + eleName + " is not displayed",
				// logInfo);
				result = true;
			}
		} catch (Exception e) {
			// // addScreenshotAndLog("", "Verified " + eleName + " is not displayed",
			// logInfo);
			result = true;
		}
		return result;
	}
	
	/**
	 * Clicks an element, with fallback to JavaScript click and selector type handling.
	 *
	 * @param eleName Name for logging (OR or otherwise).
	 * @param ele     XPATH key or XPATH string.
	 * @param logInfo ExtentTest logger.
	 * @return true if click succeeded.
	 * @throws Throwable on failure.
	 */
	public boolean excepClick(String eleName, String ele, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		try {
			if (isElementPresent(ele, logInfo)) {
				if (ele.toLowerCase().startsWith("sel")) {
					if (eleName == "OR") {
						getDriver().findElement(By.xpath(prop.getProperty(ele))).click();
						// // addScreenshotAndLog("Clicked on " + ele + " using selenium", "Clicked on "
						// + ele, logInfo);
					} else {
						getDriver().findElement(By.xpath(ele)).click();
						// // addScreenshotAndLog("Clicked on " + eleName + " using selenium", "Clicked
						// on " + eleName,
						// logInfo);
					}
					result = true;
				} else {
					if (eleName == "OR") {
						JavascriptExecutor jse = (JavascriptExecutor) getDriver();
						jse.executeScript("arguments[0].click();",
								getDriver().findElement(By.xpath(prop.getProperty(ele))));
						// // addScreenshotAndLog("Clicked on " + ele + " using java script executor",
						// "Clicked on " + ele,
						// logInfo);
					} else {
						JavascriptExecutor jse = (JavascriptExecutor) getDriver();
						jse.executeScript("arguments[0].click();", getDriver().findElement(By.xpath(ele)));
						// // addScreenshotAndLog("Clicked on " + eleName + " using java script
						// executor",
						// "Clicked on " + eleName, logInfo);
					}
					result = true;
				}
			}
		} catch (Exception e) {
			MyLogger.info("Exception occurred : " + e);
			e.printStackTrace();
			throw e;
		}
		return result;
	}
	
	/**
	 * Clicks an element found by its OR property key, retries on Stale/Timeout exceptions.
	 * Supports standard and JS click, logs each attempt.
	 *
	 * @param ele     OR key for XPATH.
	 * @param logInfo ExtentTest logger.
	 * @return true if click succeeded.
	 * @throws Throwable on error.
	 */
	public boolean click(String ele, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		try {
			try {
				if (isElementPresent(ele, logInfo)) {
					if (ele.toLowerCase().startsWith("sel")) {
						getDriver().findElement(By.xpath(prop.getProperty(ele))).click();
						addScreenshotAndLog("Clicked on " + ele, "Clicked on " + ele, logInfo);
						MyLogger.info("Clicked on" + ele);
						result = true;
					} else {
						JavascriptExecutor jse = (JavascriptExecutor) getDriver();
						jse.executeScript("arguments[0].click();",
								getDriver().findElement(By.xpath(prop.getProperty(ele))));
						MyLogger.info("Clicked on" + ele);
						addScreenshotAndLog("Clicked on " + ele, "Clicked on " + ele, logInfo);
						result = true;
					}
				}
			} catch (Exception e) {
				MyLogger.info("Exception occurred : " + e);
				e.printStackTrace();
				throw e;
			}
		} catch (StaleElementReferenceException e) {
			MyLogger.info("Retry click on StaleElement Exception" + e.getMessage());
			try {
				result = excepClick("OR", ele, logInfo);
			} catch (Exception excep) {
				MyLogger.info("Exception occurred : " + excep);
				excep.printStackTrace();
				throw excep;
			}
			MyLogger.info("StaleElement Exception Handled" + e.getMessage());
		} catch (TimeoutException e) {
			MyLogger.info("Retry click on TimeoutException" + e.getMessage());
			try {
				result = excepClick("OR", ele, logInfo);
			} catch (Exception excep) {
				MyLogger.info("Exception occurred : " + excep);
				excep.printStackTrace();
				throw excep;
			}
			MyLogger.info("TimeoutException Handled" + e.getMessage());
		} catch (Exception e) {
			MyLogger.info("Exception occurred : " + e.getMessage());
			e.printStackTrace();
			// addScreenshotAndLog("", "Not Clicked on " + ele, logInfo);
			throw e;
		}
		return result;
	}
	
	/**
	 * Optionally clicks an element if present. Skips action and logs if not found.
	 * Tries standard and JS click, handles any errors gracefully.
	 *
	 * @param ele     OR key for XPATH.
	 * @param logInfo ExtentTest logger.
	 * @return true always, false only if element not present.
	 */
	public boolean clickOptional(String ele, ExtentTest logInfo) throws Throwable {
		try {
			boolean result = false;

			// 1) Check presence first
			if (!isDynamicElementPresent(ele, logInfo)) {
				MyLogger.info("Optional element '" + ele + "' not present; skipping click.");
				logInfo.info("Optional element '" + ele + "' not present; skipping click.");
				return false;
			}

			// 2) Try the normal click, fall back to JS if needed
			By locator = By.xpath(prop.getProperty(ele));
			WebElement el = getDriver().findElement(locator);
			try {
				el.click();
			} catch (Exception clickEx) {
				MyLogger.info("Standard click failed on '" + ele + "': " + clickEx.getMessage()
						+ " → falling back to JS click");
				try {
					((JavascriptExecutor) getDriver()).executeScript("arguments[0].click();", el);
				} catch (Exception jsEx) {
					MyLogger.info("JS click also failed on '" + ele + "': " + jsEx.getMessage());
					// will retry below via excepClick, or ultimately give up
				}
			}

		} catch (Throwable t) {
			// catches *any* unexpected error in the whole method
			MyLogger.info("Optional element '" + ele + "' not present; skipping click.");
			logInfo.info("Optional element '" + ele + "' not present; skipping click.");
			return true;
		}
		return true;
	}

	/**
	 * Dynamically clicks an element by OR property key, supports retries for Stale/Timeout.
	 * Supports both Selenium and JS click based on selector.
	 *
	 * @param eleName Label for reporting.
	 * @param ele     OR key for XPATH.
	 * @param logInfo ExtentTest logger.
	 * @return true if click succeeded.
	 * @throws Throwable on error.
	 */
	public boolean dynamicClick(String eleName, String ele, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		String xpath = prop.getProperty(ele);
		try {
			if (isDynamicElementPresent(ele, logInfo)) {
				if (ele.toLowerCase().startsWith("sel")) {
					getDriver().findElement(By.xpath(xpath)).click();
					addScreenshotAndLog("Clicked on " + eleName + " using selenium", "Clicked on " + eleName, logInfo);
					result = true;
				} else {
					JavascriptExecutor jse = (JavascriptExecutor) getDriver();
					jse.executeScript("arguments[0].click();", getDriver().findElement(By.xpath(xpath)));
					addScreenshotAndLog("Clicked on " + eleName + " using java script executor",
							"Clicked on " + eleName, logInfo);
					result = true;
				}
			}
		} catch (StaleElementReferenceException e) {
			MyLogger.info("Retry click on StaleElement Exception" + e.getMessage());
			try {
				result = excepClick(eleName, ele, logInfo);
			} catch (Exception excep) {
				MyLogger.info("Exception occurred : " + excep);
				excep.printStackTrace();
				throw excep;
			}
			MyLogger.info("StaleElement Exception Handled" + e.getMessage());
		} catch (TimeoutException e) {
			MyLogger.info("Retry click on TimeoutException" + e.getMessage());
			try {
				result = excepClick(eleName, ele, logInfo);
			} catch (Exception excep) {
				MyLogger.info("Exception occurred : " + excep);
				excep.printStackTrace();
				throw excep;
			}
			MyLogger.info("TimeoutException Handled" + e.getMessage());
		} catch (Exception e) {
			MyLogger.info("Exception occurred : " + e.getMessage());
			e.printStackTrace();
			addScreenshotAndLog("", "Not Clicked on " + ele, logInfo);
			throw e;
		}
		return result;
	}
	
	/**
	 * Enters a value into a field by XPATH or property key, supports fallback with JS.
	 * Handles both standard and JS value entry.
	 *
	 * @param ele     OR key or XPATH.
	 * @param text    Value to enter.
	 * @param logInfo ExtentTest logger.
	 * @return true if entry succeeded.
	 * @throws Throwable on error.
	 */
	public boolean excepEntersValue(String ele, String text, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		try {
			if (isElementPresent(ele, logInfo)) {
				if (ele.toLowerCase().startsWith("sel")) {
					JavascriptExecutor jse = (JavascriptExecutor) getDriver();
					jse.executeScript("arguments[0].value='';",
							getDriver().findElement(By.xpath(prop.getProperty(ele))));
					getDriver().findElement(By.xpath(prop.getProperty(ele))).clear();
					getDriver().findElement(By.xpath(prop.getProperty(ele))).sendKeys(text);
					if (ele.contains("siginin_password")) {
						// addScreenshotAndLog("Entered encrypted text in " + ele + " using send keys",
						// "Entered Encrypted Password in " + ele, logInfo);
					} else {
						// addScreenshotAndLog("Entered value " + text + " in " + ele + " using send
						// keys",
						// "Entered " + text + " in " + ele, logInfo);
					}
					result = true;
				} else {
					JavascriptExecutor jse = (JavascriptExecutor) getDriver();
					jse.executeScript("arguments[0].value='';",
							getDriver().findElement(By.xpath(prop.getProperty(ele))));
					jse.executeScript("arguments[0].value='" + text + "';",
							getDriver().findElement(By.xpath(prop.getProperty(ele))));
					if (ele.contains("siginin_password")) {
						// addScreenshotAndLog("Entered encrypted text in " + ele + " using java script
						// executor",
						// "Entered Encrypted Password in " + ele, logInfo);
					} else {
						// addScreenshotAndLog("Entered value " + text + " in " + ele + " using java
						// script executor",
						// "Entered " + text + " in " + ele, logInfo);
					}
					result = true;
				}
			}
		} catch (Exception e) {
			MyLogger.info("Exception occurred : " + e);
			e.printStackTrace();
			throw e;
		}
		return result;
	}

	/**
	 * Selects an option from a dropdown <select> element by text, value, or index.
	 * 
	 * @param option  The value/text/index to select.
	 * @param ele     Key for the element in OR properties.
	 * @param mode    "text", "value", or "index"
	 * @param logInfo ExtentTest logger.
	 * @return true if option was selected, false otherwise.
	 * @throws Throwable if selection fails.
	 */
	public boolean selectDropdown(String option, String ele, String mode, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		try {
			// 1) make sure the dropdown exists
			if (!isElementPresent(ele, logInfo)) {
				MyLogger.info("Dropdown '" + ele + "' not present; skipping select.");
				logInfo.info("Dropdown '" + ele + "' not present; skipping select.");
				return false;
			}

			WebElement dropdownEl = getDriver().findElement(By.xpath(prop.getProperty(ele)));
			addScreenshotAndLog("About to select [" + mode + "=" + option + "] on " + ele, "Before select", logInfo);

			Select dropdown = new Select(dropdownEl);
			String m = mode.trim().toLowerCase(Locale.UK);

			switch (m) {
			case "value":
				dropdown.selectByValue(option);
				break;
			case "index":
				try {
					int idx = Integer.parseInt(option);
					dropdown.selectByIndex(idx);
				} catch (NumberFormatException nfe) {
					throw new IllegalArgumentException("Invalid index '" + option + "' for selectDropdown", nfe);
				}
				break;
			case "text":
			default:
				// default to visible text
				dropdown.selectByVisibleText(option);
				break;
			}

			addScreenshotAndLog("Selected [" + mode + "=" + option + "] on " + ele, "After select", logInfo);
			MyLogger.info("Dropdown '" + ele + "' selected by " + mode + ": " + option);
			result = true;

		} catch (Exception e) {
			MyLogger.info("Exception in selectDropdown(" + ele + "): " + e.getMessage());
			addScreenshotAndLog("Failed selecting [" + mode + "=" + option + "] on " + ele, "selectDropdown error",
					logInfo);
			throw e;
		}
		return result;
	}

	/**
	 * Enters a value into a text field, clearing the field first. Handles
	 * StaleElement and TimeoutException by retrying.
	 * 
	 * @param text    Text to enter.
	 * @param ele     Key for the element in OR properties.
	 * @param logInfo ExtentTest logger.
	 * @return true if value was entered successfully.
	 * @throws Throwable if entry fails.
	 */
	public boolean entersValue(String text, String ele, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		try {
			if (isElementPresent(ele, logInfo)) {
				if (ele.toLowerCase().startsWith("sel")) {
					JavascriptExecutor jse = (JavascriptExecutor) getDriver();
					jse.executeScript("arguments[0].value='';",
							getDriver().findElement(By.xpath(prop.getProperty(ele))));
					getDriver().findElement(By.xpath(prop.getProperty(ele))).clear();
					getDriver().findElement(By.xpath(prop.getProperty(ele))).sendKeys(text);
				} else {
					JavascriptExecutor jse = (JavascriptExecutor) getDriver();
					jse.executeScript("arguments[0].value='';",
							getDriver().findElement(By.xpath(prop.getProperty(ele))));
					jse.executeScript("arguments[0].value='" + text + "';",
							getDriver().findElement(By.xpath(prop.getProperty(ele))));
				}
				addScreenshotAndLog("Entered " + text + " on " + ele, "Entered" + ele, logInfo);
				result = true;
			}
		} catch (StaleElementReferenceException StaleExcep) {
			MyLogger.info("Retry entering value on StaleElement Exception : " + StaleExcep.getMessage());
			try {
				result = excepEntersValue(ele, text, logInfo);
			} catch (Exception excep) {
				MyLogger.info("Exception occurred : " + excep);
				excep.printStackTrace();
				throw excep;
			}
			MyLogger.info("StaleElement Exception Handled" + StaleExcep.getMessage());
		} catch (TimeoutException e) {
			MyLogger.info("Retry entering value on TimeoutException" + e.getMessage());
			try {
				result = excepEntersValue(ele, text, logInfo);
			} catch (Exception excep) {
				MyLogger.info("Exception occurred : " + excep);
				excep.printStackTrace();
				throw excep;
			}
			MyLogger.info("TimeoutException Handled" + e.getMessage());
		} catch (Exception e) {
			MyLogger.info("Exception occurred : " + e.getMessage());
			e.printStackTrace();
			addScreenshotAndLog("", "Cound not enter value on " + ele, logInfo);
			throw e;
		}
		return result;
	}

	/**
	 * Enters a value into a text field, clearing the field first. Handles
	 * StaleElement and TimeoutException by retrying.
	 * 
	 * @param text    Text to enter.
	 * @param ele     Key for the element in OR properties.
	 * @param logInfo ExtentTest logger.
	 * @return true if value was entered successfully.
	 * @throws Throwable if entry fails.
	 */
	public boolean entersAriaValue(String text, String ele, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		try {
			if (isElementPresent(ele, logInfo)) {
				if (ele.toLowerCase().startsWith("sel")) {
					WebElement element = getDriver().findElement(By.xpath(prop.getProperty(ele)));
					element.clear();
					element.sendKeys(text);
					// Thread.sleep(2000);
					// element.sendKeys(Keys.ARROW_DOWN);
					// element.sendKeys(Keys.ENTER);
				} else {
					JavascriptExecutor jse = (JavascriptExecutor) getDriver();
					jse.executeScript("arguments[0].value='';",
							getDriver().findElement(By.xpath(prop.getProperty(ele))));
					jse.executeScript("arguments[0].value='" + text + "';",
							getDriver().findElement(By.xpath(prop.getProperty(ele))));
				}
				addScreenshotAndLog("Entered " + text + " on " + ele, "Entered" + ele, logInfo);
				result = true;
			}
		} catch (StaleElementReferenceException StaleExcep) {
			MyLogger.info("Retry entering value on StaleElement Exception : " + StaleExcep.getMessage());
			try {
				result = excepEntersValue(ele, text, logInfo);
			} catch (Exception excep) {
				MyLogger.info("Exception occurred : " + excep);
				excep.printStackTrace();
				throw excep;
			}
			MyLogger.info("StaleElement Exception Handled" + StaleExcep.getMessage());
			addScreenshotAndLog("", "Cound not enter value on " + ele, logInfo);
		} catch (TimeoutException e) {
			MyLogger.info("Retry entering value on TimeoutException" + e.getMessage());
			try {
				result = excepEntersValue(ele, text, logInfo);
			} catch (Exception excep) {
				MyLogger.info("Exception occurred : " + excep);
				excep.printStackTrace();
				throw excep;
			}
			MyLogger.info("TimeoutException Handled" + e.getMessage());
			addScreenshotAndLog("", "Cound not enter value on " + ele, logInfo);
		} catch (Exception e) {
			MyLogger.info("Exception occurred : " + e.getMessage());
			e.printStackTrace();
			addScreenshotAndLog("", "Cound not enter value on " + ele, logInfo);
			throw e;
		}
		return result;
	}

	/**
	 * Verifies all search result cards contain a keyword in title or summary, across pages.
	 * Logs results, failures, and throws exception if any mismatch is found.
	 *
	 * @param keyword Keyword to verify.
	 * @param logInfo ExtentTest logger.
	 * @return true if all cards matched.
	 * @throws Exception if mismatch found.
	 */
	public boolean verifySearchResults(String keyword, ExtentTest logInfo) throws Exception {
		MyLogger.info("=== Entering verifySearchResults(keyword=" + keyword + ") ===");
		List<String> failures = new ArrayList<>();
		int pageOffset = 0;
		boolean hasNext = true;

		try {
			while (hasNext) {
				List<WebElement> cards = getCardsOnPage(logInfo);
				ensureCardsExist(cards, logInfo);

				for (int i = 1; i <= cards.size(); i++) {
					int globalIndex = pageOffset + i;
					verifyCard(keyword, globalIndex, failures, logInfo, i);
				}

				pageOffset += cards.size();
				hasNext = navigateToNextPage(logInfo);
			}

			reportFailuresIfAny(failures, logInfo);
			MyLogger.info("verifySearchResults completed successfully");
			return true;

		} catch (Exception e) {
			MyLogger.info("verifySearchResults encountered exception: " + e.getMessage());
			logInfo.fail("Error in verifySearchResults: " + e.getMessage());
			throw new Exception("verifySearchResults failed: " + e.getMessage(), e);
		}
	}

	
	/**
	 * Returns all job card elements currently visible on the search results page.
	 * Handles exceptions by logging and returning an empty list if none found.
	 *
	 * @param logInfo ExtentTest logger for reporting errors.
	 * @return List of WebElements representing job cards; empty list if not found or error occurs.
	 */
	private List<WebElement> getCardsOnPage(ExtentTest logInfo) {
		MyLogger.info("Fetching cards on current page");
		try {
			List<WebElement> cards = getDriver().findElements(By.xpath("//li[@data-test='search-result']"));
			MyLogger.info("Found " + cards.size() + " cards");
			return cards;
		} catch (Exception e) {
			MyLogger.info("Error fetching cards on page: " + e.getMessage());
			logInfo.fail("Error fetching cards on page: " + e.getMessage());
			return Collections.emptyList();
		}
	}

	
	/**
	 * Ensures that at least one job card element is present on the page.
	 * Throws an exception if no cards are found to prevent proceeding with empty results.
	 *
	 * @param cards   List of WebElements representing job cards.
	 * @param logInfo ExtentTest logger for reporting.
	 * @throws Exception if no cards are found on the page.
	 */
	private void ensureCardsExist(List<WebElement> cards, ExtentTest logInfo) throws Exception {
		MyLogger.info("Ensuring at least one card exists");
		if (cards.isEmpty()) {
			MyLogger.info("No search results found on page");
			logInfo.info("No search results found on page.");
			throw new Exception("No search results found on page.");
		}
	}

	
	/**
	 * Verifies a single job card for the presence of the specified keyword in the title or summary.
	 * If the title does not contain the keyword, opens the card and checks the job summary text.
	 * Logs failures and appends failed cases to the failures list.
	 *
	 * @param keyword     Keyword to check in title or summary.
	 * @param globalIndex Global card index (across all pages).
	 * @param failures    List to accumulate failure messages.
	 * @param logInfo     ExtentTest logger.
	 * @param localIndex  Index of card on the current page (1-based).
	 */
	private void verifyCard(String keyword, int globalIndex, List<String> failures, ExtentTest logInfo,
			int localIndex) {
		MyLogger.info("Verifying card #" + globalIndex + " for keyword: " + keyword);
		String titleXpath = "(" + "//li[@data-test='search-result']//a[@data-test='search-result-job-title']" + ")["
				+ localIndex + "]";

		try {
			WebElement titleEl = getDriver().findElement(By.xpath(titleXpath));
			String title = titleEl.getText().trim();
			MyLogger.info("Card #" + globalIndex + " title = \"" + title + "\"");
			String keyLower = keyword.toLowerCase(Locale.UK);
			String titleLower = title.toLowerCase(Locale.UK);

			if (titleLower.contains(keyLower)) {
				MyLogger.info("Card #" + globalIndex + ": title contains keyword");
				logInfo.pass("Card #" + globalIndex + ": title \"" + title + "\" contains \"" + keyword + "\"");
			} else {
				MyLogger.info("Card #" + globalIndex + ": title does NOT contain keyword, checking summary");
				logInfo.info("Card #" + globalIndex + ": title \"" + title + "\" does NOT contain \"" + keyword
						+ "\" → checking summary");
				titleEl.click();

				waitForBody(logInfo);

				String bodyText = getDriver().findElement(By.tagName("body")).getText().toLowerCase(Locale.UK);

				if (bodyText.contains(keyLower)) {
					MyLogger.info("Card #" + globalIndex + ": summary contains keyword");
					logInfo.pass("Card #" + globalIndex + ": summary contains \"" + keyword + "\"");
				} else {
					String msg = String.format(
							"Card #%d FAILED: neither title nor summary contains \"%s\" (title was: \"%s\")",
							globalIndex, keyword, title);
					MyLogger.info("Card #" + globalIndex + " failed keyword check");
					logInfo.fail(msg);
					failures.add(msg);
				}

				navigateBackAndWait(logInfo);
			}

		} catch (Exception e) {
			String err = String.format("Card #%d ERROR: %s", globalIndex, e.getMessage());
			MyLogger.info(err);
			logInfo.fail(err);
			failures.add(err);
		}
	}

	
	/**
	 * Navigates to the next page of job search results if a 'Next' link is present.
	 * Waits for the content of the page to update before returning.
	 * Logs and returns false if navigation fails or no next page exists.
	 *
	 * @param logInfo ExtentTest logger for reporting.
	 * @return true if navigation to next page succeeded; false if at last page or on error.
	 */
	private boolean navigateToNextPage(ExtentTest logInfo) {
		MyLogger.info("Checking for next page link");
		try {
			List<WebElement> next = getDriver().findElements(By.xpath("//a[@data-test='search-next-page']"));
			if (next.isEmpty()) {
				MyLogger.info("No Next link – end of pagination");
				return false;
			}

			String firstBefore = getDriver()
					.findElement(By.xpath(
							"(" + "//li[@data-test='search-result']//a[@data-test='search-result-job-title']" + ")[1]"))
					.getText().trim();

			MyLogger.info("Clicking Next page…");
			logInfo.info("Clicking Next page…");
			next.get(0).click();

			new WebDriverWait(getDriver(), Duration.ofSeconds(2))
					.until(ExpectedConditions.not(ExpectedConditions.textToBe(By.xpath(
							"(" + "//li[@data-test='search-result']//a[@data-test='search-result-job-title']" + ")[1]"),
							firstBefore)));

			MyLogger.info("Navigated to next page successfully");
			return true;

		} catch (Exception e) {
			MyLogger.info("Error navigating to next page: " + e.getMessage());
			logInfo.fail("Error navigating to next page: " + e.getMessage());
			return false;
		}
	}

	
	/**
	 * Waits until the <body> tag is present in the DOM, indicating that the new page or details are loaded.
	 * Logs any exceptions encountered during the wait.
	 *
	 * @param logInfo ExtentTest logger for reporting.
	 */
	private void waitForBody(ExtentTest logInfo) {
		MyLogger.info("Waiting for <body> to be present");
		try {
			new WebDriverWait(getDriver(), Duration.ofSeconds(2))
					.until(ExpectedConditions.presenceOfElementLocated(By.tagName("body")));
			MyLogger.info("<body> is present");
		} catch (Exception e) {
			MyLogger.info("Error waiting for body: " + e.getMessage());
			logInfo.fail("Error waiting for body: " + e.getMessage());
		}
	}

	
	/**
	 * Navigates the browser back to the previous page (search results) and waits until job cards are reloaded.
	 * Logs errors if navigation or waiting fails.
	 *
	 * @param logInfo ExtentTest logger for reporting.
	 */
	private void navigateBackAndWait(ExtentTest logInfo) {
		MyLogger.info("Navigating back to results");
		try {
			getDriver().navigate().back();
			new WebDriverWait(getDriver(), Duration.ofSeconds(2)).until(
					ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//li[@data-test='search-result']")));
			MyLogger.info("Returned to results and cards reloaded");
		} catch (Exception e) {
			MyLogger.info("Error navigating back: " + e.getMessage());
			logInfo.fail("Error navigating back: " + e.getMessage());
		}
	}

	
	/**
	 * Reports all accumulated failures after result card verification.
	 * Logs each failure and throws an exception if any are present.
	 *
	 * @param failures List of failure messages from card verification.
	 * @param logInfo  ExtentTest logger for reporting.
	 * @throws Exception summarizing all verification failures.
	 */
	private void reportFailuresIfAny(List<String> failures, ExtentTest logInfo) throws Exception {
		MyLogger.info("Reporting any failures");
		if (!failures.isEmpty()) {
			MyLogger.info("Found " + failures.size() + " failures");
			logInfo.fail("=== VERIFICATION SUMMARY: MISMATCHED CARDS ===");
			failures.forEach(f -> {
				MyLogger.info("Failure detail: " + f);
				logInfo.fail(f);
			});
			throw new Exception("Search verification failed on " + failures.size() + " card(s).");
		}
	}

	
	/**
	 * Checks whether a 'no results' element is present.
	 *
	 * @param ele     OR key for locator.
	 * @param logInfo ExtentTest logger.
	 * @return true if present, false otherwise.
	 * @throws Throwable on error.
	 */
	public boolean noResultsFound(String ele, ExtentTest logInfo) throws Throwable {
		MyLogger.info("Checking for 'no results' message with locator: " + ele);
		try {
			boolean result = false;
			if (isElementPresent(ele, logInfo)) {
				MyLogger.info("Element: " + ele + " present ");
				addScreenshotAndLog("Element: " + ele + " present ", ele, logInfo);
				result = true;
			} else {
				MyLogger.info("Element: " + ele + " not present ");
				addScreenshotAndLog("Element: " + ele + " not present ", ele, logInfo);
				result = false;
			}
			return result;
		} catch (Exception e) {
			MyLogger.info("Unable to verify presence: " + e.getMessage());
			logInfo.fail("Unable to verify presence" + e.getMessage());
			throw e;
		}
		}
	
		
		/**
		 * Verifies all job cards are posted by the specified employer, across all pages.
		 * Logs mismatches and throws if any fail.
		 *
		 * @param keyword Employer name expected.
		 * @param logInfo ExtentTest logger.
		 * @return true if all cards match employer.
		 * @throws Exception if verification fails.
		 */
	public boolean verifyEmployerSearchResults(String keyword, ExtentTest logInfo) throws Exception {
		List<String> failures = new ArrayList<>();
		int pageOffset = 0;
		boolean hasNext = true;

		try {
			while (hasNext) {
				List<WebElement> cards = getCardsOnPage(logInfo);
				ensureCardsExist(cards, logInfo);

				for (int i = 1; i <= cards.size(); i++) {
					int globalIndex = pageOffset + i;
					verifyEmployerCard(keyword, globalIndex, failures, logInfo, i);
				}

				pageOffset += cards.size();
				hasNext = navigateToNextPage(logInfo);
			}

			reportFailuresIfAny(failures, logInfo);
			return true;
		} catch (Exception e) {
			logInfo.fail("Error in verifySearchResults: " + e.getMessage());
			// re-throw so the calling step fails
			throw e;
		}
	}

	
	/**
	 * Verifies that a single job card was posted by the expected employer.
	 * <p>
	 * Extracts the employer name from the card at the specified index and checks it matches
	 * the expected employer name (case-insensitive). On failure, logs the mismatch and
	 * adds details to the failures list. Handles and logs any exceptions that occur during lookup.
	 *
	 * @param expectedEmployer The expected employer name to match (case-insensitive).
	 * @param globalIndex      The global index of the card (across all paginated results).
	 * @param failures         The list to which failure messages should be added.
	 * @param logInfo          ExtentTest logger for reporting pass/fail steps.
	 * @param localIndex       The index of the card on the current page (1-based).
	 */
	private void verifyEmployerCard(String expectedEmployer, int globalIndex, List<String> failures, ExtentTest logInfo,
			int localIndex) {
		// XPath for the employer name in the Nth card
		String employerXpath = "(" + "//li[@data-test='search-result']//h3[@class='nhsuk-u-font-weight-bold']" + ")["
				+ localIndex + "]";

		MyLogger.info("Verifying employer for Card #" + globalIndex + " using XPath: " + employerXpath);

		try {
			WebElement employerEl = getDriver()
					.findElement(By.xpath("//li[@data-test='search-result']//h3[@class='nhsuk-u-font-weight-bold']"));

			String fullText = employerEl.getText().trim();
			// split on any line‐break, take only the first segment (the employer name)
			String employer = fullText.split("\\r?\\n")[0].replace("\"", "") // strip stray quotes, if you still have
																				// them
					.trim();

			MyLogger.info("Employer = " + employer);
			MyLogger.info("Card #" + globalIndex + ": found employer text = \"" + employer + "\"");

			if (employer.equalsIgnoreCase(expectedEmployer.trim())) {
				String passMsg = "Card #" + globalIndex + ": employer \"" + employer + "\" matches expected \""
						+ expectedEmployer + "\"";
				MyLogger.info(passMsg);
				logInfo.pass(passMsg);
			} else {
				String msg = String.format("Card #%d FAILED: employer is \"%s\" (expected: \"%s\")", globalIndex,
						employer, expectedEmployer);
				MyLogger.info(msg);
				logInfo.fail(msg);
				failures.add(msg);
			}
		} catch (Exception e) {
			String err = String.format("Card #%d ERROR: %s", globalIndex, e.getMessage());
			MyLogger.info(err);
			logInfo.fail(err);
			failures.add(err);
		}
	}

	
	/**
	 * Paginates through all job search result pages, clicking 'Next' until last page.
	 * Logs each page visited.
	 *
	 * @param log ExtentTest logger.
	 * @return true if pagination successful.
	 * @throws Throwable on error.
	 */
	public boolean traverseAllPages(ExtentTest log) throws Throwable {
		String cardsXpath = "//li[@data-test='search-result']";
		String nextLinkXpath = "//a[@data-test='search-next-page']";
		int page = 1;

		try {
			while (true) {
				List<WebElement> nextLinks = getDriver().findElements(By.xpath(nextLinkXpath));
				if (nextLinks.isEmpty()) {
					log.pass("Pagination complete after page " + page);
					MyLogger.info("No 'Next' link on page " + page + ". Reached last page.");
					break;
				}

				WebElement nextBtn = nextLinks.get(0);
				WebElement firstCard = getDriver().findElement(By.xpath(cardsXpath));

				MyLogger.info("Clicking 'Next' on page " + page);
				log.info("Clicking 'Next' on page " + page);

				// --- NEW: scroll into view via JavaScript before clicking ---
				((JavascriptExecutor) getDriver()).executeScript("arguments[0].scrollIntoView({block: 'center'});",
						nextBtn);
				// small pause to let browser reflow
				Thread.sleep(250);

				nextBtn.click();

				// wait for the old page to unload
				new WebDriverWait(getDriver(), Duration.ofSeconds(10)).until(ExpectedConditions.stalenessOf(firstCard));

				page++;
			}
			return true;

		} catch (Exception e) {
			log.fail("Error during pagination: " + e.getMessage());
			MyLogger.error("Error during pagination: " + e.getMessage());
			throw e;
			
		}
	}

	
	/**
	 * Verifies search result cards are sorted by date posted in descending order (newest first) across pages.
	 *
	 * @param extentStep ExtentTest logger.
	 * @return true if sorted descending, false otherwise.
	 * @throws Exception if date parsing fails or unsorted.
	 */
	public boolean verifySortedByDateDesc(ExtentTest extentStep) throws Exception {
		DateTimeFormatter fmt = DateTimeFormatter.ofPattern("d MMMM yyyy", Locale.UK);

		String cardsXpath = "//li[@data-test='search-result']";
		String dateXpath = "//li[@data-test='search-result-publicationDate']//strong";
		String titleRelXpath = ".//a[@data-test='search-result-job-title']";
		String firstDateXpath = "(" + dateXpath + ")[1]";
		String nextLinkXpath = "//a[@data-test='search-next-page']";

		List<String> failures = new ArrayList<>();
		boolean hasNext = true;

		while (hasNext) {
			// 1) find all cards and their dates on this page
			List<WebElement> cards = getDriver().findElements(By.xpath(cardsXpath));
			List<WebElement> dateEls = getDriver().findElements(By.xpath(dateXpath));
			int perPageTotal = Math.min(cards.size(), dateEls.size());
			if (perPageTotal == 0) {
				extentStep.fail("No search-result cards or dates found on this page");
				throw new Exception("No search results found on page.");
			}

			// parse dates
			List<LocalDate> dates = new ArrayList<>(perPageTotal);
			for (int i = 0; i < perPageTotal; i++) {
				String txt = dateEls.get(i).getText().trim();
				try {
					dates.add(LocalDate.parse(txt, fmt));
				} catch (DateTimeParseException e) {
					String msg = "Failed to parse date \"" + txt + "\": " + e.getMessage();
					extentStep.fail(msg);
					throw new Exception(msg, e);
				}
			}

			// 2) compare each pair of neighboring cards
			for (int i = 1; i < perPageTotal; i++) {
				WebElement prevCard = cards.get(i - 1);
				WebElement currCard = cards.get(i);
				String prevTitle = prevCard.findElement(By.xpath(titleRelXpath)).getText().trim();
				String currTitle = currCard.findElement(By.xpath(titleRelXpath)).getText().trim();
				LocalDate prevDate = dates.get(i - 1);
				LocalDate currDate = dates.get(i);

				// log the titles and dates
				extentStep
						.info(String.format("Card #%d: \"%s\" – Date Posted: %s", i, prevTitle, prevDate.format(fmt)));
				extentStep.info(
						String.format("Card #%d: \"%s\" – Date Posted: %s", i + 1, currTitle, currDate.format(fmt)));

				// verify descending order
				if (currDate.isAfter(prevDate)) {
					String msg = String.format("Order violation: Card #%d date %s is after Card #%d date %s", i + 1,
							currDate.format(fmt), i, prevDate.format(fmt));
					extentStep.fail(msg);
					failures.add(msg);
				} else {
					extentStep.pass(String.format("Dates OK: Card #%d (%s) ≥ Card #%d (%s)", i, prevDate.format(fmt),
							i + 1, currDate.format(fmt)));
				}
			}

			// 3) pagination: click Next if present
			List<WebElement> nextLinks = getDriver().findElements(By.xpath(nextLinkXpath));
			if (nextLinks.isEmpty()) {
				hasNext = false;
			} else {
				String oldFirstDate = dateEls.get(0).getText().trim();
				extentStep.info("Clicking Next page…");
				nextLinks.get(0).click();
				// wait for the first date to change
				new WebDriverWait(getDriver(), Duration.ofSeconds(10)).until(
						ExpectedConditions.not(ExpectedConditions.textToBe(By.xpath(firstDateXpath), oldFirstDate)));
			}
		}

		// 4) final summary
		if (!failures.isEmpty()) {
			extentStep.fail("=== VERIFICATION SUMMARY: UNSORTED DATES ===");
			failures.forEach(extentStep::fail);
			return false;
		} else {
			extentStep.pass("All pages are sorted by Date Posted (newest first)");
			return true;
		}
	}

	
	/**
	 * Performs an action (click, hover, scroll, sendkeys) on an element.
	 *
	 * @param operation Action string ("click", "hover", "scroll", or "sendkeys:<text>").
	 * @param ele       OR key for XPATH.
	 * @param logInfo   ExtentTest logger.
	 * @return true if action performed.
	 * @throws Throwable on error.
	 */
	public boolean performAction(String operation, String ele, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		try {
			WebElement element = getDriver().findElement(By.xpath(prop.getProperty(ele)));
			if (element.isDisplayed()) {
				Actions act = new Actions(getDriver());
				if (operation.equalsIgnoreCase("click") || operation.equalsIgnoreCase("clicks")) {
					act.moveToElement(element).click().build().perform();
					// addScreenshotAndLog("", "Clicked on " + ele, logInfo);
					result = true;
				} else if (operation.equalsIgnoreCase("hover")) {
					act.moveToElement(element).build().perform();
					// addScreenshotAndLog("", "Hovered on " + ele, logInfo);
					result = true;
				} else if (operation.equalsIgnoreCase("scroll")) {
					act.moveToElement(element).build().perform();
					// addScreenshotAndLog("", "Scrolled on " + ele, logInfo);
					result = true;
				} else if (operation.startsWith("sendkeys")) {
					String text = operation.substring(9);
					act.sendKeys(element, text).build().perform();
					// addScreenshotAndLog("", "Entered " + text + "in" + ele, logInfo);
					result = true;
				}
			}
		} catch (Exception e) {
			MyLogger.info("Exception occurred : " + e.getMessage());
			e.printStackTrace();
			// addScreenshotAndLog("", "Action is not performed on " + ele, logInfo);
			throw e;
		}
		return result;
	}

	
	/**
	 * Clicks a button and switches to a newly opened child window.
	 * Waits for expected number of windows, and switches context.
	 *
	 * @param expWind  Expected window count after click.
	 * @param btn      OR key for button XPATH.
	 * @param logInfo  ExtentTest logger.
	 * @return true if switched to child window.
	 * @throws Throwable on error.
	 */
	public boolean switchWindow(int expWind, String btn, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		try {
			parWinAdd = getDriver().getWindowHandle();
			click(btn, logInfo);
			new WebDriverWait(getDriver(), Duration.ofSeconds(30))
					.until(ExpectedConditions.numberOfWindowsToBe(expWind));
			Set<String> windAdd = getDriver().getWindowHandles();
			for (String each : windAdd) {
				if (!parWinAdd.equals(each)) {
					childWinAdd = each;
					getDriver().switchTo().window(childWinAdd);
					// addScreenshotAndLog("Clicked on " + btn + " and Switched to window",
					// "Clicked on " + btn + " and Switched to window", logInfo);
					result = true;
				}
			}
		} catch (Exception e) {
			MyLogger.info("Exception occurred : " + e.getMessage());
			e.printStackTrace();
			// addScreenshotAndLog(btn + "is not Clicked and Switched to window",
			// btn + "is not Clicked and Switched to window", logInfo);
			throw e;
		}
		return result;
	}

	
	/**
	 * Closes the current child window and switches focus back to parent window.
	 *
	 * @param logInfo ExtentTest logger.
	 * @return true if child window closed.
	 * @throws Throwable on error.
	 */
	public boolean closeChildWindow(ExtentTest logInfo) throws Throwable {
		boolean result = true;
		try {
			Set<String> Windows = getDriver().getWindowHandles();
			for (String currWindow : Windows) {
				if (currWindow.equals(childWinAdd)) {
					getDriver().switchTo().window(childWinAdd);
					getDriver().close();
					MyLogger.info("childwindow closed " + childWinAdd);
					getDriver().switchTo().window(parWinAdd);
					// addScreenshotAndLog("Closed child window and switched to Parent Window",
					// "Closed child window and switched to Parent Window", logInfo);
					result = true;
				}
			}
		} catch (Exception e) {
			MyLogger.info("Exception occurred : " + e.getMessage());
			e.printStackTrace();
			// addScreenshotAndLog("Not Switched to Parent Window", "Not Switched to Parent
			// Window", logInfo);
			throw e;
		}
		return result;
	}

	
	/**
	 * Waits until a specific element (signin_emailaddress) becomes visible.
	 *
	 * @param logInfo ExtentTest logger.
	 * @return false always (design oversight?).
	 * @throws Throwable on error.
	 */
	public boolean waitUntilForElements(ExtentTest logInfo) throws Throwable {
		try {
			WebDriverWait wait = new WebDriverWait(getDriver(), Duration.ofSeconds(5));
			WebElement ele = wait.until(
					ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("signin_emailaddress"))));
		} catch (Exception e) {
			MyLogger.info("Exception Occured : " + e.getMessage());
			addScreenshotAndLog("", "Not Logged into SimbaChain Application", logInfo);
			throw e;
		}
		return false;
	}

	
	/**
	 * Launches a given URL in the browser, logs result.
	 *
	 * @param url     URL to navigate.
	 * @param logInfo ExtentTest logger.
	 * @return true if navigation succeeded.
	 * @throws Throwable on error.
	 */
	public boolean launchURL(String url, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		Boolean resultArr[] = new Boolean[20];
		int count = 0;
		try {
			getDriver().get(url);
			addScreenshotAndLog("", "URL launched: " + url, logInfo);
			result = true;

		} catch (Exception e) {
			MyLogger.info("Exception occurred : " + e.getMessage());
			e.printStackTrace();
			addScreenshotAndLog("", "URL not launched: " + url, logInfo);
			throw e;
		}
		return result;
	}

	
	/**
	 * Returns a list of WebElements by property key XPATH, after ensuring visibility.
	 *
	 * @param xpath   OR key for XPATH.
	 * @param logInfo ExtentTest logger.
	 * @return List of WebElements.
	 * @throws Throwable on error.
	 */
	public List<WebElement> getWebElementList(String xpath, ExtentTest logInfo) throws Throwable {
		List<WebElement> list = null;
		try {
			WebDriverWait wait = new WebDriverWait(getDriver(), Duration.ofSeconds(40));
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty(xpath))));
			if (isElementPresent(xpath, logInfo)) {
				list = getDriver().findElements(By.xpath(prop.getProperty(xpath)));
			}
			return list;
		} catch (Exception e) {
			e.printStackTrace();
			addScreenshotAndLog("", "Failed to fetch webelements with xpath: " + xpath, logInfo);
			throw e;
		}
	}	
	/**
	 * Gets text for an element specified by property key, using Selenium or JS.
	 *
	 * @param element OR key for XPATH.
	 * @param logInfo ExtentTest logger.
	 * @return true if text fetched, false otherwise.
	 * @throws Throwable on error.
	 */
	public boolean getTextForORElement(String element, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		Boolean resultArr[] = new Boolean[10];
		int count = 0;
		try {
			if (isElementPresent(element, logInfo)) {
				if (element.toLowerCase().startsWith("sel")) {
					elementTextValue = getDriver().findElement(By.xpath(prop.getProperty(element))).getText();
					MyLogger.info(
							"Fetched text: '" + elementTextValue + "' for element " + element + " using selenium");
					resultArr[count++] = true;
				} else {
					JavascriptExecutor jse = (JavascriptExecutor) getDriver();
					elementTextValue = (String) jse.executeScript("return arguments[0].textContent;",
							getDriver().findElement(By.xpath(prop.getProperty(element))));
					MyLogger.info(
							"Fetched text: '" + elementTextValue + "' for element " + element + " using selenium");
					resultArr[count++] = true;
				}
			}
			result = !Arrays.asList(resultArr).contains(false);
		} catch (Exception e) {
			e.printStackTrace();
			addScreenshotAndLog("", "Fetch text fail : " + element, logInfo);
			throw e;
		}
		return result;
	}

	
	/**
	 * Gets text for a given WebElement, highlights it for visibility.
	 * Handles StaleElementReferenceException by retrying.
	 *
	 * @param elementName Name for reporting.
	 * @param element     WebElement instance.
	 * @param logInfo     ExtentTest logger.
	 * @return Text value of element.
	 * @throws Throwable on error.
	 */
	public String getTextForElement(String elementName, WebElement element, ExtentTest logInfo) throws Throwable {
		String textValue;
		try {
			scrollHighLight(element);
			textValue = element.getText();
			if (textValue != null) {
				MyLogger.info("Fetched text: '" + textValue + "' for element " + elementName + " using selenium");
			} else {
				MyLogger.info(
						"Failed to Fetch text: '" + textValue + "' for element " + elementName + " using selenium");
			}
		} catch (StaleElementReferenceException staleExcep) {
			try {
				scrollHighLight(element);
				textValue = element.getText();
				if (textValue != null) {
					MyLogger.info("Fetched text: '" + textValue + "' for element " + elementName + " using selenium");
				} else {
					MyLogger.info(
							"Failed to Fetch text: '" + textValue + "' for element " + elementName + " using selenium");
				}
			} catch (Exception e1) {
				throw e1;
			}
		} catch (Exception e) {
			e.printStackTrace();
			addScreenshotAndLog("", "Fetch text fail : " + elementName, logInfo);
			throw e;
		}
		return textValue;
	}

	
	/**
	 * Gets text for an element using an explicit XPATH, highlights for visibility.
	 *
	 * @param elementName Name for logging.
	 * @param xpath       XPATH locator.
	 * @param logInfo     ExtentTest logger.
	 * @return Text value.
	 * @throws Throwable on error.
	 */
	public String getTextForXpath(String elementName, String xpath, ExtentTest logInfo) throws Throwable {
		try {
			WebElement element = getDriver().findElement(By.xpath(xpath));
			scrollHighLight(element);
			String textValue = element.getText();
			if (textValue != null) {
				MyLogger.info("Fetched text: '" + textValue + "' for element " + elementName + " using selenium");
			} else {
				MyLogger.info(
						"Failed to Fetch text: '" + textValue + "' for element " + elementName + " using selenium");
			}
			return textValue;
		} catch (Exception e) {
			e.printStackTrace();
			addScreenshotAndLog("", "Fetch text fail : " + elementName, logInfo);
			throw e;
		}
	}

	
	/**
	 * Gets the count of elements found by OR key's XPATH.
	 *
	 * @param ele     OR key for XPATH.
	 * @param logInfo ExtentTest logger.
	 * @return true if count succeeded.
	 * @throws Throwable on error.
	 */
	public boolean getSize(String ele, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		try {
			size = getDriver().findElements(By.xpath(prop.getProperty(ele))).size();
			result = true;
		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		}
		return result;
	}
	/**
	 * Gets a specified attribute value for an element.
	 *
	 * @param ele       OR key for XPATH.
	 * @param attribute Attribute name.
	 * @param logInfo   ExtentTest logger.
	 * @return Attribute value, or null.
	 * @throws Throwable on error.
	 */
	public String getAttributeForElement(String ele, String attribute, ExtentTest logInfo) throws Throwable {
		String attributeValue = null;
		try {
			takescreenshot = "no";
			if (isElementPresent(ele, logInfo)) {
				attributeValue = getDriver().findElement(By.xpath(prop.getProperty(ele))).getAttribute(attribute);
				MyLogger.info(" Fetched attribute value: " + ele + " : " + attributeValue);
			}
		} catch (Exception e) {
			addScreenshotAndLog("Exception occurred : " + e.getMessage(), "Fetch attribute failed: " + ele, logInfo);
			throw e;
		}
		return attributeValue;
	}

	
	/**
	 * Checks if an attribute is present for a given element.
	 *
	 * @param ele       OR key for XPATH.
	 * @param attribute Attribute name.
	 * @param logInfo   ExtentTest logger.
	 * @return true if attribute present.
	 * @throws Throwable on error.
	 */
	public boolean isAtributePresentForElement(String ele, String attribute, ExtentTest logInfo) throws Throwable {
		boolean result = false;
		try {
			takescreenshot = "no";
			if (isElementPresent(ele, logInfo)) {
				String attributeValue = getDriver().findElement(By.xpath(prop.getProperty(ele)))
						.getAttribute(attribute);
				MyLogger.info(" Fetched attribute value: " + ele + " : " + attributeValue);
				result = true;
			}
		} catch (Exception e) {
			MyLogger.info(" Attribute :" + attribute + " is not present");
		}
		return result;
	}
	/**
	 * Waits for a 'loading' element to disappear (loading spinner etc.), logs wait.
	 *
	 * @param xpath   XPATH string or OR key.
	 * @param logInfo ExtentTest logger.
	 * @throws Exception on error.
	 */
	public void isLoading(String xpath, ExtentTest logInfo) throws Exception {
		boolean eleFind = false;
		WebElement ele;
		try {
			if (xpath.startsWith("//")) {
				ele = getDriver().findElement(By.xpath((xpath)));
			} else {
				ele = getDriver().findElement(By.xpath(prop.getProperty(xpath)));
			}
			if (ele.isDisplayed()) {
				eleFind = true;
			}
			if (eleFind) {
				MyLogger.info("Loading.............");
				WebDriverWait wait = new WebDriverWait(getDriver(), Duration.ofSeconds(10));
				wait.until(ExpectedConditions.invisibilityOf(ele));
				MyLogger.info("Waited until load complete");
			}
		} catch (Exception e) {
			MyLogger.info("Loaded fast");
		}
	}

	
	/**
	 * Waits until a given element is clickable, logs wait.
	 *
	 * @param xpath   XPATH string or OR key.
	 * @param logInfo ExtentTest logger.
	 * @throws Exception on error.
	 */
	public void waitUntilElementIsClickable(String xpath, ExtentTest logInfo) throws Exception {
		boolean eleNotClickable = false;
		WebElement ele;
		try {
			if (xpath.startsWith("xpath")) {
				ele = getDriver().findElement(By.xpath((xpath)));
			} else {
				ele = getDriver().findElement(By.xpath(prop.getProperty(xpath)));
			}
			if (!ele.isDisplayed()) {
				eleNotClickable = true;
			}
			if (eleNotClickable) {
				MyLogger.info("Wait.............");
				WebDriverWait wait = new WebDriverWait(getDriver(), Duration.ofSeconds(5));
				wait.until(ExpectedConditions.elementToBeClickable(ele));
				MyLogger.info("Waited until element clickable");
			}
		} catch (Exception e) {
			MyLogger.info("Element Clickable");
		}
	}

}