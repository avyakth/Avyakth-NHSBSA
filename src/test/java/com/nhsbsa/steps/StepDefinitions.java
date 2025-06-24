package com.nhsbsa.steps;

import java.io.File;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.Duration;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.stream.Collectors;
import org.openqa.selenium.Keys;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import com.aventstack.extentreports.ExtentTest;
import com.nhsbsa.util.ExcelUtilResponse;
import com.nhsbsa.util.ITestListenerImpl;
import com.nhsbsa.util.MyLogger;
import com.nhsbsa.util.ResponseToRequest;

import io.cucumber.datatable.DataTable;
import io.cucumber.java.After;
import io.cucumber.java.Before;
import io.cucumber.java.Scenario;
import io.cucumber.java.en.And;
import io.cucumber.java.en.Given;
import io.cucumber.java.en.Then;
import io.cucumber.java.en.When;

public class StepDefinitions extends CommonFunctions {
	boolean result;
	ExtentTest extentStep;
	Properties prop = getORProp();
	WebDriver driver = getDriver();
	public String validEmailAddress = null;
	public String memberForRoleFilterCheck = null;
	public String TCName = null;
	public String scenarioName;
	public String status = null; 
	@Before
	public void startScenario(Scenario Scenario) {
		Collection<String> tags = Scenario.getSourceTagNames();
		MyLogger.info("");
		MyLogger.info("");
		String[] TCNames = Scenario.getName().split("_");
		TCName = TCNames[0];
		System.out.println("Scenario Name from feature file extraction" +TCName);
		MyLogger.info("******Start Scenario: " + Scenario.getName() + " " + tags.toString() + " : "
				+ getBrowserName().toUpperCase() + "  " + Thread.currentThread().getId() + " ******");
		scenarioName = Scenario.getName();
		scenario = extent.createTest(Scenario.getName()).assignDevice(getBrowserName().toUpperCase());
	}
	



	public void beforeStep() {
		MyLogger.info("");
		MyLogger.info("======Start Step: " + StepListener.stepName.get() + " : " + getBrowserName().toUpperCase() + " "
				+ Thread.currentThread().getId() + " =====");
		// if (!StepListener.stepName.get().contains("verif")) {
		// extentStep = scenario.createNode(StepListener.stepName.get());
		// }
		extentStep = scenario.createNode(StepListener.stepName.get());
		result = false;
	}

	public void afterStep(boolean result) throws Throwable {
		MyLogger.info("======End Step: " + StepListener.stepName.get() + " : " + getBrowserName().toUpperCase() + " "
				+ Thread.currentThread().getId() + " =====");
		if (!result && getDefaultProp().getProperty("FailedScreenshots").equalsIgnoreCase("Y")) {
			extentStep.addScreenCaptureFromPath(captureScreenShot(driver), StepListener.stepName.get());
		}
		MyLogger.info("=========step result======== "+result);
		logReport(result, StepListener.stepName.get(), extentStep, null);

	}

	public void afterStep(boolean result, Exception e) throws Throwable {
		MyLogger.info("======End Step: " + StepListener.stepName.get() + " : " + getBrowserName().toUpperCase() + " "
				+ Thread.currentThread().getId() + " =====");
		if (getDefaultProp().getProperty("FailedScreenshots").equalsIgnoreCase("Y")) {
			extentStep.addScreenCaptureFromPath(captureScreenShot(driver), StepListener.stepName.get());
		}
		MyLogger.info("=========step result======== "+result);
		logReport(result, StepListener.stepName.get(), extentStep, e);
	}

	@After
	public void closeScenario(Scenario Scenario) throws Exception {
		MyLogger.info("******End Scenario: " + Scenario.getName() + " : " + getBrowserName().toUpperCase() + "  "
				+ Thread.currentThread().getId() + "******");
		String scenarioStatus = scenario.getStatus().toString();
		boolean scenarioStatusFinal = scenarioStatus.equalsIgnoreCase("PASS") ? true : false;
		MyLogger.info("==========Scenario Status " + Scenario.getStatus() + "================");
		if (Scenario.getStatus().toString().equalsIgnoreCase("UNDEFINED")) {
			MyLogger.info("***>> Scenario '" + Scenario.getName() + "' failed at line(s) " + Scenario.getLine()
					+ " with status '" + Scenario.getStatus() + "");
			scenario.fail("SCENARIO STATUS : " + Scenario.getStatus().toString());
		}
		// rg.get().report(scenarioStatusFinal, "", Scenario.getName());
	}

	public void logReport(boolean result, String step, ExtentTest logInfo, Exception e) {
		try {
			
			if (result) {
				logResult("PASS", driver, logInfo, step, e);
				Assert.assertTrue(true);
			} else if (!result) {
				logResult("FAIL", driver, logInfo, step, e);
				Assert.fail(step, e);
			}
		} catch (Exception e1) {
			logResult("FAIL", driver, logInfo, step, e1);
			Assert.fail(step, e);
		}
	}
	
	@Given("I am on the NHS Jobs search page")
	public void user_launches_NHS_application(DataTable table) throws Throwable {
		try {
			this.beforeStep();
			Map<String, String> data = table.asMap(String.class, String.class);
			Boolean resultArr[] = new Boolean[10];
			resultArr[0] = launchURL(data.get("url"),extentStep);
			result = !Arrays.asList(resultArr).contains(false);
			this.afterStep(result);
		} catch (Exception e) {
			e.printStackTrace();
			this.afterStep(result, e);
		}
	}
	
	  @When("I search for jobs with:")
	  public void i_search_for_jobs_with(DataTable table) throws Throwable {
	    boolean[] resultArr = new boolean[10];
	    boolean result = false;
	    try {
	      this.beforeStep();
	      List<Map<String,String>> rows = table.asMaps(String.class, String.class);
	      Map<String,String> data = rows.get(0);        // your single row 	
	      resultArr[0] = clickOptional("sel_accept_cookies", extentStep);
	      resultArr[1] = click("sel_more_search_options", extentStep);
	      if ((data.get("keyword") != null && !(data.get("keyword").isEmpty()))) {
	      resultArr[2] = entersValue(data.get("keyword"),"sel_search_keyword", extentStep);
	      }
	      if ((data.get("location") != null && !(data.get("location").isEmpty()))) {
	      resultArr[3] = entersAriaValue(data.get("location"),"sel_search_location", extentStep);
	      }
	      if ((data.get("distance") != null && !(data.get("distance").isEmpty()))) {
	      resultArr[4] = selectDropdown(data.get("distance"),"sel_search_distance","text", extentStep);
	      }
	      if ((data.get("employer") != null && !(data.get("employer").isEmpty()))) {
	      resultArr[5] = entersValue(data.get("employer"),"sel_search_epmloyer", extentStep);
	      }
	      if ((data.get("payrange") != null && !(data.get("payrange").isEmpty()))) {
	      resultArr[6] = selectDropdown(data.get("payrange"),"sel_search_payrange","text", extentStep);
	      }
	      resultArr[7] = click("search_button", extentStep);
	      result = !Arrays.asList(resultArr).contains(false);
	      this.afterStep(result);
	    } catch (Exception e) {
	    	e.printStackTrace();
	      this.afterStep(result, e);
	    }
	  }

	  @Then("I sort my search results with the {string}")
	  public void then_i_sort_my_search_results_with_the_newest_date_posted(String sortBy)throws Throwable {
		  boolean[] resultArr = new boolean[10];
		    boolean result = false;
		    try {
		      this.beforeStep();
		      resultArr[0] = selectDropdown(sortBy,"sel_sortBy","text", extentStep);	  
		      result = !Arrays.asList(resultArr).contains(false);
		      this.afterStep(result);
		    }
		    catch (Exception e) {
		    	e.printStackTrace();
		    	this.afterStep(result, e);
		    }
		  }
	  
	  @Then("I should see a No result found message")
	  public void then_i_should_see_a_no_result_found_message()throws Throwable {
		  boolean[] resultArr = new boolean[10];
		    boolean result = false;
		    try {
		      this.beforeStep();
		      resultArr[0] = isElementPresent("sel_no_result",extentStep);
		      result = !Arrays.asList(resultArr).contains(false);
		      this.afterStep(result);
		    }
		    catch (Exception e) {
		    	e.printStackTrace();
		    	this.afterStep(result, e);
		    }
		  }
	  
	  @Then("I should see only jobs from employer {string}")
	  public void then_i_should_see_jobs_from_employer(String employer)throws Throwable {
		  boolean[] resultArr = new boolean[10];
		    boolean result = false;
		    try {
		      this.beforeStep();
		      resultArr[0] = verifyEmployerSearchResults(employer,extentStep);
		      result = !Arrays.asList(resultArr).contains(false);
		      this.afterStep(result);
		    }
		    catch (Exception e) {
		    	e.printStackTrace();
		    	this.afterStep(result, e);
		    }
		  }
	  
	  @Then("results should be sorted by newest date posted")
	  public void then_results_should_be_sorted_by_newest_Date_Posted() throws Throwable {
		  boolean[] resultArr = new boolean[10];
		    boolean result = false;
		    try {
		      this.beforeStep();
		      resultArr[0] = verifySortedByDateDesc(extentStep);	  
		      result = !Arrays.asList(resultArr).contains(false);
		      this.afterStep(result);
		    }
		    catch (Exception e) {
		    	e.printStackTrace();
		    	this.afterStep(result, e);
		    }
		  }
	  
	  @Then("I should see only jobs matching {string}")
	  public void then_i_should_see_only_jobs_matching(String keyword)throws Throwable {
		  boolean[] resultArr = new boolean[10];
		    boolean result = false;
		    try {
		      this.beforeStep();
		      resultArr[0] = verifySearchResults(keyword,extentStep);
		      result = !Arrays.asList(resultArr).contains(false);
		      this.afterStep(result);
		    }
		    catch (Exception e) {
		    	e.printStackTrace();
		    	this.afterStep(result, e);
		    }
		  }
	  
	  @Then("I should be able to traverse to each Next page until there are none")
	    public void thenIShouldTraverseAllNextPages() throws Throwable {
	        boolean[] resultArr = new boolean[10];
	        boolean result = false;
	        try {
	            this.beforeStep();
	            resultArr[0] = traverseAllPages(extentStep);
	            result = ! Arrays.asList(resultArr).contains(false);
	            this.afterStep(result);
	        } catch (Exception e) {
	            e.printStackTrace();
	            this.afterStep(result, e);
	        }
	    }

}