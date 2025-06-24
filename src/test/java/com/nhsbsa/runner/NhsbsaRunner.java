package com.nhsbsa.runner;

import org.testng.ITestContext;

import org.testng.annotations.BeforeClass;

import com.nhsbsa.util.MyLogger;

import io.cucumber.testng.CucumberOptions;

@CucumberOptions(plugin = { "pretty", "html:target/site/cucumber-pretty", "json:target/cucumber/cucumber.json", "com.nhsbsa.steps.StepListener", "timeline:test-output-thread/" }, features = {
		"src/test/java/com/nhsbsa/features/" }, glue = {
				"com/nhsbsa/steps" }, monochrome = true, publish = false, tags = "@happy-path or  @edge-case or @filter or @pagination")
public class NhsbsaRunner extends CustomAbstractTestNGCucumberTest {
	@BeforeClass(alwaysRun = true)
	public void setUpClassLog(ITestContext context) {
		MyLogger.startTest(this.getClass().getSimpleName() + "_" + context.getName());
	}
}
