package com.vtiger.practice;

import org.testng.Reporter;
import org.testng.annotations.Test;
import org.testng.reporters.jq.Main;

public class TestNgAnnotPracticeTest extends TestNgBasicConfAnnotationTest{

	@Test(groups = "Regression")
	public void practiceTest()
	{
		Reporter.log("Hii",true);
	}
	@Test(groups = "Sanity")
	public void practice2Test()
	{
		Reporter.log("Helloooo");
	}

@Test(groups = {"Sanity","Regression"})
public void practice3Test()
{
	Reporter.log("Helloooo");
}

}