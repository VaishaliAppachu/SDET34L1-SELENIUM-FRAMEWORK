package com.sdet34l1.genericUtility;

import org.openqa.selenium.WebDriver;
import org.testng.ITestContext;
import org.testng.ITestListener;
import org.testng.ITestResult;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;
import com.vtiger.ObjectRepository.HomePage;
import com.vtiger.ObjectRepository.LoginPage;
/***
 * 
 * @author VAISHALI
 *
 */
public class ListenerImplementation implements ITestListener
{
	 ExtentReports report; 
	 ExtentTest test;
	 
	 /**
	  * This method gets invoked before any test method gets invoked
	  */
	 @Override
		public void onStart(ITestContext context)
		{
		    

		    System.out.println("onStart");	
		   
		    ExtentSparkReporter spark=new ExtentSparkReporter("./extentReport/extentReport.html");
		    spark.config().setDocumentTitle("SDET34L1");
		    spark.config().setTheme(Theme.DARK);
		    spark.config().setReportName("TestNG Report Name");
		    
		    report = new ExtentReports();
		    report.attachReporter(spark);
		    report.setSystemInfo("Environment", "Testing Environment");
		    report.setSystemInfo("Reporter Name", "Vaishali Appachu");
		    report.setSystemInfo("Platform", "windows 10");
		    report.setSystemInfo("Unit Testing Tool", "TESTNG");
		    report.setSystemInfo("Build Management Tool", "Maven");
		    report.setSystemInfo("Automation Tool", "Selenium");	       
		}
 

	@Override
	public void onTestStart(ITestResult result)
	{
    System.out.println("onTestStart");	  
    test = report.createTest(result.getMethod().getMethodName());
	}
	/**
	 * This method is called on the success of any Test
	 */
	@Override
	public void onTestSuccess(ITestResult result)
	{
	    System.out.println("onTestSuccess");		
	    test.log(Status.PASS, result.getMethod().getMethodName()+" is pass");
	}
	/**
	 * This method is called on the failure of any Test
	 */
	@Override
	public void onTestFailure(ITestResult result) 
	{
	    System.out.println("onTestFailure");
	    test.log(Status.FAIL, result.getMethod().getMethodName()+ "is fail");
	    test.log(Status.FAIL, result.getThrowable());
	    	    
	  test.addScreenCaptureFromPath(WebDriverOffice.takesScreenshot(result.getMethod().getMethodName(),BaseClass.staticdriver));
	}
/**
 * This method is called on skipped of any Test
 */
	@Override
	public void onTestSkipped(ITestResult result) 
	{
	    System.out.println("onTestSkipped");	
		test.log(Status.SKIP, result.getMethod().getMethodName()+ "is skipped");
		test.log(Status.SKIP, result.getThrowable());
	}
/**
 * This method is called on each time Test fails but is within success percentage
 */
	@Override
	public void onTestFailedButWithinSuccessPercentage(ITestResult result)
	{
	    System.out.println("onTestFailedButWithinSuccessPercentage");	
		
	}
/**
 * This method gets invoked after all test method gets invoked
 */
	@Override
	public void onFinish(ITestContext context) 
	{
	    System.out.println("onFinish");	
	    report.flush();
		
	}

}
