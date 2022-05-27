package com.vtiger.practice;

import org.testng.Reporter;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class TestNgBasicConfAnnotationTest {

//	@Test
//	public void test1Test()
//	{
//		Reporter.log("test1",true);
//	}
//	@Test
//		public void test2Test()
//		{
//			Reporter.log("test2",true);
//		}
		
	
	@BeforeSuite(groups = "baseclass")
	public void beforeSuite1Test()
	{
		Reporter.log("beforesuite1",true);
		
	}
	
	@AfterSuite (groups = "baseclass")
	public void aftersuite1Test()
	{
		Reporter.log("aftersuite1",true);

	}
	@BeforeClass(groups = "baseclass")
	public void beforeClass1Test()
	{
		Reporter.log("beforeclass1",true);
		
	}
	@AfterClass(groups = "baseclass")
		public void afterclass1Test()
		{
			Reporter.log("afterclass1",true);
			
		}
	@BeforeTest(groups = "baseclass")
	public void beforeTest1Test()
	{
		Reporter.log("beforeTest1",true);
		
	}
	@AfterTest(groups = "baseclass")
	public void afterTest1Test()
	{
		Reporter.log("afterTest1",true);
		
	}
	@BeforeMethod(groups = "baseclass")
	public void beforeMethod1Test()
	{
		Reporter.log("beforeMethod1",true);
		
	}
	@AfterMethod(groups = "baseclass")
	public void afterMethod1Test()
	{
		Reporter.log("afterMethod1",true);
		
	}
	@BeforeMethod(groups = "baseclass")
	public void beforeMethod2Test()
	{
		Reporter.log("beforeMethod2",true);
		
	}
	}

