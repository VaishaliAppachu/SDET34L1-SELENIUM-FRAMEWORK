package com.vtiger.practice;

import static org.testng.Assert.assertTrue;

import org.testng.Assert;
import org.testng.Reporter;
import org.testng.annotations.Test;

public class TestNgAssert 
{
	@Test
	public static void practice1() {
		Reporter.log("pass1",true);
		Assert.assertEquals("abc", "abc");
	}
	@Test(dependsOnMethods = "practice1")
	public static void practice2() {
		Reporter.log("pass2",true);
		Assert.assertEquals("efg", "456");
	}
	@Test(invocationCount = 2)
	public static void practice3() {
		Reporter.log("pass3",true);
		Assert.assertEquals("hij", "789");
	}
}