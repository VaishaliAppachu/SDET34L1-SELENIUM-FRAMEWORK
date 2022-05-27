package com.vtiger.practice;

import java.io.IOException;

import org.testng.Reporter;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class DataProviderTestNgPracticeTest 
{
@Test
(dataProvider = "loginData")
	public void testNgPracticeTest(String un,String pwd) throws  IOException, InterruptedException 
	{
		Reporter.log(un+"----->   "+pwd,true);
	}
@DataProvider
public Object[][] loginData()
{
	Object[][] data=new Object[5][2];
	data[0][0]="username";
	data[0][1]="password";

	data[1][0]="username";
	data[1][1]="password";

	data[2][0]="username";
	data[2][1]="password";

	data[3][0]="username";
	data[3][1]="password";

	data[4][0]="username";
	data[4][1]="password";
    
	return data;
}
}
