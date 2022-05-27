package com.sdet34l1.genericUtility;

import org.openqa.selenium.WebDriver;

import com.vtiger.ObjectRepository.HomePage;
import com.vtiger.ObjectRepository.LoginPage;

public class GettersandSettersForListeners 
{
private static GettersandSettersForListeners instance;
String username;
String password;
public long longtimeout;
public WebDriver driver;
public LoginPage loginpage;
public HomePage homepage;
public String filepath;

private GettersandSettersForListeners()
{
}
	public static GettersandSettersForListeners getInstance()
	{
		if(instance==null)
		{
			instance=new GettersandSettersForListeners();
		}
		return instance;
	}
	public String getUsername()
	{
		return username;
	}
	public String getPassword() 
	{
		return password;
	}
	public long getLongtimeout()
	{
		return longtimeout;
	}
	public WebDriver getDriver() 
	{
		return driver;
	}
	public LoginPage getLoginpage() 
	{
		return loginpage;
		
	}
	public HomePage getHomepage() 
	{
		return homepage;
	}
	public String getFilepath()
	{
		return filepath;
	}
}
