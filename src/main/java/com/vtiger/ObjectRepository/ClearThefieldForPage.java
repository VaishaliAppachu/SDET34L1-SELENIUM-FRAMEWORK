package com.vtiger.ObjectRepository;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class ClearThefieldForPage {

	@FindBy(name="pagenum")
	private WebElement ClearAndSend;
	
	
	
	public ClearThefieldForPage(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
	 public void SendInt(String number)
	 {
		ClearAndSend.clear();
		ClearAndSend.sendKeys(number);
	 }

}
