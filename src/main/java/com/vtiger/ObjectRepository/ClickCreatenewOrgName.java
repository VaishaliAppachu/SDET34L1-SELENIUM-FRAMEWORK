package com.vtiger.ObjectRepository;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class ClickCreatenewOrgName {

	@FindBy(xpath="//img[@alt='Create Organization...']")
	private WebElement clickplus;
	
	public ClickCreatenewOrgName (WebDriver driver1 )
	{
		PageFactory.initElements(driver1, this);
	}
	
	public void orgClickAction()
	{
		clickplus.click();
	}
	
}
