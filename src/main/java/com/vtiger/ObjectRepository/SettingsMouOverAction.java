package com.vtiger.ObjectRepository;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.sdet34l1.genericUtility.WebDriverOffice;

public class SettingsMouOverAction {

	@FindBy(xpath="//img[@src='themes/softed/images/mainSettings.PNG']")
	private WebElement settingsMouOver;
	
	@FindBy(xpath="//a[.='CRM Settings']")
	private WebElement clickOnSettings;
	
	public SettingsMouOverAction(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
	
	public void mouOverOnSettings(WebDriver driver)
	{
		WebDriverOffice.mouseHoverOntheElement(settingsMouOver, driver);
		clickOnSettings.click();
	}
	
	
}
