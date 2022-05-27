package com.vtiger.ObjectRepository;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.sdet34l1.genericUtility.JavaOffice;

public class QuickCreateDropDown {

	@FindBy(name="lastname")
	private WebElement lastNameofCon;
	
	@FindBy(xpath="//table[@class='qcTransport']//input[@accesskey='X']")
	private WebElement cancelButton;
	
	
	public QuickCreateDropDown(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
	
	public void clickOnQuickCreate(String lastName)
	{
		JavaOffice.customWait(lastNameofCon, 2, 10);
		lastNameofCon.sendKeys(lastName);
	}
	public void clearTheField()
	{
		lastNameofCon.clear();
	}
	public void cancelBtn()
	{
		cancelButton.click();
	}
	
}
