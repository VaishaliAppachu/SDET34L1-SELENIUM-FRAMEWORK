package com.vtiger.ObjectRepository;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class OrgInformationValidateDropDown {

	@FindBy(xpath="//span[@id='dtlview_Industry']")
	private WebElement indusName;
	
	@FindBy(xpath="//span[@id='dtlview_Type']")
	private WebElement accountType;
	
	public OrgInformationValidateDropDown(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
	
	public String dropdownofIndus()
	{
		return indusName.getText();
	}
	public String dropdownofAccount()
	{
		return accountType.getText();
	}
}
