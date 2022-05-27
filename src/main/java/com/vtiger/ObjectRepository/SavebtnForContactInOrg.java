package com.vtiger.ObjectRepository;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class SavebtnForContactInOrg {
	@FindBy(xpath = "//input[@name='lastname']")
	private WebElement contactLastName;
	@FindBy(xpath = "//input[@value='  Save  ']")
	private WebElement contactSaveBtn;
	@FindBy(xpath="//input[@name='account_name']/..//img[@title='Select']")
	private WebElement plusForOrg;
	
	public SavebtnForContactInOrg(WebDriver driver) {
	PageFactory.initElements(driver, this);	
	}

 public void contactInformationPageAction(String lastName) {
	 contactLastName.sendKeys(lastName);
 }
 public void saveContact()
 {
	 contactSaveBtn.click();

 }
 public void plusForOrgName()
 {
	 plusForOrg.click();
 }
 
}
