package com.vtiger.ObjectRepository;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.sdet34l1.genericUtility.WebDriverOffice;

public class MailMergeTemplate {

	@FindBy(xpath="//img[@title='Mail Merge']")
	private WebElement mailMerige;
	
	@FindBy(xpath="//input[@class='crmButton create small']")
	private WebElement addTemplate;
	
	public MailMergeTemplate(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
	public void clickOnMailMerge(WebDriver driver)
	{
		mailMerige.click();
	}
	public void clickAddTemplates()
	{
		addTemplate.click();
	}
}
