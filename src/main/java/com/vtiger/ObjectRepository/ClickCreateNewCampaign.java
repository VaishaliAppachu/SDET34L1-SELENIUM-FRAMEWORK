package com.vtiger.ObjectRepository;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.sdet34l1.genericUtility.JavaOffice;

public class ClickCreateNewCampaign {
@FindBy(xpath="//img[@alt='Create Campaign...']")
  private WebElement clickonplus;

   public ClickCreateNewCampaign(WebDriver driver)
   {
	   PageFactory.initElements(driver, this);
   }

	public void ClickonPlus()
	{
		JavaOffice.customWait(clickonplus, 2, 10);
		clickonplus.click();
	}
}
