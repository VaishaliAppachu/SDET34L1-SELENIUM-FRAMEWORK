package com.vtiger.ObjectRepository;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.sdet34l1.genericUtility.JavaOffice;

public class SwitchToWindowForOrg {

	@FindBy(name="search_text")
	private WebElement searchTxt;
	
	@FindBy(name="search")
	private WebElement searchClick;

	@FindBy(xpath="//tr[2]/td/a")
	private WebElement actualorg;
	
	public SwitchToWindowForOrg(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
	
   public void searchText(String text) throws InterruptedException
   {
       searchTxt.sendKeys(text);
   }
   public void clickSearch() throws InterruptedException
   {
	   searchClick.click();
   }
   public void actualOrg(WebDriver driver, String orgname)
   {
	   driver.findElement(By.linkText(orgname)).click();
   }
	
}

	
	
	
	
	
	
	
	