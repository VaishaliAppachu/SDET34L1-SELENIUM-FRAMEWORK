package com.vtiger.ObjectRepository;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class CreateNewProductPage {
@FindBy(name="productname")
private WebElement lastName;
@FindBy(xpath="//input[@accesskey='S']") 
private WebElement saveProductLink;
	
	public CreateNewProductPage(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
	public void productLastName(String LastName)
	{
		lastName.sendKeys(LastName);
		saveProductLink.click();
	}
}
