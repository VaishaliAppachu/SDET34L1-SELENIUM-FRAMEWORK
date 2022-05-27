package com.vtiger.ObjectRepository;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class CreateProductPage {
@FindBy(xpath="//img[@alt='Create Product...']")
private WebElement ClickPlusForProduct;

	public CreateProductPage(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
	
	public void clickForProduct()
	{
		ClickPlusForProduct.click();
	}
	
}
