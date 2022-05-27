package com.vtiger.ObjectRepository;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.sdet34l1.genericUtility.JavaOffice;

public class SearchInsideForProduct {
@FindBy(name="search_text")
private WebElement clickSearchProduct;

@FindBy(xpath="//input[@name='search']")
private WebElement  searchForProduct;

@FindBy(xpath="//tr[2]/td/a")
private WebElement currentProduct;


   public SearchInsideForProduct(WebDriver driver)
   {
	   PageFactory.initElements(driver, this);
   }
	
	public void searchTheCreatedProduct(String productName) throws InterruptedException
	{
		clickSearchProduct.sendKeys(productName);
		searchForProduct.click();
		Thread.sleep(5000);

	}
	public void   getTheCreatedProduct() throws InterruptedException
	{
		Thread.sleep(5000);
		currentProduct.click();
	}
	
}
