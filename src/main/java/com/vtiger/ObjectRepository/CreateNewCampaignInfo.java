package com.vtiger.ObjectRepository;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class CreateNewCampaignInfo {
	@FindBy(name="campaignname")
	private WebElement CampaignlastName;
	
	@FindBy(xpath="//input[@accesskey='S']")
	private WebElement Savecampaign;
	
	@FindBy(xpath="//input[@name='product_name']/..//img[@title='Select']")
	private WebElement clickForProduct;
	
	public CreateNewCampaignInfo(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}

	public void CreateCampaignlastName(String lastName)
	{
		CampaignlastName.sendKeys(lastName);
	}
	public void campaingSave()
	{
		Savecampaign.click();

	}
    public void clickForProductSearch()
    {
    	clickForProduct.click();
    }
}
