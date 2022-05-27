package com.vtiger.ObjectRepository;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.sdet34l1.genericUtility.WebDriverOffice;

public class ClickOnCampaign {
	@FindBy(xpath="//img[@src='themes/softed/images/menuDnArrow.gif']")
    private WebElement MouseOverOnMore;
	
	@FindBy(name="Campaigns")
	private WebElement CampaignLink;
	
	public ClickOnCampaign(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
	public void clickCampainLink(WebDriver driver)
	{
		WebDriverOffice.mouseHoverOntheElement(MouseOverOnMore, driver);
        CampaignLink.click();
	}
}
