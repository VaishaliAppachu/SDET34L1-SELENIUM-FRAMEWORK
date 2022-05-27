package com.vtiger.ObjectRepository;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.sdet34l1.genericUtility.JavaOffice;
import com.sdet34l1.genericUtility.WebDriverOffice;

public class HomePage {
	@FindBy(xpath ="//a[.='Organizations']")
	private WebElement OrganizationsLink;
	@FindBy(xpath = "//img[@src='themes/softed/images/user.PNG']")
	private WebElement adminstrationIcon;
	@FindBy(xpath = "//a[.='Sign Out']")
	private WebElement signOutLnk;
	@FindBy(xpath = "//a[.='Documents']")
	private WebElement documentLink;
	@FindBy(xpath = "//a[.='Contacts']")
	private WebElement contactLink;
	@FindBy(xpath="//a[.='Products']")
	private WebElement productLink;
	@FindBy (id="qccombo")
	private WebElement quickCreate;

	
	public HomePage(WebDriver driver) {
		PageFactory.initElements(driver, this);
	}
	
	public void ClickOnQuickCreate()
	{
		quickCreate.click();
	}
	
	public void homePageForProduct()
	{
       productLink.click();
	}
	public void homePageOrganizationAction() {
		OrganizationsLink.click();
	}
	
	public void signOutIconMouseHoverAction(WebDriver driver) {
		JavaOffice.customWait(adminstrationIcon, 2, 10);
		WebDriverOffice.initializeActionClass(driver);
		WebDriverOffice.mouseHoverOntheElement(adminstrationIcon, driver);
		signOutLnk.click();
	}
	
	public void homePageDocumentLink() {
		documentLink.click();
	}
	public void HomePageContactLink() {
		JavaOffice.customWait(contactLink, 3, 10);
		contactLink.click();
	}
	public void quickCreateDrop(String name)
	{
		WebDriverOffice.selectByVisibleText(name, quickCreate);
	}
}
