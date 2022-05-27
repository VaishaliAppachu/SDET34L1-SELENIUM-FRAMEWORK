package com.vtiger.ObjectRepository;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.sdet34l1.genericUtility.JavaOffice;
import com.sdet34l1.genericUtility.WebDriverOffice;

public class CreateNewOrgInfoPage {

	@FindBy(xpath="//input[@name='accountname']")
	private WebElement newOrgName;
	
	@FindBy(name="industry")
	private WebElement industryNameDropdown;
	
	@FindBy(name="accounttype")
	private WebElement accountTypeDropdown;
	
	@FindBy(xpath = "//img[@title='Create Organization...']")
	private WebElement organizationSelectBtn;
	
	@FindBy(xpath="//input[@accesskey='S']")
	private WebElement savebtn;
	
	@FindBy(xpath="//span[@class='dvHeaderText']")
	private WebElement title;
	
	public CreateNewOrgInfoPage(WebDriver driver1)
	{
		PageFactory.initElements(driver1, this);
	}
	public void selectOrg()
	{
		organizationSelectBtn.click();

	}
	public void orgLastName(String lastName) 
	{
		

		newOrgName.sendKeys(lastName);
	}
	
	public void industryNameOfOrg(String industryName )
	{
		WebDriverOffice.selectByVisibleText(industryName, industryNameDropdown);

	}
	public void accountTypeOfOrg(String accountType)
	{
		WebDriverOffice.selectByValue(accountTypeDropdown, accountType);

	}
	public void saveBtnToOrg(WebDriver driver) 
	{

		savebtn.click();
		WebDriverOffice.WaitvisibleText(driver, 10, title);
}
}