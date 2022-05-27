package com.vtiger.ObjectRepository;

	import org.openqa.selenium.WebDriver;
	import org.openqa.selenium.WebElement;
	import org.openqa.selenium.support.FindBy;
	import org.openqa.selenium.support.PageFactory;

	public class CreateNewContactInformationPage {
		@FindBy(xpath = "//input[@name='lastname']")
		private WebElement contactLastName;
		@FindBy(xpath = "//input[@value='  Save  ']")
		private WebElement contactSaveBtn;
		
		
		public CreateNewContactInformationPage(WebDriver driver) {
		PageFactory.initElements(driver, this);	
		}

	 public void contactInformationPageAction(String lastName) {
		 contactLastName.sendKeys(lastName);
		 contactSaveBtn.click();
	 }
	}

