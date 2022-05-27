package com.vtiger.ObjectRepository;

	import org.openqa.selenium.WebDriver;
	import org.openqa.selenium.WebElement;
	import org.openqa.selenium.support.FindBy;
	import org.openqa.selenium.support.PageFactory;

	public class ContactInformationPage {
		@FindBy(id="dtlview_Last Name")
		private WebElement actualContactLastName;
		
		public ContactInformationPage(WebDriver driver) {
			PageFactory.initElements(driver, this);
		}
		
		public  String ContactInformationPageAction() {
			 return actualContactLastName.getText();	
		}

	}


