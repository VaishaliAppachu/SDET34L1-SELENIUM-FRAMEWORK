package com.vtiger.ObjectRepository;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class ContactsGetText {

	@FindBy(id="dtlview_Last Name")
	private WebElement lastname;
	
	public ContactsGetText(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
   public void getTextContact()
   {
	   lastname.getText();
   }
}
