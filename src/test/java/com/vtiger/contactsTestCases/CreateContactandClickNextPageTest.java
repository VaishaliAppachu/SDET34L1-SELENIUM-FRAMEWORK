package com.vtiger.contactsTestCases;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.Test;

import com.sdet34l1.genericUtility.BaseClass;
import com.sdet34l1.genericUtility.ExcelOffice;
import com.sdet34l1.genericUtility.FileOffice;
import com.sdet34l1.genericUtility.IconstantPathOffice;
import com.sdet34l1.genericUtility.JavaOffice;
import com.sdet34l1.genericUtility.WebDriverOffice;
import com.vtiger.ObjectRepository.ContactPage;
import com.vtiger.ObjectRepository.ContactsGetText;
import com.vtiger.ObjectRepository.CreateNewContactInformationPage;
import com.vtiger.ObjectRepository.HomePage;
import com.vtiger.ObjectRepository.LoginPage;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CreateContactandClickNextPageTest extends BaseClass{
	@Test
	public void createContactTC2() throws EncryptedDocumentException, IOException
	{
		String contactname = ExcelOffice.getDataFromExcel("contacts", 2, 1)+JavaOffice.getRandomnumber(1000);

		if(driver.getTitle().contains("vtiger"))
		{
			ExcelOffice.setExcelData("contacts", 25, 7, "login page is displayed");
			ExcelOffice.setExcelData("contacts", 25, 8, "Pass");

		}
		JavaOffice.printStatement("login page is dispalyed");
		if(driver.getTitle().contains("Home"))
		{
			ExcelOffice.setExcelData("contacts", 26, 7, "home page is displayed");
			ExcelOffice.setExcelData("contacts", 26, 8, "Pass");

		}
		JavaOffice.printStatement("home page is dispalyed");
		homepage.HomePageContactLink();
		ContactPage cp = new ContactPage(driver);
		cp.contactPageAction();	 
		if(driver.getTitle().contains("Contacts"))
		{
			ExcelOffice.setExcelData("contacts", 27, 7, "contact page is displayed");
			ExcelOffice.setExcelData("contacts", 27, 8, "Pass");
		}
		JavaOffice.printStatement("contact page is displayed");		
		CreateNewContactInformationPage cncip = new CreateNewContactInformationPage(driver);
		cncip.contactInformationPageAction(contactname);        
		ContactsGetText text = new ContactsGetText(driver);
		text.getTextContact();
		JavaOffice.printStatement("name is same");		
		if(driver.getTitle().contains("Contacts"))
		{
			ExcelOffice.setExcelData("contacts", 28, 7, "Contact information updated today");
			ExcelOffice.setExcelData("contacts", 28, 8, "Pass");
		}    
		if(driver.getTitle().contains("Contacts"))
		{
			ExcelOffice.setExcelData("contacts", 29, 7, "previous contact is displayed");
			ExcelOffice.setExcelData("contacts", 29, 8, "Pass");
		}    
		JavaOffice.printStatement("Sam contact information updated today");
	}
}

