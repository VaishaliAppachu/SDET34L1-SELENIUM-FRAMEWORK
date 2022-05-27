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
import com.vtiger.ObjectRepository.ClearThefieldForPage;
import com.vtiger.ObjectRepository.ContactPage;
import com.vtiger.ObjectRepository.ContactsGetText;
import com.vtiger.ObjectRepository.CreateNewContactInformationPage;
import com.vtiger.ObjectRepository.HomePage;
import com.vtiger.ObjectRepository.LoginPage;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CreateContactandClickPreviousPageTest extends BaseClass{
	@Test
	public void createContactTC3() throws EncryptedDocumentException, IOException, InterruptedException
	{

		String contactname = ExcelOffice.getDataFromExcel("contacts", 2, 1)+JavaOffice.getRandomnumber(1000);

		if(driver.getTitle().contains("Home"))
		{
			ExcelOffice.setExcelData("contacts", 32, 7, "login page is displayed");
			ExcelOffice.setExcelData("contacts", 32, 8, "Pass");
		}		
		JavaOffice.printStatement("home page is dispalyed");
		homepage.HomePageContactLink();
		ContactPage cp = new ContactPage(driver);
		cp.contactPageAction();		
		JavaOffice.printStatement("home page is dispalyed");		
		cp.contactPageAction();		 
		if(driver.getTitle().contains("Contacts"))
		{
			ExcelOffice.setExcelData("contacts", 33, 7, "home page is displayed");
			ExcelOffice.setExcelData("contacts", 33, 8, "Pass");
		}		
		JavaOffice.printStatement("contact page is displayed");
		CreateNewContactInformationPage cncip = new CreateNewContactInformationPage(driver);
		cncip.contactInformationPageAction(contactname);		
		ContactsGetText text = new ContactsGetText(driver);
		text.getTextContact();					
		JavaOffice.printStatement("name is same");		
		if(driver.getTitle().contains("Contacts"))
		{
			ExcelOffice.setExcelData("contacts", 34, 7, "Contact page is displayed");
			ExcelOffice.setExcelData("contacts", 34, 8, "Pass");
		}        
		JavaOffice.printStatement("Sam contact information updated today");	
		cp.contactPageAction();
		if(driver.getTitle().contains("Contacts"))
		{			
			ExcelOffice.setExcelData("contacts", 35, 7, "Contact information updated today");
			ExcelOffice.setExcelData("contacts", 35, 8, "Pass");
		}	
		if(driver.getTitle().contains("Contacts"))
		{			
			ExcelOffice.setExcelData("contacts", 36, 7, "all contact pages displayed");
			ExcelOffice.setExcelData("contacts", 36, 8, "Pass");
		}	
	}
}