package com.vtiger.practice;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.testng.Reporter;
import org.testng.annotations.Test;

import com.sdet34l1.genericUtility.BaseClass;
import com.sdet34l1.genericUtility.ExcelOffice;
import com.sdet34l1.genericUtility.JavaOffice;
import com.vtiger.ObjectRepository.ContactInformationPage;
import com.vtiger.ObjectRepository.ContactPage;
import com.vtiger.ObjectRepository.CreateNewContactInformationPage;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CreateContactsTest extends BaseClass{
	@Test
	public void createContactsTest()throws EncryptedDocumentException, IOException
 {
		String expcontactname = ExcelOffice.getDataFromExcel("contacts", 2, 1)+JavaOffice.getRandomnumber(1000);
		homepage.HomePageContactLink();
		ContactPage cp = new ContactPage(driver);
		cp.contactPageAction();
		CreateNewContactInformationPage cncip = new CreateNewContactInformationPage(driver);
		cncip.contactInformationPageAction(expcontactname);
		ContactInformationPage cap=new ContactInformationPage(driver);
		String actual = cap.ContactInformationPageAction();
		JavaOffice.assertionthroughifCondition(actual, expcontactname, "contacts");
		Reporter.log("contacts has been created");
	}
}

