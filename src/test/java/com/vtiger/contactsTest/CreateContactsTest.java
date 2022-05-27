package com.vtiger.contactsTest;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;
import java.util.Random;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.Reporter;
import org.testng.annotations.Test;

import com.sdet34l1.genericUtility.BaseClass;
import com.sdet34l1.genericUtility.ExcelOffice;
import com.sdet34l1.genericUtility.FileOffice;
import com.sdet34l1.genericUtility.IconstantPathOffice;
import com.sdet34l1.genericUtility.JavaOffice;
import com.sdet34l1.genericUtility.WebDriverOffice;
import com.vtiger.ObjectRepository.ContactInformationPage;
import com.vtiger.ObjectRepository.ContactPage;
import com.vtiger.ObjectRepository.CreateNewContactInformationPage;
import com.vtiger.ObjectRepository.HomePage;
import com.vtiger.ObjectRepository.LoginPage;

import io.github.bonigarcia.wdm.WebDriverManager;
import io.qameta.allure.Description;
import io.qameta.allure.Epic;
import io.qameta.allure.Severity;
import io.qameta.allure.SeverityLevel;
import io.qameta.allure.Step;
import io.qameta.allure.Story;

public class CreateContactsTest extends BaseClass{
	@Test(groups = "Regression",description ="testng:- CreateContactsTest")
    @Description("Desription:- CreateContactsTest")
	@Epic("Epic:- CreateContactsTest")
	@Story("Story:- CreateContactsTest")
	@Step("Step:- CreateContactsTest")
	@Severity(SeverityLevel.BLOCKER)

	public void fetchDataofContactsTest() throws EncryptedDocumentException, IOException
	{
		String expcontactname = ExcelOffice.getDataFromExcel("contacts", 2, 1)+JavaOffice.getRandomnumber(1000);
		homepage.HomePageContactLink();
		ContactPage contactpage = new ContactPage(driver);
		contactpage.contactPageAction();
		CreateNewContactInformationPage createnewcontactinformationPage = new CreateNewContactInformationPage(driver);
		createnewcontactinformationPage.contactInformationPageAction(expcontactname);
		ContactInformationPage contactinformationpage=new ContactInformationPage(driver);
		String actual = contactinformationpage.ContactInformationPageAction();
		JavaOffice.assertionthroughifCondition(actual, expcontactname, "contacts");
		Reporter.log("contacts has been created");
	}
}


