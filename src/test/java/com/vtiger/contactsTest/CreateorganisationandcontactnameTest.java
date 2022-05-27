package com.vtiger.contactsTest;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Properties;
import java.util.Random;
import java.util.Set;
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
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.Reporter;
import org.testng.annotations.Test;

import com.sdet34l1.genericUtility.BaseClass;
import com.sdet34l1.genericUtility.ExcelOffice;
import com.sdet34l1.genericUtility.FileOffice;
import com.sdet34l1.genericUtility.IconstantPathOffice;
import com.sdet34l1.genericUtility.JavaOffice;
import com.sdet34l1.genericUtility.WebDriverOffice;
import com.vtiger.ObjectRepository.ClickCreatenewOrgName;
import com.vtiger.ObjectRepository.ContactPage;
import com.vtiger.ObjectRepository.CreateNewContactInformationPage;
import com.vtiger.ObjectRepository.CreateNewOrgInfoPage;
import com.vtiger.ObjectRepository.HomePage;
import com.vtiger.ObjectRepository.LoginPage;
import com.vtiger.ObjectRepository.SavebtnForContactInOrg;
import com.vtiger.ObjectRepository.SwitchToWindowForOrg;

import io.github.bonigarcia.wdm.WebDriverManager;
import io.qameta.allure.Description;
import io.qameta.allure.Epic;
import io.qameta.allure.Severity;
import io.qameta.allure.SeverityLevel;
import io.qameta.allure.Step;
import io.qameta.allure.Story;

public class CreateorganisationandcontactnameTest extends BaseClass{
	@Test(groups = "Sanity")
	@Description("Desription:- CreateContactsTest")
	@Epic("Epic:- CreateContactsTest")
	@Story("Story:- CreateContactsTest")
	@Step("Step:- CreateContactsTest")
	@Severity(SeverityLevel.MINOR)
	
	public void createorganisationandcontactnameTest()  throws EncryptedDocumentException, IOException, SQLException, InterruptedException {	
		String orgname = ExcelOffice.getDataFromExcel("organization", 2, 1)+JavaOffice.getRandomnumber(1000);  
		ClickCreatenewOrgName clickcreateneworgName = new  ClickCreatenewOrgName(driver);
		CreateNewOrgInfoPage orginfopage = new CreateNewOrgInfoPage(driver);
		ContactPage contactpage = new ContactPage(driver);
		CreateNewOrgInfoPage createnewOrginfoPage=new CreateNewOrgInfoPage(driver);
		SavebtnForContactInOrg savebutton = new SavebtnForContactInOrg(driver);
		if(driver.getTitle().contains("Home"))
		{
			ExcelOffice.setExcelData("contacts", 10, 7, "home page is displayed");
			ExcelOffice.setExcelData("contacts", 10, 8, "Pass");
		}
		homepage.homePageOrganizationAction();	
		clickcreateneworgName.orgClickAction();      
		createnewOrginfoPage.orgLastName(orgname);
		createnewOrginfoPage.saveBtnToOrg(driver);
		JavaOffice.printStatement("organisation created");	
			
		
		homepage.HomePageContactLink();
		contactpage.contactPageAction();	 
		savebutton.contactInformationPageAction("smith");
		savebutton.plusForOrgName();     
		if(driver.getTitle().contains("Contacts"))
		{
			ExcelOffice.setExcelData("contacts", 11, 7, "contact page is displayed");
			ExcelOffice.setExcelData("contacts", 11, 8, "Pass");
		}		
		if(driver.getTitle().contains("Contacts"))
		{
			ExcelOffice.setExcelData("contacts", 12, 7, "create contact page is displayed");
			ExcelOffice.setExcelData("contacts", 12, 8, "Pass");
		}
		if(driver.getTitle().contains("Contacts"))
		{
			ExcelOffice.setExcelData("contacts", 13, 7, " contact name is added");
			ExcelOffice.setExcelData("contacts", 13, 8, "Pass");  
		}
		WebDriverOffice.switchToWindow(driver, "Accounts");
		SwitchToWindowForOrg switchwin = new SwitchToWindowForOrg(driver);
		switchwin.searchText(orgname);
		switchwin.clickSearch();
		switchwin.actualOrg(driver,orgname);
		WebDriverOffice.switchToWindow(driver, "Contacts");
		savebutton.saveContact();
		JavaOffice.printStatement("contact created with org name");		  
		Reporter.log("contact created with org name");
	}
}


