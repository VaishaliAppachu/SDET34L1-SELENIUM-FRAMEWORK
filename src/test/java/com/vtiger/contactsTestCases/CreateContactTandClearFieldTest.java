package com.vtiger.contactsTestCases;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.Test;

import com.sdet34l1.genericUtility.BaseClass;
import com.sdet34l1.genericUtility.ExcelOffice;
import com.sdet34l1.genericUtility.FileOffice;
import com.sdet34l1.genericUtility.IconstantPathOffice;
import com.sdet34l1.genericUtility.JavaOffice;
import com.sdet34l1.genericUtility.WebDriverOffice;
import com.vtiger.ObjectRepository.HomePage;
import com.vtiger.ObjectRepository.LoginPage;
import com.vtiger.ObjectRepository.QuickCreateDropDown;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CreateContactTandClearFieldTest extends BaseClass{
	@Test
	public void createContactTC4() throws EncryptedDocumentException, IOException, InterruptedException
	{
		String contactname = ExcelOffice.getDataFromExcel("contacts", 2, 1)+JavaOffice.getRandomnumber(1000);
		if(driver.getTitle().contains("vtiger"))
		{
			ExcelOffice.setExcelData("contacts", 38, 7, "login page is displayed");
			ExcelOffice.setExcelData("contacts", 38, 8, "Pass");
		}
		JavaOffice.printStatement("login page is displayed");
		homepage.HomePageContactLink();
		if(driver.getTitle().contains("Administrator"))
		{
			ExcelOffice.setExcelData("contacts", 39, 7, "home page is displayed");
			ExcelOffice.setExcelData("contacts", 39, 8, "Pass");
		}
		JavaOffice.printStatement("home page is dispalyed");
		homepage.ClickOnQuickCreate();
		homepage.quickCreateDrop("New Contact");
		QuickCreateDropDown q = new QuickCreateDropDown(driver);
		q.clickOnQuickCreate(contactname);
		q.clearTheField();
		q.cancelBtn();
		if(driver.getTitle().contains("Administrator"))
		{
			ExcelOffice.setExcelData("contacts", 40, 7, "dropdown displayed");
			ExcelOffice.setExcelData("contacts", 40, 8, "Pass");
		}
		if(driver.getTitle().contains("Administrator"))
		{
			ExcelOffice.setExcelData("contacts", 41, 7, "new contact page displayed");
			ExcelOffice.setExcelData("contacts", 41, 8, "Pass");
		} 
		if(driver.getTitle().contains("Administrator"))
		{
			ExcelOffice.setExcelData("contacts", 42, 7, "all field cleared");
			ExcelOffice.setExcelData("contacts", 42, 8, "Pass");
		}
	}
}