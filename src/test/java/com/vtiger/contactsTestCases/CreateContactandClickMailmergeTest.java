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
import com.vtiger.ObjectRepository.HomePage;
import com.vtiger.ObjectRepository.LoginPage;
import com.vtiger.ObjectRepository.MailMergeTemplate;
import com.vtiger.ObjectRepository.SettingsMouOverAction;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CreateContactandClickMailmergeTest extends BaseClass{
	@Test
	public void createContactTC5() throws EncryptedDocumentException, IOException, InterruptedException {
		String contactname = ExcelOffice.getDataFromExcel("contacts", 2, 1) + JavaOffice.getRandomnumber(1000);	
		if (driver.getTitle().contains("vtiger"))
		{
			ExcelOffice.setExcelData("contacts", 45, 7, "login page is displayed");
			ExcelOffice.setExcelData("contacts", 45, 8, "Pass");
		}
		JavaOffice.printStatement("login page is dispalyed");
		if (driver.getTitle().contains("Home")) {
			ExcelOffice.setExcelData("contacts", 46, 7, "home page is displayed");
			ExcelOffice.setExcelData("contacts", 46, 8, "Pass");
		}
		JavaOffice.printStatement("home page is displayed");
		
		SettingsMouOverAction setting = new SettingsMouOverAction(driver);
		setting.mouOverOnSettings(driver);

		if (driver.getTitle().contains("Administrator")) {
			ExcelOffice.setExcelData("contacts", 47, 7, "Settings page is displayed");
			ExcelOffice.setExcelData("contacts", 47, 8, "Pass");
			Thread.sleep(1000);
			MailMergeTemplate mail = new MailMergeTemplate(driver);
			mail.clickOnMailMerge(driver);
			JavaOffice.printStatement("Mail Merge Template is displayed");
			if(driver.getTitle().contains("Settings")) 
			{
				ExcelOffice.setExcelData("contacts", 48, 7,"Settings mailmerge page is displayed"); 
				ExcelOffice.setExcelData("contacts", 48, 8, "Pass");
			}
			mail.clickAddTemplates();
		}

	}
}