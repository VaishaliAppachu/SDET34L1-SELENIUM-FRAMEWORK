package com.vtiger.documents;

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
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.testng.Reporter;
import org.testng.annotations.Test;

import com.sdet34l1.genericUtility.BaseClass;
import com.sdet34l1.genericUtility.ExcelOffice;
import com.sdet34l1.genericUtility.FileOffice;
import com.sdet34l1.genericUtility.IconstantPathOffice;
import com.sdet34l1.genericUtility.JavaOffice;
import com.sdet34l1.genericUtility.WebDriverOffice;
import com.vtiger.ObjectRepository.CreateNewDocumentPage;
import com.vtiger.ObjectRepository.DocumentPage;
import com.vtiger.ObjectRepository.HomePage;
import com.vtiger.ObjectRepository.LoginPage;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CreatedocumentsTest extends BaseClass{
	@Test(groups = {"Regression","Sanity"})
	public void createdocumentsFrameTest() throws EncryptedDocumentException, IOException, InterruptedException {
		String documents = ExcelOffice.getDataFromExcel("Documents", 2, 1)+JavaOffice.getRandomnumber(1000);
		String documents1 = ExcelOffice.getDataFromExcel("Documents", 2, 2);
		homepage.homePageDocumentLink();
		DocumentPage docu = new DocumentPage(driver);
		docu.documentPageAction();
		CreateNewDocumentPage newDocPage=new CreateNewDocumentPage(driver);
		newDocPage.documentPageSendTitle(documents);
		WebDriverOffice.switchToFrame(driver, 0);
		newDocPage.documentPageSendTextAndCopyAll(documents1);
		WebDriverOffice.switchBackToHomepage(driver);
		newDocPage.documentPageBoldAndItalicText();
		newDocPage.documentPageFileUpload("C:\\Users\\VAISHALI\\Desktop\\LOC.png");
		JavaOffice.printStatement(documents1);
		Reporter.log("all done");
	}
}
