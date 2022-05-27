package com.vtiger.campaignTest;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.ArrayList;
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
import org.openqa.selenium.interactions.Actions;
import org.testng.Reporter;
import org.testng.annotations.Test;

import com.sdet34l1.genericUtility.BaseClass;
import com.sdet34l1.genericUtility.ExcelOffice;
import com.sdet34l1.genericUtility.FileOffice;
import com.sdet34l1.genericUtility.IconstantPathOffice;
import com.sdet34l1.genericUtility.JavaOffice;
import com.sdet34l1.genericUtility.WebDriverOffice;
import com.vtiger.ObjectRepository.ClickCreateNewCampaign;
import com.vtiger.ObjectRepository.ClickOnCampaign;
import com.vtiger.ObjectRepository.CreateNewCampaignInfo;
import com.vtiger.ObjectRepository.CreateNewProductPage;
import com.vtiger.ObjectRepository.CreateProductPage;
import com.vtiger.ObjectRepository.HomePage;
import com.vtiger.ObjectRepository.LoginPage;
import com.vtiger.ObjectRepository.SearchInsideForProduct;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CreateCampaignWithProductnameTest extends BaseClass{

	@Test(groups = "Sanity")
	public void createCampaignWithProductname()  throws EncryptedDocumentException, IOException, SQLException, InterruptedException {
		String campaignname = ExcelOffice.getDataFromExcel("campaign", 2, 1)+JavaOffice.getRandomnumber(1000);
		String productname = ExcelOffice.getDataFromExcel("products", 2, 1)+JavaOffice.getRandomnumber(1000);
		homepage.homePageForProduct();
		CreateProductPage createproductpage = new CreateProductPage(driver);
		createproductpage.clickForProduct();
		CreateNewProductPage createnewproductpage = new CreateNewProductPage(driver);
		createnewproductpage.productLastName(productname);
		JavaOffice.printStatement("product name has been created");	
		WebDriverOffice.explicitlyWait(longtimeout, driver);
		ClickOnCampaign clickoncampaign = new ClickOnCampaign(driver);
		clickoncampaign.clickCampainLink(driver);
		ClickCreateNewCampaign clickcreatecampaign = new ClickCreateNewCampaign(driver);
		clickcreatecampaign.ClickonPlus();
		CreateNewCampaignInfo createcampaigninfo = new CreateNewCampaignInfo(driver);
		createcampaigninfo.CreateCampaignlastName(campaignname);
		createcampaigninfo.clickForProductSearch();
		WebDriverOffice.switchToWindow(driver, "Products");			               
		WebDriverOffice.maximizeBrowser(driver);	
		WebDriverOffice.explicitlyWait(longtimeout, driver);
		SearchInsideForProduct search = new SearchInsideForProduct(driver);
		search.searchTheCreatedProduct(productname);
		search.getTheCreatedProduct();
		WebDriverOffice.switchToWindow(driver, "Campaigns");
		JavaOffice.printStatement("campaign and product name has been created");
		Reporter.log("campaign and product name has been created");
	}

}
