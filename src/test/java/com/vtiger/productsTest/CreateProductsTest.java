package com.vtiger.productsTest;

import static org.testng.Assert.fail;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.SQLException;
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
import org.testng.Assert;
import org.testng.Reporter;
import org.testng.annotations.Test;

import com.sdet34l1.genericUtility.BaseClass;
import com.sdet34l1.genericUtility.ExcelOffice;
import com.sdet34l1.genericUtility.FileOffice;
import com.sdet34l1.genericUtility.IconstantPathOffice;
import com.sdet34l1.genericUtility.JavaOffice;
import com.sdet34l1.genericUtility.WebDriverOffice;
import com.vtiger.ObjectRepository.ContactPage;
import com.vtiger.ObjectRepository.CreateNewProductPage;
import com.vtiger.ObjectRepository.CreateProductPage;
import com.vtiger.ObjectRepository.HomePage;
import com.vtiger.ObjectRepository.LoginPage;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CreateProductsTest extends BaseClass{
	@Test(groups = "Sanity")
	public void createProductsTest()  throws EncryptedDocumentException, IOException, SQLException {
		String products = ExcelOffice.getDataFromExcel("products", 2, 1);
		homepage.homePageForProduct();
		CreateProductPage createproductpage = new CreateProductPage(driver);
		createproductpage.clickForProduct();
		CreateNewProductPage createnewproductpage = new CreateNewProductPage(driver);
		createnewproductpage.productLastName(products);
		JavaOffice.printStatement("product created");		  
		Reporter.log("product created");
	}
}