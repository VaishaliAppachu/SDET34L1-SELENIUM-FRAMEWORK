package com.sdet34l1.genericUtility;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.Reporter;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Parameters;

import com.vtiger.ObjectRepository.HomePage;
import com.vtiger.ObjectRepository.LoginPage;

import io.github.bonigarcia.wdm.WebDriverManager;
/**
 * This class contains implementation of all basic configuration annotations
 * @author VAISHALI
 * This base class includes all before and after method
 */
public class BaseClass {
	String username;
	String password;
	public long longtimeout;
	public WebDriver driver;
	public LoginPage loginpage;
	public HomePage homepage;
	public String filepath;
	public static WebDriver staticdriver;

/**
 * In this annotation we open the database,open the property file,open the excel sheet
 */
	@BeforeSuite(groups = "baseclass")
	public void beforesuiteTest()
	{
		FileOffice.openPropertyFile(IconstantPathOffice.PROPERTYFILEPATH);
		ExcelOffice.openExcel(IconstantPathOffice.EXCELPATH);
	}
	/**
	 * In this annotation we fetch the common data from the property file,
	 * create the instance for all generic utility
	 * launch browser, do browser settings 
	 * create the instance for common object repository 
	 */
	//@Parameters("browser")
	@BeforeClass(groups = "baseclass")
	//String browser
	public void beforeclassTest()
	{
		String url=FileOffice.getDataFromPropertyFile("url");
		String browser=FileOffice.getDataFromPropertyFile("browser");
		username= FileOffice.getDataFromPropertyFile("username");
		password= FileOffice.getDataFromPropertyFile("password");
		longtimeout =JavaOffice.stringToLong( FileOffice.getDataFromPropertyFile("timeout"));
		switch(browser){
		case "chrome":
			WebDriverManager.chromedriver().setup();
			driver=new ChromeDriver();
			break;
		case "firefox":
			WebDriverManager.firefoxdriver().setup();
			driver=new FirefoxDriver();
			break;
		default:
			System.out.println("Please specify proper browser key");
			WebDriverManager.chromedriver().setup();
			driver=new ChromeDriver();
			break;
		}
		staticdriver=driver;
		WebDriverOffice.navigateApp(url, driver);
		WebDriverOffice.browserSettings(longtimeout, driver);
		WebDriverOffice.waitTillPageLoad(longtimeout, driver);
		loginpage= new LoginPage(driver);
		homepage= new HomePage(driver);
	}
/**
 * In this annotation we do login action
 */

	@BeforeMethod(groups = "baseclass")
	public void beforeMethodTest()
	{
		loginpage.loginAction(username, password);			
	}
	/**
	 * In this annotation we do logout action
	 */
	@AfterMethod(groups = "baseclass")
	public void afterMethodTest()
	{
		homepage.signOutIconMouseHoverAction(driver);
	}	
	@AfterClass(groups = "baseclass")
	/**
	 * In this annotation we do quit action
	 */
	public void afterClassTest()
	{
		WebDriverOffice.quitBrowser(driver);
	}
	
	@AfterSuite(groups = "baseclass")
	/**
	 * In this annotation we close the excel
	 */
	public void afterSuiteTest()
	{
	    ExcelOffice.saveExcel(IconstantPathOffice.EXCELPATH);	
	    ExcelOffice.closeExcel();
		
	}
}

