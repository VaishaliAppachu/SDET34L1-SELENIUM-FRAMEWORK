package com.vtiger.practice;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

import com.sdet34l1.genericUtility.ExcelOffice;
import com.sdet34l1.genericUtility.FileOffice;
import com.sdet34l1.genericUtility.IconstantPathOffice;
import com.sdet34l1.genericUtility.JavaOffice;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Demoautomation {

	public static void main(String[] args) {
		WebDriver driver1=null;
		FileOffice.openPropertyFile(IconstantPathOffice.PROPERTYFILEPATH);
		ExcelOffice.openExcel(IconstantPathOffice.EXCELPATH);
		
		int Randomnumber = JavaOffice.getRandomnumber(1000);
		
		String fullname = ExcelOffice.getDataFromExcel("Fullname", 1, 1);
		   String email = ExcelOffice.getDataFromExcel("email", 2, 1);
		String caddress = ExcelOffice.getDataFromExcel("caddress", 3, 1);
		String paddress = ExcelOffice.getDataFromExcel("paddress", 4, 1);
		
		String url=FileOffice.getDataFromPropertyFile("url");
		String browser=FileOffice.getDataFromPropertyFile("browser");
		String username= FileOffice.getDataFromPropertyFile("username");
		String password= FileOffice.getDataFromPropertyFile("password");
		String timeout= FileOffice.getDataFromPropertyFile("timeout");
             
		long longtimeout = JavaOffice.stringToLong(timeout);
		
		switch(browser)
		{
		case "chrome":
			WebDriverManager.chromedriver().setup();
			driver1=new ChromeDriver();
			break;
		case "firefox":
			WebDriverManager.firefoxdriver().setup();
			driver1=new FirefoxDriver();
			break;
		default:
			System.out.println("Please specify proper browser key");
			WebDriverManager.chromedriver().setup();
			driver1=new ChromeDriver();
			break;
		}

	}
	


}
