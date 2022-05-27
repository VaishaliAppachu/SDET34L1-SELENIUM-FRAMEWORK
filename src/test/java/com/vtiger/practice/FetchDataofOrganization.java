package com.vtiger.practice;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
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

import com.mysql.cj.jdbc.Driver;
import com.sdet34l1.genericUtility.FileOffice;
import com.sdet34l1.genericUtility.IconstantPathOffice;
import com.sdet34l1.genericUtility.JavaOffice;

import io.github.bonigarcia.wdm.WebDriverManager;

public class FetchDataofOrganization {

	public static void main(String[] args)  throws EncryptedDocumentException, IOException, SQLException {
		FileInputStream fis = new FileInputStream("./src/test/resources/TestData.xlsx");
	    Workbook wb = WorkbookFactory.create(fis);
	    Sheet sh = wb.getSheet("organization");
	    Row row = sh.getRow(2);
	    Cell cell = row.getCell(1);
	    Random ran = new Random();
	    int random = ran.nextInt(1000);
	    String data = cell.getStringCellValue()+random;
	    System.out.println(data);
	   
		FileOffice.openPropertyFile(IconstantPathOffice.PROPERTYFILEPATH);

		   
	    String url=FileOffice.getDataFromPropertyFile("url");
	    String browser=FileOffice.getDataFromPropertyFile("browser");
	    String username= FileOffice.getDataFromPropertyFile("username");
	    String password= FileOffice.getDataFromPropertyFile("password");
	    String timeout= FileOffice.getDataFromPropertyFile("timeout");
	    
	    long longtimeout =JavaOffice.stringToLong(timeout);
        
	     int Randomnumber = JavaOffice.getRandomnumber(1000);
	     
	    WebDriver driver1=null;
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
	      driver1.get(url);
	      driver1.manage().window().maximize();
		  driver1.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		  driver1.findElement(By.xpath("//input[@type='text']")).sendKeys("admin");
		  driver1.findElement(By.xpath("//input[@type='password']")).sendKeys("admin");
		  driver1.findElement(By.xpath("//input[@id='submitButton']")).click();
		  driver1.findElement(By.linkText("Organizations")).click();
			driver1.findElement(By.xpath("//img[@alt='Create Organization...']")).click();
		    driver1.findElement(By.xpath("//input[@name='accountname']")).sendKeys(data);
		    driver1.findElement(By.xpath("//input[@accesskey='S']")).click();
		    System.out.println("organisation created");


		  
	     }
}


