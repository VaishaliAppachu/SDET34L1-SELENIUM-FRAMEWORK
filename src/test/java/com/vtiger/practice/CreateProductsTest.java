package com.vtiger.practice;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;

import com.sdet34l1.genericUtility.FileOffice;
import com.sdet34l1.genericUtility.IconstantPathOffice;
import com.sdet34l1.genericUtility.JavaOffice;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CreateProductsTest {

	public static void main(String[] args) throws IOException {
	        
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
	  driver1.manage().timeouts().implicitlyWait(longtimeout, TimeUnit.SECONDS);
	  
	    driver1.findElement(By.xpath("//input[@type='text']")).sendKeys("admin");
		driver1.findElement(By.xpath("//input[@type='password']")).sendKeys("admin");
		driver1.findElement(By.xpath("//input[@id='submitButton']")).click();
		driver1.findElement(By.linkText("Products")).click();
		driver1.findElement(By.xpath("//img[@alt='Create Product...']")).click();
		driver1.findElement(By.name("productname")).sendKeys("ssvv");
        driver1.findElement(By.xpath("//input[@accesskey='S']")).click();
		
		System.out.println("product has been created");
		
		 WebElement t = driver1.findElement(By.xpath("//img[@src='themes/softed/images/user.PNG']"));
			Actions act1=new Actions(driver1);
			act1.moveToElement(t).perform();
			driver1.findElement(By.xpath("//a[.='Sign Out']")).click();
			driver1.close();
	}
	}



