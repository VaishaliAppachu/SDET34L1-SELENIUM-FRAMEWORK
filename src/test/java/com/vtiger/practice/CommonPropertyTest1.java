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

import io.github.bonigarcia.wdm.WebDriverManager;

public class CommonPropertyTest1 {

	public static void main(String[] args) throws IOException   {
    FileInputStream fis = new FileInputStream("./src/test/resources/commonData.properties");
    Properties property = new Properties();
    property.load(fis);
    String url=property.getProperty("url");
    String browser=property.getProperty("browser");
    String username= property.getProperty("username");
    String password= property.getProperty("password");
    String timeout= property.getProperty("timeout");
    
    System.out.println(url);
    System.out.println(browser);
    System.out.println(username);
    System.out.println(password);
    System.out.println(timeout);
    long longtimeout = Long.parseLong(timeout);
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
		driver1.findElement(By.linkText("Contacts")).click();
		driver1.findElement(By.xpath("//img[@title='Create Contact...']")).click();
		
		String firstName="sdet3";
		String LastName="L6";
		driver1.findElement(By.name("firstname")).sendKeys(firstName);		
		driver1.findElement(By.name("lastname")).sendKeys(LastName);
		driver1.findElement(By.xpath("//input[@accesskey='S']")).click();
		
		WebElement actualfn = driver1.findElement(By.id("dtlview_First Name"));
		WebElement actualln = driver1.findElement(By.id("dtlview_Last Name"));
    
    if(actualfn.getText().equalsIgnoreCase(firstName)&&actualln.getText().equalsIgnoreCase(LastName))
    {
  	  System.out.println("contact creates successfully");
    }
    
    WebElement t = driver1.findElement(By.xpath("//img[@src='themes/softed/images/user.PNG']"));
		Actions act=new Actions(driver1);
		act.moveToElement(t).perform();
		driver1.findElement(By.xpath("//a[.='Sign Out']")).click();
		driver1.close();
}
}


