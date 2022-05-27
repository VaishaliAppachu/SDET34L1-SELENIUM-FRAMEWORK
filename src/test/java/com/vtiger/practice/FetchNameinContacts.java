package com.vtiger.practice;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Properties;
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
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.sdet34l1.genericUtility.FileOffice;
import com.sdet34l1.genericUtility.IconstantPathOffice;
import com.sdet34l1.genericUtility.JavaOffice;

import io.github.bonigarcia.wdm.WebDriverManager;

public class FetchNameinContacts {

	public static void main(String[] args)  throws EncryptedDocumentException, IOException, SQLException, InterruptedException {
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
		  driver1.manage().timeouts().implicitlyWait(longtimeout, TimeUnit.SECONDS);
		  driver1.findElement(By.xpath("//input[@type='text']")).sendKeys(username);
		  driver1.findElement(By.xpath("//input[@type='password']")).sendKeys(password);
		  driver1.findElement(By.xpath("//input[@id='submitButton']")).click();
		  driver1.findElement(By.linkText("Organizations")).click();
		  driver1.findElement(By.xpath("//img[@alt='Create Organization...']")).click();
		  driver1.findElement(By.xpath("//input[@name='accountname']")).sendKeys(data);
		  driver1.findElement(By.xpath("//input[@accesskey='S']")).click();
		  System.out.println("organisation created");
		  String data1 = driver1.findElement(By.id("dtlview_Organization Name")).getText();
		  System.out.println(data1);
		  //contacts
		  //Thread.sleep(1000);
		  driver1.findElement(By.xpath("//a[.='Contacts']")).click();
		  //driver1.findElement(By.linkText("Contacts")).click();
		  
		  
		 //WebElement ele = driver1.findElement(By.xpath("//a[.='Contacts']"));
		 //Actions act =new Actions(driver1);
		 //act.moveToElement(ele).perform();
		 //ele.click();
		  driver1.findElement(By.xpath("//img[@title='Create Contact...']")).click();
		  driver1.findElement(By.xpath("//input[@name='lastname']")).sendKeys("rupi");
		  driver1.findElement(By.xpath("(//img[@title='Select'])[1]")).click();
		  
		  //switch to another tab
		  Set<String> drv = driver1.getWindowHandles();
		  ArrayList<String> child = new ArrayList<String>(drv);
		  String newchild = child.get(1);
		  
		  driver1.switchTo().window(newchild);
		  
		  driver1.manage().window().maximize();
		  
		  driver1.findElement(By.name("search_text")).sendKeys(data);
		  
	      driver1.findElement(By.name("search")).click();
	      Thread.sleep(2000);
	      driver1.findElement(By.xpath("//tr[2]/td/a")).click();
	      
	     String newparent = child.get(0);
	     driver1.switchTo().window(newparent);
	      Thread.sleep(2000);
		 driver1.findElement(By.xpath("//input[@accesskey='S']")).click();
		 
		  System.out.println("contact created");
		  
		  
	}

}
