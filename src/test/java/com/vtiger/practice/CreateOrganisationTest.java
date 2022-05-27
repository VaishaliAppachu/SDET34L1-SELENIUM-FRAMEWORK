package com.vtiger.practice;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CreateOrganisationTest {
		public static void main(String[] args) {
			WebDriverManager.firefoxdriver().setup();
			WebDriver driver=new FirefoxDriver();
			driver.get("http://localhost:8888");
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			driver.findElement(By.xpath("//input[@type='text']")).sendKeys("admin");
			driver.findElement(By.xpath("//input[@type='password']")).sendKeys("admin");
			driver.findElement(By.xpath("//input[@id='submitButton']")).click();
			driver.findElement(By.linkText("Organizations")).click();
			driver.findElement(By.xpath("//img[@alt='Create Organization...']")).click();
		    driver.findElement(By.xpath("//input[@name='accountname']")).sendKeys("abc17");
			driver.findElement(By.xpath("//input[@accesskey='S']")).click();
			String z = driver.findElement(By.id("dtlview_Organization Name")).getText();
			if(z.contains("abc17"))
			{
				System.out.print("organisation added");
			}
			else
			{
				System.out.print("organisation not added");
			}
			WebElement t = driver.findElement(By.xpath("//img[@src='themes/softed/images/user.PNG']"));
			Actions act=new Actions(driver);
			act.moveToElement(t).perform();
			driver.findElement(By.xpath("//a[.='Sign Out']")).click();
			driver.close();
		}

	}
   


