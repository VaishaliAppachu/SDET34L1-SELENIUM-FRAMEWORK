package com.vtiger.practice;

import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import com.sdet34l1.genericUtility.WebDriverOffice;

import io.github.bonigarcia.wdm.WebDriverManager;

public class JavaScriptExecutor {

	public static void main(String[] args) throws IOException {
     WebDriverManager.chromedriver().setup();
     WebDriver driver=new ChromeDriver();
     WebDriverOffice.initializeJscript(driver);  
     driver.manage().window().maximize();
     WebDriverOffice.navigateAppthroughJscript("http://localhost:8888/");     
     WebDriverOffice.enterDataThroughJscript(driver.findElement(By.name("user_name")),"admin");
     WebDriverOffice.enterDataThroughJscript(driver.findElement(By.name("user_password")),"admin");
     WebDriverOffice.clickThroghJscript(driver.findElement(By.id("submitButton")));
     WebDriverOffice.scrollTitleElementJscript(driver.findElement(By.id("stuff_7")));
     String fileName = new JavaScriptExecutor().getClass().getName();
     WebDriverOffice.takesScreenshot(fileName, driver);
	}

}
