package com.vtiger.practice;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;

import com.mysql.cj.jdbc.Driver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class newpract {
	
public static void main(String[] args) throws SQLException 
{
	
	 Connection connection = null;
	 try {
		 
	Driver driver=new Driver();
	DriverManager.registerDriver(driver);
	connection= DriverManager.getConnection("jdbc:mysql://localhost:3306/projects", "root", "root");
	Statement statement = connection.createStatement();
	statement.executeUpdate("insert into project values('TY_PROJ_143','smitha','28/4/2022','sdet433','created',24)");
	 }
	 finally {
		 connection.close(); 
 }WebDriver driver1 = null;
	 try {
		WebDriverManager.firefoxdriver().setup();
		 driver1= new FirefoxDriver();
		driver1.get("http://localhost:8084");
		driver1.manage().window().maximize();
		driver1.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver1.findElement(By.xpath("//input[@id='usernmae']")).sendKeys("rmgyantra");
		driver1.findElement(By.xpath("//input[@id='inputPassword']")).sendKeys("rmgy@9999");
		driver1.findElement(By.xpath("//button[text()='Sign in']")).click();
		driver1.findElement(By.xpath("//a[text()='Projects']")).click();
		List<WebElement> ele = driver1.findElements(By.xpath("//label[.='Project Status ']/following-sibling::select"));
        for(WebElement project:ele)
        {
        	if(project.getText().equalsIgnoreCase("sdet42"))
        	{
        		System.out.print("project is visible in GUI");
        		System.out.print("TC pass");
        		break;
        	}
        }
	}
	finally 
	{
    driver1.quit();
	
	}
	}
}

