package com.vtiger.practice;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

import com.mysql.cj.jdbc.Driver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class VtigerTest2 {
	public static void main(String[] args) throws SQLException 
	{
		  WebDriverManager.firefoxdriver().setup();
			WebDriver driver1=new FirefoxDriver();
			driver1.get("http://localhost:8888");
			driver1.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			driver1.findElement(By.xpath("//input[@type='text']")).sendKeys("admin");
			driver1.findElement(By.xpath("//input[@type='password']")).sendKeys("admin");
			driver1.findElement(By.xpath("//input[@id='submitButton']")).click();
    Driver driver = new Driver();
    DriverManager.registerDriver(driver);
    Connection connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/vtiger", "root", "root");
    Statement statement = connection.createStatement();
    ResultSet result = statement.executeQuery("select * from vtigertab;");
    while(result.next())
    {
    	String res = result.getString(1);
  
	driver1.findElement(By.linkText("Organizations")).click();
	driver1.findElement(By.xpath("//img[@alt='Create Organization...']")).click();
    driver1.findElement(By.xpath("//input[@name='accountname']")).sendKeys(res);
    driver1.findElement(By.xpath("//input[@accesskey='S']")).click();
    String z = driver1.findElement(By.id("dtlview_Organization Name")).getText();
    if(z.equals(res))
    {
    	System.out.println("Organization added");
    	System.out.println("TC passed");
    }
	}
    connection.close();
	}
}


