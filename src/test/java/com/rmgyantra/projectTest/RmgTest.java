package com.rmgyantra.projectTest;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;

import com.mysql.cj.jdbc.Driver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class RmgTest {

	
		public static void main(String[] args) throws SQLException 
		{
			String var=null;
			Driver driver = new Driver();
		    DriverManager.registerDriver(driver);
			Connection connt = DriverManager.getConnection("jdbc:mysql://localhost:3306/projects", "root", "root");
			Statement stmnt = connt.createStatement();
			//int result = stmnt.executeUpdate("insert into project values('TY_PROJ_123','ROOPAtm','27/04/2022','DATABASE12','On Going','0')");

			WebDriverManager.firefoxdriver().setup();
			WebDriver driver1= new FirefoxDriver();
			driver1.get("http://localhost:8084");
			driver1.manage().window().maximize();
			driver1.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			driver1.findElement(By.xpath("//input[@id='usernmae']")).sendKeys("rmgyantra");
			driver1.findElement(By.xpath("//input[@id='inputPassword']")).sendKeys("rmgy@9999");
			driver1.findElement(By.xpath("//button[text()='Sign in']")).click();

			driver1.findElement(By.xpath("//a[text()='Projects']")).click();
			driver1.findElement(By.xpath("//span[text()='Create Project']")).click();
			driver1.findElement(By.xpath("//input[@name='projectName']")).sendKeys("sdetsdet11");
			driver1.findElement(By.xpath("//input[@name='createdBy']")).sendKeys("mohon");
			WebElement ele = driver1.findElement(By.xpath("(//select[@name='status'])[2]"));
			Select s = new Select(ele);
			s.selectByIndex(2);
			driver1.findElement(By.xpath("//input[@value='Add Project']")).click();

			 String qyery = "select * from project";
		     ResultSet res = stmnt.executeQuery(qyery);
			
			while(res.next())
			{
					var = res.getString("Project_name");
								
				
			}
			
			if(var.equals("sdetsdet11"))
					{
				System.out.println("proj presesnt");
					}
			else
			{
				System.out.println("proj not presesnt");
			}
	}

}
