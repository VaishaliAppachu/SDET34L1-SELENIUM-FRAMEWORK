package com.rmgyantra.projectTest;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.List;
import java.util.concurrent.TimeUnit;

import javax.xml.crypto.Data;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

import com.mysql.cj.jdbc.Driver;
import com.sdet34l1.genericUtility.DatabaseOffice;
import com.sdet34l1.genericUtility.ExcelOffice;
import com.sdet34l1.genericUtility.FileOffice;
import com.sdet34l1.genericUtility.IconstantPathOffice;
import com.sdet34l1.genericUtility.JavaOffice;
import com.sdet34l1.genericUtility.WebDriverOffice;

import io.github.bonigarcia.wdm.WebDriverManager;

public class JdbcguiTest {

	public static void main(String[] args) throws SQLException  {
		FileOffice.openPropertyFile(IconstantPathOffice.RMGPROPERTYFILEPATH);
		ExcelOffice.openExcel(IconstantPathOffice.RMGEXCELPATH);
		int Randomnumber = JavaOffice.getRandomnumber(1000);
         String projectName = ExcelOffice.getDataFromExcel("Projects", 1, 1);
         System.out.println(projectName);
		 DatabaseOffice.openDBconnection(IconstantPathOffice.DATABASEPATHURL+FileOffice.getDataFromPropertyFile("dbName"), FileOffice.getDataFromPropertyFile("dbuser"),FileOffice.getDataFromPropertyFile("dbpassword"));
         DatabaseOffice.setDataInDataBase("insert into project values('TY_RMG_"+Randomnumber+"','chan','28/4/2022','"+Randomnumber+"','on going',3)");
         DatabaseOffice.closeDatabase();
        
		   WebDriverManager.chromedriver().setup();
			WebDriver driver1= new ChromeDriver();
             WebDriverOffice.browserSettings(10, driver1);
            WebDriverOffice.navigateApp("http://localhost:8084", driver1);
			driver1.findElement(By.xpath("//input[@id='usernmae']")).sendKeys("rmgyantra");
			driver1.findElement(By.xpath("//input[@id='inputPassword']")).sendKeys("rmgy@9999");
			driver1.findElement(By.xpath("//button[text()='Sign in']")).click();
			driver1.findElement(By.xpath("//a[text()='Projects']")).click();
			//driver1.findElement(By.xpath("//span[text()='Create Project']")).click();
			
			 List<WebElement> ele = driver1.findElements(By.xpath("//td[2]"));
//			 
//			 for(int i=ele.size()-1;i>=0;i--)
//			 {
//				WebElement n=ele.get(i);
//				t=n.getText();
//				if(t.contentEquals("sdetsdet11"))
//				{
//					break;
//					
//				}
//			 }
			 for(WebElement project:ele)
			 {
				 if(project.getText().equalsIgnoreCase(projectName))
				 {
					 System.out.println("proj name is presesnt in gui");
					 System.out.println("TC pass");
					 break;
					 
				 }
			 }
//			 if(t.contentEquals("sdetsdet11"))
//			 {
//					System.out.println("proj presesnt in gui");
//
//			 }
//			 else
//			 {
//					System.out.println("proj not presesnt in gui");
//
//			 }
			 WebDriverOffice.quitBrowser(driver1);
	}

}
