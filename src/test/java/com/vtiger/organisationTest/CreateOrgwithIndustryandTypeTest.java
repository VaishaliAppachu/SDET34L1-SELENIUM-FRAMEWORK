package com.vtiger.organisationTest;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;
import java.util.Random;
import java.util.concurrent.TimeUnit;

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
import org.openqa.selenium.support.ui.Select;
import org.testng.Reporter;
import org.testng.annotations.Test;

import com.sdet34l1.genericUtility.BaseClass;
import com.sdet34l1.genericUtility.ExcelOffice;
import com.sdet34l1.genericUtility.FileOffice;
import com.sdet34l1.genericUtility.IconstantPathOffice;
import com.sdet34l1.genericUtility.JavaOffice;
import com.sdet34l1.genericUtility.WebDriverOffice;
import com.vtiger.ObjectRepository.ClickCreatenewOrgName;
import com.vtiger.ObjectRepository.CreateNewOrgInfoPage;
import com.vtiger.ObjectRepository.HomePage;
import com.vtiger.ObjectRepository.LoginPage;
import com.vtiger.ObjectRepository.OrgInformationValidateDropDown;

import io.github.bonigarcia.wdm.WebDriverManager;

public class CreateOrgwithIndustryandTypeTest extends BaseClass {
@Test(groups = "Regression")
	public void createOrgwithIndustryandType() throws IOException, InterruptedException {
	    String orgname = ExcelOffice.getDataFromExcel("organization", 2, 1)+JavaOffice.getRandomnumber(1000);
	    ClickCreatenewOrgName clickcreateneworgname = new ClickCreatenewOrgName(driver);
        CreateNewOrgInfoPage orginfopage = new CreateNewOrgInfoPage(driver);
        homepage.homePageOrganizationAction();
        clickcreateneworgname.orgClickAction();
		orginfopage.orgLastName(orgname);
		orginfopage.industryNameOfOrg("Education");
		orginfopage.accountTypeOfOrg("Press");
		orginfopage.saveBtnToOrg(driver);
        JavaOffice.printStatement("Organization created with indusrty and type");	
        Reporter.log("Organization created with indusrty and type");
	}


}


