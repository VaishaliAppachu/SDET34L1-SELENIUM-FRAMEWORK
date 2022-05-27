package com.vtiger.organisationTest;

import static org.testng.Assert.fail;

import java.io.IOException;

import org.testng.Assert;
import org.testng.Reporter;
import org.testng.annotations.Test;

import com.sdet34l1.genericUtility.BaseClass;
import com.sdet34l1.genericUtility.ExcelOffice;
import com.sdet34l1.genericUtility.JavaOffice;
import com.sdet34l1.genericUtility.WebDriverOffice;
import com.vtiger.ObjectRepository.ClickCreatenewOrgName;
import com.vtiger.ObjectRepository.CreateNewOrgInfoPage;

public class CreateOrganisationTest extends BaseClass {
	@Test(groups = "Sanity")
	public void CreateOrganisationTest1() throws IOException, InterruptedException {

		String orgname = ExcelOffice.getDataFromExcel("organization", 2, 1)+JavaOffice.getRandomnumber(1000);
		homepage.homePageOrganizationAction();
		ClickCreatenewOrgName clickcreateneworgname = new ClickCreatenewOrgName(driver);
		clickcreateneworgname.orgClickAction();
		CreateNewOrgInfoPage orginfopage = new CreateNewOrgInfoPage(driver);
		orginfopage.orgLastName(orgname);
		orginfopage.saveBtnToOrg(driver);
		WebDriverOffice.explicitlyWait(longtimeout, driver);
		Reporter.log("Organization is created");
		
	}
}