package com.vtiger.campaignTest;

import java.io.IOException;

import org.testng.Reporter;
import org.testng.annotations.Test;

import com.sdet34l1.genericUtility.BaseClass;
import com.sdet34l1.genericUtility.ExcelOffice;
import com.sdet34l1.genericUtility.JavaOffice;
import com.vtiger.ObjectRepository.ClickCreateNewCampaign;
import com.vtiger.ObjectRepository.ClickOnCampaign;
import com.vtiger.ObjectRepository.CreateNewCampaignInfo;

public class CreatecampaignTest extends BaseClass{

	@Test(groups = "Regression")

	public void createcampaignTest() throws IOException {
		String campname = ExcelOffice.getDataFromExcel("campaign", 2, 1)+JavaOffice.getRandomnumber(1000);
		ClickOnCampaign clickcampaign = new ClickOnCampaign(driver);
		clickcampaign.clickCampainLink(driver);
		ClickCreateNewCampaign clickcreatecampaign = new ClickCreateNewCampaign(driver);
		clickcreatecampaign.ClickonPlus();
		CreateNewCampaignInfo createcampaigninfo = new CreateNewCampaignInfo(driver);
		createcampaigninfo.CreateCampaignlastName(campname);
		createcampaigninfo.campaingSave();
		JavaOffice.printStatement("campaign name has been created");
		Reporter.log("campaign name has been created");
	}
}
