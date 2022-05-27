package com.vtiger.practice;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.testng.Reporter;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import com.sdet34l1.genericUtility.ExcelOffice;
import com.sdet34l1.genericUtility.IconstantPathOffice;

public class TestNgMultipleData extends TestNgBasicConfAnnotationTest{
	@Test(dataProvider = "multiple")
	
	public void testNgMultipleData(String un,String pwd) throws  IOException, InterruptedException 
	{
		Reporter.log(un+"----->   "+pwd,true);
	}
@DataProvider
public Object[][] multiple() throws EncryptedDocumentException, IOException
{
	ExcelOffice.openExcel(IconstantPathOffice.EXCELPATH);
	return ExcelOffice.getMultipleDataFromExcel("multiple");
}
}
