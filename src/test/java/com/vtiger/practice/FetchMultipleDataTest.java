package com.vtiger.practice;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class FetchMultipleDataTest {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
	FileInputStream excf = new FileInputStream("./src/test/resources/TestData.xlsx");
	Workbook wb = WorkbookFactory.create(excf);
	Sheet sh = wb.getSheet("multiple");
	for(int i=0; i<=sh.getLastRowNum(); i++)
	{
	  for(int j=0; j<sh.getRow(i).getLastCellNum(); j++)
	  {
		  System.out.println(sh.getRow(i).getCell(j).getStringCellValue());
	  }
	}
	}

}
