package com.sdet34l1.genericUtility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import javax.imageio.stream.FileImageInputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * This class is used to maintain all excel specific common methods
 * @author VAISHALI
 *
 */
public class ExcelOffice {
	static Workbook wb;
/**
 * This method is used to get the data from the excel
 * @param sheetName
 * @param rowNumber
 * @param cellNumber
 * @return
 */
	public static String getDataFromExcel(String sheetName,int rowNumber,int cellNumber )
	{
		String data = wb.getSheet(sheetName).getRow(rowNumber).getCell(cellNumber).getStringCellValue();
		return data;
	   
	}
	/**
	 * This method is used to open the excel
	 * @param filepath
	 * @throws EncryptedDocumentException
	 * @throws IOException
	 */
	public static void openExcel(String filepath)
	{
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(filepath);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			wb=WorkbookFactory.create(fis);
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	/**
	 * This method is used to save the excel
	 * @param savepath
	 * @throws IOException
	 */
	public static void saveExcel(String savepath)
	{
		FileOutputStream fis = null;
		try {
			fis = new FileOutputStream(savepath);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			wb.write(fis);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	/**
	 * This method is used to close the excel
	 * @param filepath
	 * @throws IOException
	 */
	public static void closeExcel()
	{
		try {
			wb.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	/**
	 * This method is used to Set the excel
	 * @param sheetName
	 * @param rowNumber
	 * @param cellNumber
	 * @param text
	 */
	public static void setExcelData(String sheetName, int rowNumber, int cellNumber, String text)
	{
		wb.getSheet(sheetName).getRow(rowNumber).createCell(cellNumber).setCellValue(text);
	}
	/**
	 * This method is used to get multiple data from excel
	 */
	public static Object[][] getMultipleDataFromExcel(String sheetName)
	{
		Sheet sheet = wb.getSheet(sheetName);
		Object[][] arr=new Object[sheet.getLastRowNum()+1][sheet.getRow(0).getLastCellNum()];
		for(int i=0;i<=sheet.getLastRowNum();i++)
		{
			for(int j=0;j<sheet.getRow(i).getLastCellNum();j++)
			{
				arr[i][j]=sheet.getRow(i).getCell(j).getStringCellValue();
			}
		}
		return arr;
	}
}
