package com.sdet34l1.genericUtility;

import java.sql.Date;
import java.text.SimpleDateFormat;
import java.util.Random;

import org.openqa.selenium.WebElement;
/**
 * 
 * @author VAISHALI
 *
 */
public class JavaOffice 
{
	/**
	 * This method is used to convert from String value to long data type
	 * @param value
	 * @return
	 */
public static long stringToLong(String value)
{
	return Long.parseLong(value);
}

/**
 * This method is used to get the Random Number
 * @param limit
 * @return
 */
public static int getRandomnumber(int limit)
{
	Random ran = new Random();
	return ran.nextInt(limit);
}
/**
 * This method is used to print the statement
 * @param message
 */
public static void printStatement(String message)
{
	System.out.println(message);
}
/**
 * This method is used to wait until element clickable
 * @param element
 * @param polingTime
 * @param duration
 * @throws InterruptedException
 */
public static void customWait(WebElement element,long polingTime,int duration) 
{
	int count=0;
	while(count<=duration) {
		try {
		element.click();
		break;
	        }
		catch(Exception e) {
			try {
				Thread.sleep(polingTime);
			} catch (InterruptedException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			count++;
		}
}
}
/**
 * This method is used for the conditions
 * @param actualresult
 * @param expectedresult
 * @param Testcasename
 */
public static void assertionthroughifCondition(String actualresult,String expectedresult,String Testcasename)
{
    if(actualresult.equalsIgnoreCase(expectedresult))
    		{
    	    System.out.println(Testcasename+"created succesfully");
    	    System.out.println("Pass");

    		}
/**
 * This method is used to date and Time Format
 */
}
public String dateTimeInFormat()
{
	return (new SimpleDateFormat("dd-mm-yyy-HH-ss").format(new java.util.Date()));
}
}