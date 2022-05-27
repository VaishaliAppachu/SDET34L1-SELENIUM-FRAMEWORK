package com.sdet34l1.genericUtility;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

/**
 * 
 * @author VAISHALI
 *
 */
public class FileOffice {
	static Properties property;
/**
 * This method is used to open the property file
 * @param filepath
 * @throws IOException
 */
	public static void openPropertyFile(String filepath){
		 FileInputStream fis1 = null;
		try {
			fis1 = new FileInputStream(filepath);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		    property = new Properties();
		    try {
		    	
				property.load(fis1);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		
	}
	/**
	 * This method is used to fetch Data From Property File
	 * @param key
	 * @return 
	 */
	public static String getDataFromPropertyFile(String key) {
		String value = property.getProperty(key);
		return value;
	}
	
}
