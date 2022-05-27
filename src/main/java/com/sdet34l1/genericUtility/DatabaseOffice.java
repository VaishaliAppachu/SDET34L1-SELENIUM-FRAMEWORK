package com.sdet34l1.genericUtility;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

import com.mysql.cj.jdbc.Driver;

/**
 * 
 * @author VAISHALI
 *
 */
public class DatabaseOffice {
	static Connection connection;
	static Statement statement;
/**
 * This method is used to get the data from database
 * @param columnName
 * @param query
 * @return
 */
	public static ArrayList<String> getDataFromDatabase(String columnName,String query)
	{
		ArrayList<String> list=new ArrayList<>();
	    ResultSet result = null;
		try {
			result = statement.executeQuery(query);
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			while(result.next())
			{
					list.add(result.getString(columnName));
								
				
			}
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return list;
	}
	/**
	 * This method is used to open the database connection
	 * @param dBUrl
	 * @param dbUserName
	 * @param dbPassword
	 * @throws SQLException
	 */
	public static void openDBconnection(String dBUrl,String dbUserName,String dbPassword) throws SQLException
	{
		Driver driver = new Driver();
	    DriverManager.registerDriver(driver);
	    connection = DriverManager.getConnection(dBUrl, dbUserName, dbPassword);
		 statement = connection.createStatement();
	}
	/**
	 * This method is used to validate the data in database
	 * @param columnName
	 * @param query
	 * @param expectedData
	 * @return
	 */
	public static boolean validatetheDatainDatabase(String columnName,String query, String expectedData) 
	{
		ArrayList<String> list = null;
		list = getDataFromDatabase(query, columnName);
		boolean flag=false;
		for(String actualData:list)
		{
			if(actualData.equalsIgnoreCase(expectedData))
			{
				flag=true;
			}
		}
		return flag;
	}
	/**
	 * This method is used to set the data in the data base
	 * @param query
	 */
	public static void setDataInDataBase(String query) 
	{
		 int result = 0;
		try {
			result = statement.executeUpdate(query);
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		 if(result>=1) {
			 System.out.println("Data entered/modified successfully");
		 }
		
	}
 /**
  * This method is used to close the data
  */
    public static void closeDatabase() 
    {
    	try {
			connection.close();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }

}
