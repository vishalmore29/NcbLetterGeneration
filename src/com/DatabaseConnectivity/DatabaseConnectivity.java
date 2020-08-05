package com.DatabaseConnectivity;

import java.sql.Connection;
import java.sql.DriverManager;

public class DatabaseConnectivity 
{
	public static Connection getDatabaseConnection() 
	{
		Connection conn = null;
		try 
		{
			Class.forName("com.mysql.jdbc.Driver");  
			conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/psu_ncb_confirma","root","admin123");
			System.out.println("Database connection estabhlished.");
		}
		catch(Exception ex) 
		{
			ex.printStackTrace();
		}
		return conn;
	}
	
	/*public static void main(String argsp[]) 
	{
		DatabaseConnectivity obj = new DatabaseConnectivity();
		obj.getDatabaseConnection();
	}*/
}
