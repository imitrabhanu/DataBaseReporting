package com.novopay.masteronboarding.report.prod;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.Properties;

import org.apache.log4j.Logger;

public class DatabaseConnection {
	private static Logger logger = Logger.getLogger("DatabaseConnection.class");
	
	/**
	 *Database Connection 
	*/
	public static Connection getConnection()
	{ Properties properties= new Properties();
		FileInputStream fileInput= null;
		Connection connection= null;
		try{
			fileInput =new FileInputStream("./Properties/DBProperties.properties");
			properties.load(fileInput);
			fileInput.close();
			//Load the driver class
			Class.forName(properties.getProperty("jdbc.driver"));
			//Create the connection
			connection=DriverManager.getConnection(properties.getProperty("jdbc.url"),
					properties.getProperty("jdbc.username"),properties.getProperty("jdbc.password"));
	
		}
		catch (SQLException se )
		{
			logger.error("SQL Exception",se);
		}
		catch (FileNotFoundException fe )
		{
			logger.error("File Not Found",fe);
		}
		catch (IOException ie )
		{
			logger.error("IO Exception",ie);
		}
		catch (Exception e )
		{
			logger.error("Exception",e);
		}
		logger.debug("Connection Successful");
		return connection;
	}

}