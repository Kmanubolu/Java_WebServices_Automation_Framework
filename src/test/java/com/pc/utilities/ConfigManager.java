/**
 * @ClassPurpose This Class is to create instance for Common and this instance will called across all other classes
 * @Scriptor Krishna Manubolu
 * @ReviewedBy
 * @ModifiedBy 
 * @LastDateModified 3/17/2017
 */

package com.pc.utilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

/**
 * This class provides methods to read values of keys from either app.properties or sys.properties from "ConfigFiles" folder
 */
public class ConfigManager 
{
	private Properties configProperty = new Properties();	
	private String configFileName;
	//private Logger logger = Logger.getLogger("ConfigManager");
	
	//The default constructor initializes and reads the key values from sys.properties file
	public ConfigManager()
	{
		configFileName = "Sys";
	}
	
	//The parameterized constructor when supplied with "app" as configFile name initializes and 
	//reads the key values from app.properties 
	public ConfigManager(String configname)
    {
		configFileName = configname;
    }
	
	/**
	 * Returns the value of given property from either sys.properties or app.properties file  
	 * @param key - ConfigParamvalue that requires to be returned from Config.properties file
	 * @return - return ConfigValue
	 * @throws Exception
	 */
	public String getProperty(String key) throws Exception
	{
        try
        {
            if(key!=null)
            {
            	this.loadProperties();         
                return configProperty.getProperty(key).trim();
            }
            else
            {
               // UtilityMethods.infoBox("One of the config parameter/property is given null - ConfigFiles->App/Sys.properties","Config Error");
               //logger.error("One of the config parameter/property is given null - ConfigFiles->App/Sys/OR.properties");
    			
            }
	    }
        catch (Exception E)
        {
        	//logger.error("Check the given Config property values"+E.getMessage()+UtilityMethods.getExceptionStackTrace(E));
           // UtilityMethods.infoBox("Check the given Config property values - Either the key "+key+" doesn't exist or mistyped \n"+UtilityMethods.getExceptionStackTrace(E),"Config Error");
           // System.exit(0);
        }
       return null;
    }


	/**
	 * To initialize and load the app.properties or sys.properties files based on the constructor called
	 * @throws IOException
	 */
	public void loadProperties() throws IOException
	{
		File directory = new File (".");
		String sConfigfilespath = directory.getCanonicalPath()+"\\Config";
		if(configFileName.equalsIgnoreCase("sys"))
        {	       
			loadSysPropertiesFile(sConfigfilespath);
        }
		else
        {
			//UtilityMethods.infoBox("Please enter valid configfile name(sys/app/OR) while creating object for configmanager class. The given name is "+configFileName,"Typo Error");
        }
	}
	
	/**
	 * @function This function use to identify the property file
	 * @param sConfigfilespath
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	private void loadSysPropertiesFile(String sConfigfilespath) throws FileNotFoundException, IOException {
		FileInputStream fis = new FileInputStream( sConfigfilespath+"\\Sys.properties");
		configProperty.load(fis);
		fis.close();
	}
}
