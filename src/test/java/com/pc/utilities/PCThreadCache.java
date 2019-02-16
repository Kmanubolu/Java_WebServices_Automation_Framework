/**
 * @ClassPurpose This Class is to create the instance of properties file and can be used across the classes
 * @Scriptor Banu Pradeep
 * @ReviewedBy
 * @ModifiedBy 
 * @LastDateModified 3/17/2017
 */
package com.pc.utilities;

import java.util.Properties;

public class PCThreadCache {
	
	private static PCThreadCache instance = new PCThreadCache();
	
	public static PCThreadCache getInstance()
	{
		return instance;
	}
	
    ThreadLocal<Properties> properties = new ThreadLocal<Properties>() 
    {
       @Override
       protected Properties initialValue()
       {
        	 return new Properties();
       }
    };
			   
   	public String getProperty(String key) 
   	{
   		return properties.get().getProperty(key);      
   	}

    public void setProperty(String key, String value)
    {
	    properties.get().setProperty(key, value);
    }
    
	public void resetProperties()
	{
		Properties prop = properties.get();
		prop.clear();
		/* for (String key : prop.stringPropertyNames()) 
		 {
			 prop.remove(key);
	     }*/
     }
}
