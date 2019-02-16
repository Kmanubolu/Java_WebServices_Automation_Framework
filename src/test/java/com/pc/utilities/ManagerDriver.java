package com.pc.utilities;

import org.openqa.selenium.WebDriver;

public class ManagerDriver 
{	
	   private static ManagerDriver instance = new ManagerDriver();
	
	   public static ManagerDriver getInstance()
	   {
	      return instance;
	   }
	   
	   ThreadLocal<WebDriver> driver = new ThreadLocal<WebDriver>();
	  	   
	   public WebDriver getWebDriver() 
	   {
	      return driver.get();
	   }
	   
	   public void setWebDriver(WebDriver webdriver) 
	   {
	        driver.set(webdriver);
	   }
}
