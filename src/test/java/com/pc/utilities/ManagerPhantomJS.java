package com.pc.utilities;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.phantomjs.PhantomJSDriverService;;
public class ManagerPhantomJS 
{	
	   private static ManagerPhantomJS instance = new ManagerPhantomJS();
	
	   public static ManagerPhantomJS getInstance()
	   {
	      return instance;
	   }
	   
	   ThreadLocal<PhantomJSDriverService> service = new ThreadLocal<PhantomJSDriverService>();
	  	   
	   public PhantomJSDriverService getPhantomJSDrivrService() 
	   {
	      return service.get();
	   }
	   
	   public void setPhantomJSDrivrService(PhantomJSDriverService phantomJSDrivrService) 
	   {
	        service.set(phantomJSDrivrService);
	   }
}
