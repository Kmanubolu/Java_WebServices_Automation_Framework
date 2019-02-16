/**
 * @ClassPurpose This Class is to run the test in the remotely in the child machines
 * @Scriptor Banu Pradeep
 * @ReviewedBy
 * @ModifiedBy 
 * @LastDateModified 3/17/2017
 */
package com.pc.utilities;

import java.io.File;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Arrays;
import java.util.Collections;

import org.apache.log4j.Logger;
import org.openqa.selenium.Platform;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.phantomjs.PhantomJSDriverService;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;

public class RemoteDriverFactory {
	static Logger log = Logger.getLogger(RemoteDriverFactory.class);
	
	private static RemoteDriverFactory instance = new RemoteDriverFactory();
	
		
	  
	
	   public static RemoteDriverFactory getInstance()
	   {
	      return instance;
	   }
	   
	   		   
	   public  WebDriver createNewDriver()
	   {
		   	  RemoteWebDriver driver = null;
	    	  String driverPath = HTML.properties.getProperty("driverPath");
	  		  String execution = HTML.properties.getProperty("TypeOfAutomation");
	  		  String browser = HTML.properties.getProperty("Browser");
	  		  PhantomJSDriverService service = null;
	  		  DesiredCapabilities cap = null;
	  		  if(execution.toUpperCase().equalsIgnoreCase("HEADLESS"))
				{
	  			 	log.info("Webdriver Headless Automation started");
					/*final String phantomJsFilePath = driverPath + "\\phantomjs.exe";
					service = new PhantomJSDriverService.Builder().usingPhantomJSExecutable(new File(phantomJsFilePath)).usingAnyFreePort()
					.build();
					try {
						service.start();
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					} 
					log.info("Headless phantomjs service started");*/
					if (browser.toUpperCase().contains("CH"))
					{	
						 cap = DesiredCapabilities.chrome();
					}
					if (browser.toUpperCase().contains("FF"))
					{
						cap = DesiredCapabilities.firefox();
					}
					if (browser.toUpperCase().contains("IE"))
					{
						cap = DesiredCapabilities.internetExplorer();
					}
					try {
						driver = new RemoteWebDriver(new URL(ManagerPhantomJS.getInstance().getPhantomJSDrivrService().getUrl().toString()), cap);
					} catch (MalformedURLException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					log.info("Headless '" + browser.toUpperCase() + "' DesiredCapabilities initiated");
					log.info("Headless '" + browser.toUpperCase() + "' driver started");	
				}
				if (execution.toUpperCase().equalsIgnoreCase("HEAD"))
				{
					log.info("Webdriver Grid Automation started");
					if(browser.toUpperCase().contains("CH"))
					{	
						ChromeOptions options =  new ChromeOptions();
						options.addArguments("no-sandbox");
						options.addArguments("start-maximized");
						options.setExperimentalOption("excludeSwitches", Collections.singletonList("enable-automation"));
						cap = DesiredCapabilities.chrome();
						cap.setCapability(ChromeOptions.CAPABILITY, options);
						cap.setBrowserName("chrome");
						cap.setPlatform(Platform.WINDOWS);
					}
					if(browser.toUpperCase().contains("FF"))
					{
						cap = DesiredCapabilities.firefox();
						cap.setBrowserName("firefox");
						cap.setPlatform(Platform.WINDOWS);
					}
					if (browser.toUpperCase().contains("IE"))
					{
						cap = DesiredCapabilities.internetExplorer();
						cap.setBrowserName("iexplore");
						cap.setPlatform(Platform.WINDOWS);
					}
					try 
					{
						driver = new RemoteWebDriver(new URL("http://localhost:4444/wd/hub"),cap);
					} catch (MalformedURLException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					log.info("Headless '" + browser.toUpperCase() + "' DesiredCapabilities initiated");
					log.info("Headless '" + browser.toUpperCase() + "' driver started");	
				}
			  return driver; // can be replaced with other browser drivers
		}

}
