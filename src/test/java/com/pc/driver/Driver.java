package com.pc.driver;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.NoSuchElementException;
import java.util.concurrent.ConcurrentLinkedQueue;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.phantomjs.PhantomJSDriverService;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;
import com.pc.constants.PCConstants;
import com.pc.utilities.Common;
import com.pc.utilities.CommonManager;
import com.pc.utilities.HTML;
import com.pc.utilities.LocalDriverFactory;
import com.pc.utilities.ManagerDriver;
import com.pc.utilities.ManagerPhantomJS;
import com.pc.utilities.PhantomJSDriverFactory;
import com.pc.utilities.RemoteDriverFactory;
import com.pc.utilities.ReportUtil;
import com.pc.utilities.XlsxReader;

public class Driver 
{	
	   static  Logger log =Logger.getLogger(Common.class);
	   
	   private static ConcurrentLinkedQueue<String> testCaseIDorGroup = new ConcurrentLinkedQueue<String>();
	   
	   /**
	    * @function This method will take all the TCID into the queue will run one after one
	    * @throws Exception
	    */
	   @BeforeSuite
	   public void loadConfig() throws Exception 
	   {
		   	boolean status = false;
		    PropertyConfigurator.configure("log4j.properties");
			try 
			{
					HTML.fnSummaryInitialization("Execution Summary Report");
					status = XlsxReader.getInstance().addTestCasesFromDataSheetName(testCaseIDorGroup);
					System.out.println("Inside BeforeSuite");
					if(HTML.properties.getProperty("DataBaseUpdate").equalsIgnoreCase("YES"))
					{
						ReportUtil.loadTestCases();
					}
						if(!status)
						{
							log.info("None of the testcase selected as 'YES' to execute");
							System.exit(0);
						}
			} 
			catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				log.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
			}	
	   }
	   
    /**
    * @function This method use to run the test case parllel according to the thread through TestNG
    * @param DataSheetName
    * @param Region
    * @throws Exception
    */
	@Parameters({ "DataSheetName","Region" })
	@Test(priority=1,enabled=true)
	public void Parllel1(String DataSheetName, String Region) 
	{
		boolean isExitLoop = false;
		boolean isExecuteByGroup = false;
		HashMap<String,Object> whereConstraint = new HashMap<String,Object>();
		ArrayList<HashMap<String,Object>> resultListRows= new ArrayList<HashMap<String,Object>>();
		String execMode = HTML.properties.getProperty("EXECUTIONMODE");
		if(PCConstants.YES.equalsIgnoreCase((String)HTML.properties.get("EXECUTE_BY_TEST_CASE_GROUP"))){
			isExecuteByGroup = true;
		}
		String testCaseName = null;
		try{
        	testCaseName = testCaseIDorGroup.remove();
        } catch(NoSuchElementException e){
        	isExitLoop = true;
        }
		if (testCaseName==null){
			isExitLoop = true;
		}
		while(!isExitLoop)
		{
			log.info("Running test case in Prallel4 = " + testCaseName);	
			if(isExecuteByGroup){
				whereConstraint.clear();
				whereConstraint.put(PCConstants.TestCaseGroup, testCaseName);
				whereConstraint.put(PCConstants.Execution, PCConstants.YES);
				resultListRows = XlsxReader.getInstance().executeSelectQuery(PCConstants.SHEET_TESTCASE, whereConstraint);
				if(!resultListRows.isEmpty())
				{
					for(HashMap<String,Object> processRow: resultListRows)
					{
						testCaseName = (String)processRow.get(PCConstants.ID);
						WebDriver driver = null;
						PhantomJSDriverService service = null;
						String execution  = HTML.properties.getProperty("TypeOfAutomation");
						if(execution.toUpperCase().contains("HRADLESS"))
						{
							service = PhantomJSDriverFactory.getInstance().createPhantomJSDriver();
							ManagerPhantomJS.getInstance().setPhantomJSDrivrService(service);
						}
						if(execMode.equalsIgnoreCase(PCConstants.executionModeLocal)){
							driver = LocalDriverFactory.getInstance().createNewDriver();
						}else {
							driver = RemoteDriverFactory.getInstance().createNewDriver();
						}
						Common common = new Common();
						CommonManager.getInstance().setCommon(common);
				        ManagerDriver.getInstance().setWebDriver(driver);
				        log.info("Thread ID = " + Thread.currentThread().getId() + " common = "+ common +" testCaseName= "+testCaseName);
				        try {
							common.RunTest("RunModeNo",testCaseName,"",Region);
						} catch (Exception e) {
							e.printStackTrace();
							log.error("Error occured while executing test case = " + testCaseName, e);
						}
				        common = null; //Mark for garbage collection
					}
				}
			} else {
				WebDriver driver = null;
				PhantomJSDriverService service = null;
				String execution  = HTML.properties.getProperty("TypeOfAutomation");
				if(execution.toUpperCase().contains("HEADLESS"))
				{
					service = PhantomJSDriverFactory.getInstance().createPhantomJSDriver();
					ManagerPhantomJS.getInstance().setPhantomJSDrivrService(service);
				}
				if(execMode.equalsIgnoreCase(PCConstants.executionModeLocal)){
					driver = LocalDriverFactory.getInstance().createNewDriver();
				}else {
					driver = RemoteDriverFactory.getInstance().createNewDriver();
				}
				Common common = new Common();
				CommonManager.getInstance().setCommon(common);
		        ManagerDriver.getInstance().setWebDriver(driver);
		        log.info("Thread ID = " + Thread.currentThread().getId() + " common = "+ common +" testCaseName= "+testCaseName);
		        try {
					common.RunTest("RunModeNo",testCaseName,"",Region);
				} catch (Exception e) {
					e.printStackTrace();
					log.error("Error occured while executing test case = " + testCaseName, e);
				}
		        common = null; //Mark for garbage collection
			}
	        try{
	        	testCaseName = testCaseIDorGroup.remove();
	        } catch(java.util.NoSuchElementException e){
	        	isExitLoop = true;
	        }
	        if (testCaseName==null){
	        	isExitLoop = true;
			}
		}
	}
		@Parameters({ "DataSheetName","Region" })
	@Test(priority=1,enabled=false)
	public void Parllel8(String DataSheetName, String Region) 
	{
		boolean isExitLoop = false;
		boolean isExecuteByGroup = false;
		HashMap<String,Object> whereConstraint = new HashMap<String,Object>();
		ArrayList<HashMap<String,Object>> resultListRows= new ArrayList<HashMap<String,Object>>();
		String execMode = HTML.properties.getProperty("EXECUTIONMODE");
		if(PCConstants.YES.equalsIgnoreCase((String)HTML.properties.get("EXECUTE_BY_TEST_CASE_GROUP"))){
			isExecuteByGroup = true;
		}
		String testCaseName = null;
		try{
        	testCaseName = testCaseIDorGroup.remove();
        } catch(NoSuchElementException e){
        	isExitLoop = true;
        }
		if (testCaseName==null){
			isExitLoop = true;
		}
		while(!isExitLoop){
			log.info("Running test case in Prallel8 = " + testCaseName);
			
			if(isExecuteByGroup){
				whereConstraint.clear();
				whereConstraint.put(PCConstants.TestCaseGroup, testCaseName);
				whereConstraint.put(PCConstants.Execution, PCConstants.YES);
				resultListRows = XlsxReader.getInstance().executeSelectQuery(PCConstants.SHEET_TESTCASE, whereConstraint);
				if(!resultListRows.isEmpty())
				{
					for(HashMap<String,Object> processRow: resultListRows)
					{
						testCaseName = (String)processRow.get(PCConstants.ID);
						WebDriver driver = null;
						if(execMode.equalsIgnoreCase(PCConstants.executionModeLocal)){
							driver = LocalDriverFactory.getInstance().createNewDriver();
						}else {
							driver = RemoteDriverFactory.getInstance().createNewDriver();
						}
						Common common = new Common();
						CommonManager.getInstance().setCommon(common);
				        ManagerDriver.getInstance().setWebDriver(driver);
				        log.info("Thread ID = " + Thread.currentThread().getId() + " common = "+ common +" testCaseName= "+testCaseName);
				        try {
							common.RunTest("RunModeNo",testCaseName,"",Region);
						} catch (Exception e) {
							e.printStackTrace();
							log.error("Error occured while executing test case = " + testCaseName, e);
						}
				        common = null; //Mark for garbage collection
					}
				}
			} else {
				WebDriver driver = null;
				if(execMode.equalsIgnoreCase(PCConstants.executionModeLocal)){
					driver = LocalDriverFactory.getInstance().createNewDriver();
				}else {
					driver = RemoteDriverFactory.getInstance().createNewDriver();
				}
				Common common = new Common();
				CommonManager.getInstance().setCommon(common);
		        ManagerDriver.getInstance().setWebDriver(driver);
		        log.info("Thread ID = " + Thread.currentThread().getId() + " common = "+ common +" testCaseName= "+testCaseName);
		        try {
					common.RunTest("RunModeNo",testCaseName,"",Region);
				} catch (Exception e) {
					e.printStackTrace();
					log.error("Error occured while executing test case = " + testCaseName, e);
				}
		        common = null; //Mark for garbage collection
			}
	        
	        try{
	        	testCaseName = testCaseIDorGroup.remove();
	        } catch(java.util.NoSuchElementException e){
	        	isExitLoop = true;
	        }
	        if (testCaseName==null){
	        	isExitLoop = true;
			}
		}
	}
	
	/**
	 * @function This function is to exit the connection of the browser
	 */
	@AfterSuite
	public void exitConfig() {
	
	    	try {
	    		XlsxReader.getInstance().closeConnections();
				HTML.fnSummaryCloseHtml(HTML.properties.getProperty("Release"));
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				log.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
			}
			System.out.println( "Inside AfterSuite"); 
	}
}
