package com.pc.driver;

import java.io.IOException;
import java.util.NoSuchElementException;
import java.util.concurrent.ConcurrentLinkedQueue;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import com.pc.utilities.HTML;
import com.pc.utilities.ReportUtil;
import com.pc.utilities.XlsxReader;

public class ParallelExecDriver 
{	
	   static  Logger log =Logger.getLogger(ParallelExecDriver.class);
//	   private String baseUrls;
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
					System.out.println( "Inside BeforeSuite"); 
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
	public void Parllel1(String DataSheetName, String Region) throws Exception
	{
		boolean isExitLoop = false;
		
		int nThreads = Integer.parseInt((String)HTML.properties.getProperty("THREAD_COUNT"));
		log.info("Total Number of Threads = "+HTML.properties.getProperty("THREAD_COUNT"));
		ExecutorService multiThreadExecutor = Executors.newFixedThreadPool(nThreads);
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
			log.info("Running test case in Prallel4 = " + testCaseName);
			ParallelExecutor parallelExecutor = new ParallelExecutor("RunModeNo",testCaseName,"",Region);
			//multiThreadExecutor.execute(parallelExecutor);
			multiThreadExecutor.submit(parallelExecutor);
		
	        try{
	        	testCaseName = testCaseIDorGroup.remove();
	        } catch(java.util.NoSuchElementException e){
	        	isExitLoop = true;
	        }
	        if (testCaseName==null){
	        	isExitLoop = true;
			}
		}
		multiThreadExecutor.shutdown();
		//while (!multiThreadExecutor.isTerminated()) {
        //}
		multiThreadExecutor.awaitTermination(24, TimeUnit.HOURS);
		log.info("Executed All Test Cases");
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
			}
			System.out.println( "Inside AfterSuite"); 
	}
}
