package com.pc.driver;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.NoSuchElementException;
import java.util.concurrent.ConcurrentLinkedQueue;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import com.pc.constants.PCConstants;
import com.pc.utilities.HTML;
import com.pc.utilities.HTML;
import com.pc.utilities.ReportUtil;
import com.pc.utilities.XlsxReader;

public class ParallelRegressionExecDriver 
{	
	   static  Logger log =Logger.getLogger(ParallelRegressionExecDriver.class);
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

	     
	@Parameters({ "DataSheetName","Region" })
	@Test(priority=1,enabled=true)
	public void Parllel1(String DataSheetName, String Region) 
	{
		boolean isExitLoop = false;
		boolean isExecuteByGroup = false;
		HashMap<String,Object> whereConstraint = new HashMap<String,Object>();
		ArrayList<HashMap<String,Object>> resultListRows= new ArrayList<HashMap<String,Object>>();
		
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
						//ExecutorService multiThreadExecutor = Executors.newFixedThreadPool(1); 
				        ParallelExecutor parallelExecutor = new ParallelExecutor("RunModeNo",testCaseName,"",Region);
						Thread t = new Thread(parallelExecutor);
						t.start();
						try {
							t.join();
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
				        //multiThreadExecutor.submit(parallelExecutor);
						//multiThreadExecutor.shutdown();
						//while (!multiThreadExecutor.isTerminated()) {
				       // }
						log.info("Executed All Test Cases");
					}
				}
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
	@Test(priority=1,enabled=true)
	public void Parllel2(String DataSheetName, String Region) 
	{
		boolean isExitLoop = false;
		boolean isExecuteByGroup = false;
		HashMap<String,Object> whereConstraint = new HashMap<String,Object>();
		ArrayList<HashMap<String,Object>> resultListRows= new ArrayList<HashMap<String,Object>>();
		
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
						//ExecutorService multiThreadExecutor = Executors.newFixedThreadPool(1); 
				        ParallelExecutor parallelExecutor = new ParallelExecutor("RunModeNo",testCaseName,"",Region);
						Thread t = new Thread(parallelExecutor);
						t.start();
						try {
							t.join();
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
				        //multiThreadExecutor.submit(parallelExecutor);
						//multiThreadExecutor.shutdown();
						//while (!multiThreadExecutor.isTerminated()) {
				       // }
						log.info("Executed All Test Cases");
					}
				}
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
	@Test(priority=1,enabled=true)
	public void Parllel3(String DataSheetName, String Region) 
	{
		boolean isExitLoop = false;
		boolean isExecuteByGroup = false;
		HashMap<String,Object> whereConstraint = new HashMap<String,Object>();
		ArrayList<HashMap<String,Object>> resultListRows= new ArrayList<HashMap<String,Object>>();
		
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
						//ExecutorService multiThreadExecutor = Executors.newFixedThreadPool(1); 
				        ParallelExecutor parallelExecutor = new ParallelExecutor("RunModeNo",testCaseName,"",Region);
						Thread t = new Thread(parallelExecutor);
						t.start();
						try {
							t.join();
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
				        //multiThreadExecutor.submit(parallelExecutor);
						//multiThreadExecutor.shutdown();
						//while (!multiThreadExecutor.isTerminated()) {
				       // }
						log.info("Executed All Test Cases");
					}
				}
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
	@Test(priority=1,enabled=true)
	public void Parllel4(String DataSheetName, String Region) 
	{
		boolean isExitLoop = false;
		boolean isExecuteByGroup = false;
		HashMap<String,Object> whereConstraint = new HashMap<String,Object>();
		ArrayList<HashMap<String,Object>> resultListRows= new ArrayList<HashMap<String,Object>>();
		
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
						//ExecutorService multiThreadExecutor = Executors.newFixedThreadPool(1); 
				        ParallelExecutor parallelExecutor = new ParallelExecutor("RunModeNo",testCaseName,"",Region);
						Thread t = new Thread(parallelExecutor);
						t.start();
						try {
							t.join();
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
				        //multiThreadExecutor.submit(parallelExecutor);
						//multiThreadExecutor.shutdown();
						//while (!multiThreadExecutor.isTerminated()) {
				       // }
						log.info("Executed All Test Cases");
					}
				}
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
	@Test(priority=1,enabled=true)
	public void Parllel5(String DataSheetName, String Region) 
	{
		boolean isExitLoop = false;
		boolean isExecuteByGroup = false;
		HashMap<String,Object> whereConstraint = new HashMap<String,Object>();
		ArrayList<HashMap<String,Object>> resultListRows= new ArrayList<HashMap<String,Object>>();
		
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
						//ExecutorService multiThreadExecutor = Executors.newFixedThreadPool(1); 
				        ParallelExecutor parallelExecutor = new ParallelExecutor("RunModeNo",testCaseName,"",Region);
						Thread t = new Thread(parallelExecutor);
						t.start();
						try {
							t.join();
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
				        //multiThreadExecutor.submit(parallelExecutor);
						//multiThreadExecutor.shutdown();
						//while (!multiThreadExecutor.isTerminated()) {
				       // }
						log.info("Executed All Test Cases");
					}
				}
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
	@Test(priority=1,enabled=true)
	public void Parllel6(String DataSheetName, String Region) 
	{
		boolean isExitLoop = false;
		boolean isExecuteByGroup = false;
		HashMap<String,Object> whereConstraint = new HashMap<String,Object>();
		ArrayList<HashMap<String,Object>> resultListRows= new ArrayList<HashMap<String,Object>>();
		
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
						//ExecutorService multiThreadExecutor = Executors.newFixedThreadPool(1); 
				        ParallelExecutor parallelExecutor = new ParallelExecutor("RunModeNo",testCaseName,"",Region);
						Thread t = new Thread(parallelExecutor);
						t.start();
						try {
							t.join();
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
				        //multiThreadExecutor.submit(parallelExecutor);
						//multiThreadExecutor.shutdown();
						//while (!multiThreadExecutor.isTerminated()) {
				       // }
						log.info("Executed All Test Cases");
					}
				}
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
	@Test(priority=1,enabled=true)
	public void Parllel7(String DataSheetName, String Region) 
	{
		boolean isExitLoop = false;
		boolean isExecuteByGroup = false;
		HashMap<String,Object> whereConstraint = new HashMap<String,Object>();
		ArrayList<HashMap<String,Object>> resultListRows= new ArrayList<HashMap<String,Object>>();
		
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
						//ExecutorService multiThreadExecutor = Executors.newFixedThreadPool(1); 
				        ParallelExecutor parallelExecutor = new ParallelExecutor("RunModeNo",testCaseName,"",Region);
						Thread t = new Thread(parallelExecutor);
						t.start();
						try {
							t.join();
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
				        //multiThreadExecutor.submit(parallelExecutor);
						//multiThreadExecutor.shutdown();
						//while (!multiThreadExecutor.isTerminated()) {
				       // }
						log.info("Executed All Test Cases");
					}
				}
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
	@Test(priority=1,enabled=true)
	public void Parllel8(String DataSheetName, String Region) 
	{
		boolean isExitLoop = false;
		boolean isExecuteByGroup = false;
		HashMap<String,Object> whereConstraint = new HashMap<String,Object>();
		ArrayList<HashMap<String,Object>> resultListRows= new ArrayList<HashMap<String,Object>>();
		
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
						//ExecutorService multiThreadExecutor = Executors.newFixedThreadPool(1); 
				        ParallelExecutor parallelExecutor = new ParallelExecutor("RunModeNo",testCaseName,"",Region);
						Thread t = new Thread(parallelExecutor);
						t.start();
						try {
							t.join();
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
				        //multiThreadExecutor.submit(parallelExecutor);
						//multiThreadExecutor.shutdown();
						//while (!multiThreadExecutor.isTerminated()) {
				       // }
						log.info("Executed All Test Cases");
					}
				}
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
	@Test(priority=1,enabled=true)
	public void Parllel9(String DataSheetName, String Region) 
	{
		boolean isExitLoop = false;
		boolean isExecuteByGroup = false;
		HashMap<String,Object> whereConstraint = new HashMap<String,Object>();
		ArrayList<HashMap<String,Object>> resultListRows= new ArrayList<HashMap<String,Object>>();
		
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
						//ExecutorService multiThreadExecutor = Executors.newFixedThreadPool(1); 
				        ParallelExecutor parallelExecutor = new ParallelExecutor("RunModeNo",testCaseName,"",Region);
						Thread t = new Thread(parallelExecutor);
						t.start();
						try {
							t.join();
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
				        //multiThreadExecutor.submit(parallelExecutor);
						//multiThreadExecutor.shutdown();
						//while (!multiThreadExecutor.isTerminated()) {
				       // }
						log.info("Executed All Test Cases");
					}
				}
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
	@Test(priority=1,enabled=true)
	public void Parllel10(String DataSheetName, String Region) 
	{
		boolean isExitLoop = false;
		boolean isExecuteByGroup = false;
		HashMap<String,Object> whereConstraint = new HashMap<String,Object>();
		ArrayList<HashMap<String,Object>> resultListRows= new ArrayList<HashMap<String,Object>>();
		
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
						//ExecutorService multiThreadExecutor = Executors.newFixedThreadPool(1); 
				        ParallelExecutor parallelExecutor = new ParallelExecutor("RunModeNo",testCaseName,"",Region);
						Thread t = new Thread(parallelExecutor);
						t.start();
						try {
							t.join();
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
				        //multiThreadExecutor.submit(parallelExecutor);
						//multiThreadExecutor.shutdown();
						//while (!multiThreadExecutor.isTerminated()) {
				       // }
						log.info("Executed All Test Cases");
					}
				}
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
			}
			System.out.println( "Inside AfterSuite"); 
	}
}
