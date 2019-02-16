package com.pc.utilities;

import java.net.InetAddress;
import java.net.UnknownHostException;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.log4j.Logger;
import org.openqa.selenium.remote.RemoteWebDriver;

import com.pc.constants.PCConstants;
import com.pc.dao.ReportUtilityDAO;

public class ReportUtil {
	static  Logger logger =Logger.getLogger(ReportUtil.class); // Logger variable
	
   public static void loadTestCases()
   {   
	    ReportUtilityDAO  reportUtilDao = new ReportUtilityDAO();
	    XlsxReader sXL = XlsxReader.getInstance();
		HashMap<String,Object> whereConstraint = new HashMap<String,Object>();
		HashMap<String,Object> updateColumnNameValues = new HashMap<String,Object>();
		ArrayList<HashMap<String,Object>> resultListRows= new ArrayList<HashMap<String,Object>>();
		whereConstraint.put(PCConstants.TestCaseType, "Regression"); 
		whereConstraint.put(PCConstants.Execution, "YES");
		resultListRows = sXL.executeSelectQuery(PCConstants.SHEET_TESTCASE, whereConstraint);
		reportUtilDao.populateTestExecStatus(resultListRows);
	    reportUtilDao.closeConnections();
   }
   
   public static void initBeginExecuction(){
	   String testCaseID = PCThreadCache.getInstance().getProperty("TCID");
	   RemoteWebDriver webDriver = (RemoteWebDriver) ManagerDriver.getInstance().getWebDriver();
	   String execMachineName = null;
	   String ipAddress = null;
		try {
			String execMode = HTML.properties.getProperty("EXECUTIONMODE");
			if(execMode.equalsIgnoreCase(PCConstants.executionModeLocal)){
				try {
					  InetAddress addr = InetAddress.getLocalHost();            
					  ipAddress = addr.getHostAddress(); //gets the ip address of current machine
					  String computerName = addr.getHostName(); //gets computer name of current machine
					  
					  execMachineName = ipAddress;
				} catch (UnknownHostException e) {
				}	
			}else {
				execMachineName = GridInfoExtracter.getHostName(webDriver.getSessionId(), HTML.properties.getProperty("MACHINE"));
			}
			
		} catch ( Exception e) {
			// TODO Auto-generated catch block
			logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
			e.printStackTrace();
		}
	   ReportUtilityDAO  reportUtilDao = new ReportUtilityDAO();
	   
	    reportUtilDao.updateTestExecMachine(testCaseID, execMachineName);
	    reportUtilDao.closeConnections();
	   
   }
   
   public static void finalizeExec(String status){
	   ReportUtilityDAO  reportUtilDao = new ReportUtilityDAO();
	   String testCaseID = PCThreadCache.getInstance().getProperty("TCID");
	   reportUtilDao.updateTestExecStatus(testCaseID, status);
	   reportUtilDao.closeConnections();
   }
   
   public static synchronized void updateDataFeed(String valueToUpdate){
	   ReportUtilityDAO  reportUtilDao = new ReportUtilityDAO();
	   String testCaseID = PCThreadCache.getInstance().getProperty("TCID");
	   reportUtilDao.populateDataFeed(testCaseID, valueToUpdate);
	   reportUtilDao.closeConnections();
   }
   
}
