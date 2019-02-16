package com.pc.dao;

import java.net.InetAddress;
import java.net.UnknownHostException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;

import org.testng.log4testng.Logger;


public class ReportUtilityDAO{


	private Connection ucaConn;
	private Statement 	st;
	private ResultSet rs;
	private static Logger log = Logger.getLogger(ReportUtilityDAO.class);
	
	public ReportUtilityDAO(){
		try {
			boolean isValidConnection = false;
			int attempt =0;
			while(!isValidConnection && attempt < 10) {
				try{
					 attempt++;
					 Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
					 //String url = "jdbc:ucanaccess:C://Selenium//WorkSpace//DB//RegressionDB//EXECUTIONREPORT.accdb";
					 String url = "jdbc:ucanaccess:C://Selenium//WorkSpace//DB//RegressionDB//EXECUTIONREPORT.accdb";
					 
					 //String url = "jdbc:ucanaccess:////ad1.prod/hig/commercial/ComPAS/temp/E2E/Regression/EXECUTIONREPORT.accdb"; 
					 ucaConn = DriverManager.getConnection(url);
				
					// ucaConn.setAutoCommit(true);
					 ucaConn.setTransactionIsolation(Connection.TRANSACTION_READ_UNCOMMITTED);
					 st = ucaConn.createStatement();
					 //TestCaseID
					 ResultSet rsTest= st.executeQuery("select 1 from dual");
					 if(rsTest.next()){
						 log.info("Acquired valid connection attempt =" + attempt);
						 isValidConnection = true;
					 }
					 
				}catch (SQLException e) {
					log.error("Error getting connection from ucanaccess driver attempt =" +attempt);
					// TODO Auto-generated catch block
					e.printStackTrace();
				} 
			}
		} catch (ClassNotFoundException e) {
			log.error("Error getting connection from ucanaccess driver 1");
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
	}
	
	
	
	public String retrieveTransKeyValue(String e2eTestCaseId, String keyName){
		String policyNumber = null;
		try {
			rs= st.executeQuery("Select TRANS_KEY_VALUE from TRANS where E2E_TCASE_ID='"+e2eTestCaseId+"' and TRANS_KEY_NM = '"+keyName+"' order by LAST_UPDT_TMSP desc;");
			if(rs.next()){
				policyNumber = rs.getString(1);
			}
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			/*try {
				st.close();
				ucaConn.close();
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}*/
		}
		
		return policyNumber;
	}
	
		
	public boolean updateE2ETrans(String e2eTestCaseId, String keyName,String keyValue){
		boolean status = false;
		Date curDate = new Date();
		SimpleDateFormat format = new SimpleDateFormat("M/d/yyyy h:m:s aaa");
		String dateToStr = format.format(curDate);
		try {
			//String query =  "MERGE INTO TRANS bi USING (Select '"+ e2eTestCaseId +"' as E2E_TCASE_ID from dual) bo ON bi.E2E_TCASE_ID = bo.E2E_TCASE_ID" +
              //      " WHEN MATCHED THEN Update set POLICYNO = '"+ policyNumber +"', APPLICATION = 'Selenium', LAST_UPDATE_TIMESTAMP ='"+dateToStr +"', E2E_TCASE_ID = '"+e2eTestCaseId +"' WHEN NOT MATCHED THEN insert(POLICYNO, E2E_TCASE_ID, APPLICATION, LAST_UPDATE_TIMESTAMP) values ( '"+ policyNumber +"','"+e2eTestCaseId +"','Selenium','"+dateToStr+ "') ";
			String query = "insert into TRANS(TRANS_KEY_NM, TRANS_KEY_VALUE,E2E_TEST_CASE_ID,LAST_UPDT_TMSP,LAST_UPDT_APPL_NM) VALUES('"+keyName+"','"+keyValue+"','"+e2eTestCaseId+"','"+dateToStr+"','PC')";
			log.info("executing query = "+query);
			status = st.execute(query);
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			/*try {
			st.close();
			ucaConn.close();
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}*/
		}
		return status;
	
	}
	
	public boolean updateE2EMasterStatus( String Id, String statusValue){
		boolean status = false;
		Date curDate = new Date();
		SimpleDateFormat format = new SimpleDateFormat("M/d/yyyy h:m:s aaa");
		String dateToStr = format.format(curDate);
		String computerName = null;
		String ipAddress = null;
		String lastUpdateTimeStamp = null;
		String duration = "";
		
		try {
			  InetAddress addr = InetAddress.getLocalHost();            
			  ipAddress = addr.getHostAddress(); //gets the ip address of current machine
			  computerName = addr.getHostName(); //gets computer name of current machine
		} catch (UnknownHostException e) {
		}
		
		try {
			if(statusValue != null && statusValue.length() >0 
					&& (statusValue.equalsIgnoreCase("Passed") || statusValue.equalsIgnoreCase("Failed"))){
					String getLastUpdateTimeStamp = "Select LAST_UPDT_TMSP from E2E_TESTCASES where SEQ_NUM = '"+Id+"'";
					Statement st1 = ucaConn.createStatement();
					rs = st1.executeQuery(getLastUpdateTimeStamp);
					if(rs.next()){
						lastUpdateTimeStamp = rs.getString(1);
					}
					
					if( lastUpdateTimeStamp != null && lastUpdateTimeStamp.length() >0){
						Date date1;
						try {
							date1 = format.parse(lastUpdateTimeStamp);
							long difference = curDate.getTime() - date1.getTime(); //milliseconds
							long diffMinutes = difference / (60 * 1000) % 60;
							duration = String.valueOf(diffMinutes);
						} catch (ParseException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						
					}
					
					st1.close();
			}
			//setting the computername as grid
			computerName = "GRID";
			String query =  "MERGE INTO E2E_TESTCASES bi USING (Select '"+ Id +"' as SEQ_NUM from dual) bo ON bi.SEQ_NUM = bo.SEQ_NUM" +
                    " WHEN MATCHED THEN Update set TEST_CASE_STAT_DESC = '"+ statusValue +"', EXEC_ELAPS_MIN='"+duration +"',EXEC_MACH_NM ='"+computerName+"',LAST_UPDT_TMSP = '"+dateToStr +"' WHEN NOT MATCHED THEN insert(TEST_CASE_STAT_DESC,EXEC_ELAPS_MIN, EXEC_MACH_NM, LAST_UPDT_TMSP) values ( '"+ statusValue +"','"+duration+"','"+computerName +"','"+dateToStr +"' ) ";
			
			log.info("executing query = "+query);
			status = st.execute(query);
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			/*try {
				st.close();
				ucaConn.close();
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}*/
		}
		return status;
	
	}
	
	public void populateTestExecStatus(ArrayList<HashMap<String,Object>> resultListRows){
		
		try {
			
			Statement stmt = ucaConn.createStatement();
			if(!resultListRows.isEmpty())
			{
				for(HashMap<String,Object> processRow: resultListRows)
				{
					
					stmt.addBatch("INSERT INTO TEST_EXEC_STATUS (TEST_CASE_ID,TEST_CASE_NM,TEST_CASE_STAT_DESC,MODULE ) "+ 
					   "VALUES ('"+(String)processRow.get("ID")+"','" +(String)processRow.get("TestCaseName")+"','" +"Initialized"+"','" +(String)processRow.get("Use Case")+"')");
				}
			}
			stmt.executeBatch();

		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	public boolean updateTestExecMachine(String testCaseID, String execMachineName){
		boolean status = false;
		Date curDate = new Date();
		SimpleDateFormat format = new SimpleDateFormat("M/d/yyyy h:m:s aaa");
		String dateToStr = format.format(curDate);
		try {
			String threadId = String.valueOf(Thread.currentThread().getId());
			//String query =  "MERGE INTO TRANS bi USING (Select '"+ e2eTestCaseId +"' as E2E_TCASE_ID from dual) bo ON bi.E2E_TCASE_ID = bo.E2E_TCASE_ID" +
              //      " WHEN MATCHED THEN Update set POLICYNO = '"+ policyNumber +"', APPLICATION = 'Selenium', LAST_UPDATE_TIMESTAMP ='"+dateToStr +"', E2E_TCASE_ID = '"+e2eTestCaseId +"' WHEN NOT MATCHED THEN insert(POLICYNO, E2E_TCASE_ID, APPLICATION, LAST_UPDATE_TIMESTAMP) values ( '"+ policyNumber +"','"+e2eTestCaseId +"','Selenium','"+dateToStr+ "') ";
			String query = "update TEST_EXEC_STATUS set EXEC_START_TMSP = '"+dateToStr+"', EXEC_THREAD_ID ='"+threadId+"', EXEC_MACH_NM = '"+ execMachineName+"', TEST_CASE_STAT_DESC ='In Progress' where TEST_CASE_ID = '"+testCaseID+"'";
			status = st.execute(query);
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			/*try {
			st.close();
			ucaConn.close();
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}*/
		}
		return status;
	
	}
	
	public boolean updateTestExecStatus(String testCaseID, String status){
		boolean isSucess = false;
		Date curDate = new Date();
		SimpleDateFormat format = new SimpleDateFormat("M/d/yyyy h:m:s aaa");
		String dateToStr = format.format(curDate);
		try {
			//String query =  "MERGE INTO TRANS bi USING (Select '"+ e2eTestCaseId +"' as E2E_TCASE_ID from dual) bo ON bi.E2E_TCASE_ID = bo.E2E_TCASE_ID" +
              //      " WHEN MATCHED THEN Update set POLICYNO = '"+ policyNumber +"', APPLICATION = 'Selenium', LAST_UPDATE_TIMESTAMP ='"+dateToStr +"', E2E_TCASE_ID = '"+e2eTestCaseId +"' WHEN NOT MATCHED THEN insert(POLICYNO, E2E_TCASE_ID, APPLICATION, LAST_UPDATE_TIMESTAMP) values ( '"+ policyNumber +"','"+e2eTestCaseId +"','Selenium','"+dateToStr+ "') ";
			String query = "update TEST_EXEC_STATUS set EXEC_END_TMSP = '"+dateToStr+"', TEST_CASE_STAT_DESC = '"+ status+"' where TEST_CASE_ID = '"+testCaseID+"'";
			isSucess = st.execute(query);
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			/*try {
			st.close();
			ucaConn.close();
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}*/
		}
		return isSucess;
	
	}
	
	public boolean populateDataFeed(String testCaseID, String valueToUpdate){
		boolean isSucess = false;
		Date curDate = new Date();
		SimpleDateFormat format = new SimpleDateFormat("M/d/yyyy h:m:s aaa");
		String dateToStr = format.format(curDate);
		
				
		try {
			
			//String query =  "MERGE INTO TRANS bi USING (Select '"+ e2eTestCaseId +"' as E2E_TCASE_ID from dual) bo ON bi.E2E_TCASE_ID = bo.E2E_TCASE_ID" +
              //      " WHEN MATCHED THEN Update set POLICYNO = '"+ policyNumber +"', APPLICATION = 'Selenium', LAST_UPDATE_TIMESTAMP ='"+dateToStr +"', E2E_TCASE_ID = '"+e2eTestCaseId +"' WHEN NOT MATCHED THEN insert(POLICYNO, E2E_TCASE_ID, APPLICATION, LAST_UPDATE_TIMESTAMP) values ( '"+ policyNumber +"','"+e2eTestCaseId +"','Selenium','"+dateToStr+ "') ";
			String query = null;
			if(valueToUpdate != null && valueToUpdate.length() >0
					&& valueToUpdate.equalsIgnoreCase("IN_PROGRESS")){
				query = "insert into DATA_FEED(TEST_CASE_ID,IN_PROGRESS,LAST_UPDT_TMSP) values('"+testCaseID+"','1','"+dateToStr+"')";
			} else if(valueToUpdate != null && valueToUpdate.length() >0
					&& valueToUpdate.equalsIgnoreCase("PASS")){
			
				query = "insert into DATA_FEED(TEST_CASE_ID,PASS,LAST_UPDT_TMSP) values('"+testCaseID+"','1','"+dateToStr+"')";
			}else if(valueToUpdate != null && valueToUpdate.length() >0
					&& valueToUpdate.equalsIgnoreCase("FAIL")){
				query = "insert into DATA_FEED(TEST_CASE_ID,FAIL,LAST_UPDT_TMSP) values('"+testCaseID+"','1','"+dateToStr+"')";
			}
			if(query != null){
				isSucess = st.execute(query);
			}
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			/*try {
			st.close();
			ucaConn.close();
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}*/
		}
		return isSucess;
	
	}
	
	public void closeConnections(){
		try {
			st.close();
			ucaConn.close();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
