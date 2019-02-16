package com.pc.driver;

import java.util.ArrayList;
import java.util.List;
import org.testng.TestNG;
import org.testng.xml.XmlClass;
import org.testng.xml.XmlSuite;
import org.testng.xml.XmlSuite.ParallelMode;
import org.testng.xml.XmlTest;

public class TestNGRunner {

	public static void main(String args[])
	{
	
		XmlSuite suite = new XmlSuite();
		suite.setName("TestAll");
		suite.setThreadCount(2);
		suite.setParallel(ParallelMode.METHODS);
		
		XmlTest test = new XmlTest(suite);
		test.setName("PC");
		
		test.addParameter("DataSheetName", "Data");
		test.addParameter("Region", "");
		List<XmlClass> classes = new ArrayList<XmlClass>();
		classes.add(new XmlClass("com.pc.driver.ParallelExecDriver"));
		//classes.add(new XmlClass("com.pc.driver.ParallelRegressionExecDriver"));
		test.setXmlClasses(classes) ;
		
		List<XmlSuite> suites = new ArrayList<XmlSuite>();
		suites.add(suite);
		TestNG tng = new TestNG();
		tng.setXmlSuites(suites);
		tng.run();
	} 
}
