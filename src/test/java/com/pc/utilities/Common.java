/**
 * @ClassPurpose This Class used to store the common methods across the project/package
 * @Scriptor Krishna Manubolu
 * @ReviewedBy
 * @ModifiedBy Sojan David
 * @LastDateModified 3/17/2017
 */
package com.pc.utilities;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.StringReader;
import java.io.StringWriter;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.NoSuchElementException;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.RandomStringUtils;
import org.apache.http.HttpResponse;
import org.apache.http.StatusLine;
import org.apache.http.client.ClientProtocolException;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.InputStreamEntity;
import org.apache.http.impl.client.HttpClientBuilder;
import org.apache.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.phantomjs.PhantomJSDriverService;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;

import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.google.common.base.Predicate;
import com.pc.constants.PCConstants;
import com.pc.elements.Elements;
import com.pc.screen.SCRCommon;

public class Common{

    public  XlsxReader sXL; // Excel Read Object
    public  Integer TCRow; // TestCase Sheet Row
    public  String  TCID; // Test Case ID
    public  String testcasename = null; // TestCase Name
    public  String methodName = null; // Component Name
    public  String  TestCaseID; //ALM TestID
    public  String  TestSetID; //ALM TestSetID
    public  String  DataSheetName; // DataSheet Name
    public  String UpdateID; //Write the excel sheet
    public  PhantomJSDriverService service = null; // Headless browser Varaible
    static  Logger logger =Logger.getLogger(Common.class); // Logger variable
    public  WebElement ele; // Safeaction Element Variabale
    public  static Elements o = new Elements(); //Object for element class
    private List<String> expectedXPaths = new ArrayList<>();

    
    /**
     * @function Constructor for common class
     */
    public Common()
    {
        
    }
    
    /**
     * @function This function use to wait until the next element to be exist
     * @param bylocator
     * @param iWaitTime
     * @return true/false
     * @throws Exception
     */
    public  boolean WaitForElementExist(By bylocator, int iWaitTime) throws Exception
    {
        boolean bFlag = false;
        WebDriverWait wait = new WebDriverWait(ManagerDriver.getInstance().getWebDriver(), iWaitTime);
        try
        {
            wait.until(ExpectedConditions.presenceOfElementLocated(bylocator)); //see if you can append ExpectedConditions.visibilityOfElementLocated(bylocator) also in Until
            if(ManagerDriver.getInstance().getWebDriver().findElement(bylocator).isDisplayed()||ManagerDriver.getInstance().getWebDriver().findElement(bylocator).isEnabled())
            {
                bFlag = true;
            }
        }
        catch (NoSuchElementException e)
        {
            e.printStackTrace();
            logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
            bFlag = false;
        }
        
        catch (Exception e)
        {
            e.printStackTrace();
            logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
            bFlag = false;
        }
        return bFlag;
    }
    
    /**
     * @function Safe Method for User Select option from list menu, waits until the element is loaded and then selects an option from list menu
     * @param bylocator
     * @param sOptionToSelect
     * @param iWaitTime
     * @return true/false
     * @throws Exception
    **/
    public  boolean SafeSelectGWListBox(By bylocator, String sOptionToSelect, int iWaitTime) throws Exception
    {
        WaitUntilClickable(bylocator, iWaitTime);
        WebElement element = ManagerDriver.getInstance().getWebDriver().findElement(bylocator);
        element.click();
        Thread.sleep(1000);
        ManagerDriver.getInstance().getWebDriver().findElement(bylocator).sendKeys(Keys.ARROW_DOWN);
        Thread.sleep(1500);
        boolean bFlag = false;
        WaitForElementExist(bylocator, iWaitTime);
        List<WebElement> gwListBox = ManagerDriver.getInstance().getWebDriver().findElements(By.tagName("LI"));
        for (int i=0; i<gwListBox.size(); i++)
        {
            String strListValue = gwListBox.get(i).getText();
            try
            {
                if (strListValue.contains(sOptionToSelect))
                {
                    System.out.println(gwListBox.get(i).getText());
                    gwListBox.get(i).click();
                    bFlag = true;
                    break;
                }
            }
            catch (Exception e)
            {
                e.printStackTrace();
                logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
                bFlag = false;
            }
        }
        return bFlag;
    }
    
    /**
     * @function This function use to wait untill the next element is ready to click
     * @param bylocator
     * @param iWaitTime
     * @return true/false
     * @throws Exception
     */
    public  boolean WaitUntilClickable(By bylocator, int iWaitTime) throws Exception
    {
        boolean bFlag = false;
        WebDriverWait wait = new WebDriverWait(ManagerDriver.getInstance().getWebDriver(), iWaitTime);
        try
        {
            wait.until(ExpectedConditions.elementToBeClickable(bylocator));
            //if(bylocator.isDisplayed())
            if(ManagerDriver.getInstance().getWebDriver().findElement((bylocator)).isDisplayed())
            {
                bFlag = true;
            }
        }
        catch (NoSuchElementException e)
        {
            e.printStackTrace();
            logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
            bFlag = false;
        }
        catch (Exception e)
        {
            e.printStackTrace();
            logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
            bFlag = false;
        }
        return bFlag;
    }
    
    /**
     * @function Highlights on current working element or locator
     * @param driver
     * @param locator
     * @throws Exception
     */
    public void highlightElement(By locator) throws Exception
    {
        //pro = new ConfigManager("sys");
        WebElement element = ManagerDriver.getInstance().getWebDriver().findElement(locator);
        if(HTML.properties.getProperty("HighlightElements").equalsIgnoreCase("true"))
        {
            String attributevalue="border:10px solid green;";//change border width and colour values if required
            JavascriptExecutor executor= (JavascriptExecutor) ManagerDriver.getInstance().getWebDriver();
            String getattrib=element.getAttribute("style");
            executor.executeScript("arguments[0].setAttribute('style', arguments[1]);", element, attributevalue);
            Thread.sleep(100);
            executor.executeScript("arguments[0].setAttribute('style', arguments[1]);", element, getattrib);
        }
    }
    
    /**
     * @function Use to perform any action in the application(click/edit/drop down/scroll)
     * @param element
     * @param value
     * @param ColumnName
     * @return true/false
     * @throws Exception
     */
    public Boolean SafeAction(By element, String value,String ColumnName) throws Exception
    {
        Boolean returnValue = true;
        Actions objActions = null;
        objActions = new Actions(ManagerDriver.getInstance().getWebDriver());
        JavascriptExecutor js = (JavascriptExecutor) ManagerDriver.getInstance().getWebDriver();
        String elementType = ColumnName.substring(0, 3);
        String objectName = ColumnName.substring(3);
        boolean elementClickable = WaitUntilClickable(element, Integer.valueOf(HTML.properties.getProperty("LONGESTWAIT")));
        if(elementClickable == true)
        {
                Boolean f = ManagerDriver.getInstance().getWebDriver().findElements(element).size()!=0;
                if (!f)
                {
                    returnValue = false;
                }
                else
                {
                    highlightElement(element);
                    try
                    {
                        ele = ManagerDriver.getInstance().getWebDriver().findElement(element);
                        returnValue = true;
                    }
                    catch(Exception e)
                    {
                        returnValue = false;
                    }
                }
        }
        else
        {
            returnValue = false;
        }
        if(returnValue)
        {
            switch (elementType.toUpperCase())
            {
                case "MEL":
                    String colName = ColumnName.toUpperCase();
                    Integer xYaxis=null;
                    Integer yYaxis=null;
                    if(colName.contains("ACCOUNT")){
                        xYaxis = 36;
                        yYaxis = 5;
                    }
                    else if(colName.contains("POLICY"))
                    {
                        xYaxis = 48;
                        yYaxis = 5;
                    }
                    else if(colName.contains("SEARCH"))
                    {
                        xYaxis = 60;
                        yYaxis = 5;
                    }else if(colName.contains("ADMINISTRATION"))
                    {
                        xYaxis = 67;
                        yYaxis = 5;
                    }else if(colName.contains("DESKTOP"))
                    {
                        xYaxis = 28;
                        yYaxis = 5;
                    }
                    Actions clickTriangle= new Actions(ManagerDriver.getInstance().getWebDriver());
                    clickTriangle.moveToElement(ele).moveByOffset(xYaxis, yYaxis).click().perform();
                    returnValue = SCRCommon.JavaScript(js);
                    logger.info("Thread ID = " + Thread.currentThread().getId() + " Clicked on '" + objectName + "' element or button or link and element '"+ element + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should click on '" + objectName + "' element or button or link", "Clicked on '" + objectName + "' element or button or link", "PASS");
                    break;
                case "ZED":
//                  ele.sendKeys(value);
//                  returnValue = true;
                    ManagerDriver.getInstance().getWebDriver().findElement(element).clear();
                    ManagerDriver.getInstance().getWebDriver().findElement(element).sendKeys(value);
                    WaitForPageToBeReady();
                    ManagerDriver.getInstance().getWebDriver().findElement(element).sendKeys(Keys.TAB);
                    WaitForPageToBeReady();
                    returnValue=SCRCommon.JavaScriptDynamicWait(ele, js);
                    logger.info("Thread ID = " + Thread.currentThread().getId() + " Value entered '"+ value + "' in '" + objectName + "' field and element '"+ element + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should enter value  '"+ value + "' in '" + objectName + "' field", "Value entered '"+ value + "' in '" + objectName + "' field", "PASS");
                    break;
                case "EDT":
                    ele.clear();
                    ele.sendKeys(value);
                    returnValue = SCRCommon.JavaScript(js);
                    logger.info("Thread ID = " + Thread.currentThread().getId() + " Value entered '"+ value + "' in '" + objectName + "' field and element '"+ element + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should enter value  '"+ value + "' in '" + objectName + "' field", "Value entered '"+ value + "' in '" + objectName + "' field", "PASS");
                    break;
                case "EDJ":
                    ele.clear();
                    ele.sendKeys(value);
                    ele.sendKeys(Keys.TAB);
                    logger.info("Thread ID = " + Thread.currentThread().getId() + " Value entered '"+ value + "' in '" + objectName + "' field and element '"+ element + "'");
//                  returnValue=SCRCommon.JavaScriptDynamicWait(ele, js);
                    returnValue = SCRCommon.JavaScript(js);
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should enter value  '"+ value + "' in '" + objectName + "' field", "Value entered '"+ value + "' in '" + objectName + "' field", "PASS");
                    break;
                case "PWD":
                    ele.clear();
                    ele.sendKeys(value);
                    returnValue = SCRCommon.JavaScript(js);
                    break;
                case "BTN":
                    ele.click();
                    returnValue=true;
                    logger.info("Thread ID = " + Thread.currentThread().getId() + "  Clicked on '" + objectName + "' element or button or link and element '"+ element + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should click on '" + objectName + "' element or button or link", "Clicked on '" + objectName + "' element or button or link", "PASS");
                    break;
                case "ELE":
                    Action objMouseClick1 = objActions
                            .click (ele)
                            .build();
                    objMouseClick1.perform();
                    returnValue = SCRCommon.JavaScript(js);
                    logger.info("Thread ID = " + Thread.currentThread().getId() + "  Clicked on '" + objectName + "' element or button or link and element '"+ element + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should click on '" + objectName + "' element or button or link", "Clicked on '" + objectName + "' element or button or link", "PASS");
                    break;
                case "ELJ":
                    Action objMouseClick2 = objActions
                                 .click (ele)
                                 .build();
                    objMouseClick2.perform();
                    logger.info("Thread ID = " + Thread.currentThread().getId() + " Clicked on '" + objectName + "' element or button or link and element '"+ element + "'");
//                    returnValue=SCRCommon.JavaScriptDynamicWait(ele, js);
                    returnValue = SCRCommon.JavaScript(js);
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should click on '" + objectName + "' element or button or link", "Clicked on '" + objectName + "' element or button or link", "PASS");
                    break;
                case "DBL":
                    objActions.click(ele);
                    Action objMousedblClick = objActions
                            .doubleClick (ele)
                            .build();
                    objMousedblClick.perform();
                    returnValue=true;
                    logger.info("Thread ID = " + Thread.currentThread().getId() + " Double Clicked on '" + objectName + "' element or button or link and element '"+ element + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should Double Click on '" + objectName + "' element or button or link", "Double Clicked on '" + objectName + "' element or button or link", "PASS");
                    break;
                case "LST":
                    //ManagerDriver.getInstance().getWebDriver().findElement(element).clear();
                    ManagerDriver.getInstance().getWebDriver().findElement(element).clear();
//                  Thread.sleep(1000);
                    ManagerDriver.getInstance().getWebDriver().findElement(element).sendKeys(value);
//                  WaitForPageToBeReady();
                    ManagerDriver.getInstance().getWebDriver().findElement(element).sendKeys(Keys.TAB);
//                  WaitForPageToBeReady();
                    returnValue = SCRCommon.JavaScript(js);
                    logger.info("Thread ID = " + Thread.currentThread().getId() + " Value available '"+ value +"' in '" + objectName + "' listbox and element '"+ element + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should Select item '"+ value +"' from '" + objectName + "' listbox", "Selected item '"+ value +"' from '" + objectName + "' listbox", "PASS");
                    break;
                case "LSJ":
                    ManagerDriver.getInstance().getWebDriver().findElement(element).clear();
//                  Thread.sleep(1000);
                    ManagerDriver.getInstance().getWebDriver().findElement(element).sendKeys(value);
                    ManagerDriver.getInstance().getWebDriver().findElement(element).sendKeys(Keys.TAB);
                    logger.info("Thread ID = " + Thread.currentThread().getId() + " Value available '"+ value +"' in '" + objectName + "' listbox and element '"+ element + "'");
//                  returnValue = SCRCommon.JavaScriptDynamicWait(ele, js);
                    returnValue = SCRCommon.JavaScript(js);
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should Select item '"+ value +"' from '" + objectName + "' listbox", "Selected item '"+ value +"' from '" + objectName + "' listbox", "PASS");
                    break;
                case "SCL":
                    ((JavascriptExecutor) ManagerDriver.getInstance().getWebDriver()).executeScript("arguments[0].scrollIntoView();",ManagerDriver.getInstance().getWebDriver().findElement(element));
                    logger.info("Thread ID = " + Thread.currentThread().getId() + " Scroll Donw to the Element " + objectName + " element or button or link and element '"+ element + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should click on '" + objectName + "' element or button or link", "Clicked on '" + objectName + "' element or button or link", "PASS");
                    returnValue = true;
                    break;
                case "RDO":
                    logger.info("Thread ID = " + Thread.currentThread().getId() + " Selected Radio " + objectName + " element or button or link and element '"+ element + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should click on '" + objectName + "' element or button or link", "Clicked on '" + objectName + "' element or button or link", "PASS");
                    returnValue = true;
                    break;
                case "CHK":
                    logger.info("Thread ID = " + Thread.currentThread().getId() + " Checked " + objectName + " element or button or link and element '"+ element + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should click on '" + objectName + "' element or button or link", "Clicked on '" + objectName + "' element or button or link", "PASS");
                    returnValue = true;
                    break;
                case "CLR":
                    logger.info("Thread ID = " + Thread.currentThread().getId() + " Checked " + objectName + " element or button or link and element '"+ element + "'");
                    ele.clear();
                    returnValue = SCRCommon.JavaScript(js);
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "", "", "PASS");
                    returnValue = true;
                    break;
                case "EJS":
                    logger.info("Thread ID = " + Thread.currentThread().getId() + " Checked " + objectName + " element or button or link and element '"+ element + "'");
                    JavascriptExecutor executor = (JavascriptExecutor)ManagerDriver.getInstance().getWebDriver();
                    executor.executeScript("arguments[0].click();", ele);
                    returnValue = SCRCommon.JavaScript(js);
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should click on '" + objectName + "' element or button or link", "Clicked on '" + objectName + "' element or button or link", "PASS");
                    break;
                case "BLI":
                    ele.click();
                    By option = By.xpath("//span[starts-with(text(),'"+value+"')]");
                        if(ManagerDriver.getInstance().getWebDriver().findElement(option).isDisplayed())
                        {
                            ManagerDriver.getInstance().getWebDriver().findElement(option).click();
                            returnValue = SCRCommon.JavaScript(js);
                            logger.info("Thread ID = " + Thread.currentThread().getId() + "  Clicked on '" + objectName + "' element or button or link and element '"+ element + "'");
                            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should click on '" + objectName + "' element or button or link", "Clicked on '" + objectName + "' element or button or link", "PASS");
                        }
                        else
                        {
                            returnValue=false;
                            logger.info("Thread ID = " + Thread.currentThread().getId() + " Object not enabled or displayed or not clickable '"+ element + "'");
                            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should display '" + objectName + "' in screen", "'" + objectName + "' not displayed in screen", "FAIL");
                        }
                    break;
                case "DRP":
                    Select sDropDown = new Select(ele);
                    sDropDown.selectByVisibleText(value);
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should select the '" + value + "' from the Dropdown", "Selected the '" + value + "' from the Dropdown", "PASS");
                    returnValue = true;
                    break;
                case "SEL":
                    boolean listStatus = false;
//                  WebElement element1 = DriverManager.getInstance().getWebDriver().findElement(element);
                    ele.click();
                    WaitForPageToBeReady();
                    ManagerDriver.getInstance().getWebDriver().findElement(element).sendKeys(Keys.ARROW_DOWN);
                    WaitForPageToBeReady();
                    List<WebElement> gwListBox = ManagerDriver.getInstance().getWebDriver().findElements(By.tagName("LI"));
                    for (int i = 0; i<gwListBox.size(); i++)
                    {
                        String strListValue = gwListBox.get(i).getText();
                        if (strListValue.toUpperCase().contains(value.toUpperCase()))
                        {
                            gwListBox.get(i).click();
                            returnValue = SCRCommon.JavaScript(js);
                            listStatus = true;
                            //logger.info("Selected item '"+ value +"' from '" + objectName + "' listbox and element '"+ element + "'");
                            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should Select item '"+ value +"' from '" + objectName + "' listbox", "Selected item '"+ value +"' from '" + objectName + "' listbox", "PASS");
                            break;
                        }
                    }
                    if(!listStatus)
                    {
                        returnValue = false;
                        //logger.info("Value not available '"+ value +"' in '" + objectName + "' listbox and element '"+ element + "'");
                        HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should Select item '"+ value +"' from '" + objectName + "' listbox", "Item '"+ value +"' not available in '" + objectName + "' listbox", "FAIL");
                        break;
                    }
                    break;
              /*  case "SLJ":
                    boolean listStatus1 = false;
//                  WebElement element1 = DriverManager.getInstance().getWebDriver().findElement(element);
                    ele.click();
                    WaitForPageToBeReady();
                    ManagerDriver.getInstance().getWebDriver().findElement(element).sendKeys(Keys.ARROW_DOWN);
                    WaitForPageToBeReady();
                    List<WebElement> gwListBox1 = ManagerDriver.getInstance().getWebDriver().findElements(By.tagName("LI"));
                    for (int i = 0; i<gwListBox1.size(); i++)
                    {
                        String strListValue = gwListBox1.get(i).getText();
                        if (strListValue.toUpperCase().contains(value.toUpperCase()))
                        {
                            gwListBox1.get(i).click();
                            returnValue = SCRCommon.JavaScript(js);
                            listStatus = true;
                            //logger.info("Selected item '"+ value +"' from '" + objectName + "' listbox and element '"+ element + "'");
                            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should Select item '"+ value +"' from '" + objectName + "' listbox", "Selected item '"+ value +"' from '" + objectName + "' listbox", "PASS");
                            break;
                        }
                    }
                    if(!listStatus1)
                    {
                        returnValue = false;
                        //logger.info("Value not available '"+ value +"' in '" + objectName + "' listbox and element '"+ element + "'");
                        HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should Select item '"+ value +"' from '" + objectName + "' listbox", "Item '"+ value +"' not available in '" + objectName + "' listbox", "FAIL");
                        break;
                    }
                    break;*/
            }
            WaitForPageToBeReady();
        }
        else
        {
            logger.info("Thread ID = " + Thread.currentThread().getId() + " Object not enabled or displayed or not clickable '"+ element + "'");
            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should display '" + objectName + "' in screen", "'" + objectName + "' not displayed in screen", "FAIL");
        }
        return returnValue;
    }
    
    /**
     * @function This will return the element
     * @param sLocator
     * @return
     */
    public  WebElement returnObject(By sLocator)
    {
        WebElement elements = null;
        try{
            elements = ManagerDriver.getInstance().getWebDriver().findElement(sLocator);
        }
        catch(Exception e)
        {
            e.printStackTrace();
            logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
        }
        return elements;
    }
    
    /**
     * function This will use to check the browser state is ready perform the next action
     */
    public void WaitForPageToBeReady()
    {
           //http://www.testingexcellence.com/webdriver-wait-page-load-example-java/
        JavascriptExecutor js = (JavascriptExecutor)ManagerDriver.getInstance().getWebDriver();
        for (int i=0; i<Integer.parseInt(HTML.properties.getProperty("VERYLONGWAIT")); i++)
        {
            try
            {
                Thread.sleep(1000);
            }catch (InterruptedException e) {}
            if (js.executeScript("return document.readyState").toString().equals("complete"))
            {
                break;
            }
          }
    }
    
    /**
     * Backup of the waitforpagetobeready function
     */
    public  void WaitForPageToBeReady1()
    {
        WebDriverWait wait = new WebDriverWait(ManagerDriver.getInstance().getWebDriver(), Integer.parseInt(HTML.properties.getProperty("VERYLONGWAIT")));
        Predicate<WebDriver> pageLoad = new Predicate<WebDriver>()
                {
                    @Override
                    public boolean apply(WebDriver input)
                    {
                        return((JavascriptExecutor) input).executeScript("return document.readyState").equals("complete");
                    }
                };
            System.out.println("Page is loaded");
            wait.until(pageLoad);
    }
    
    /**
     * @function Used to open the browser according to the environment variable
     * @return true/false
     * @throws Exception
     */
    public Boolean OpenApp() throws Exception
    {
            Boolean bStatus = false;
            String sURL = null;
            sURL = HTML.properties.getProperty(HTML.properties.getProperty("Region"));
            logger.info("Execution starting in '" + HTML.properties.getProperty("Region").toUpperCase() + "' environment");
            if(HTML.properties.getProperty("EXECUTIONMODE").equalsIgnoreCase("Remote"))
            {
                String sMachineAddress = "";
                //String sMachineAddress = RemoteDriverFactory.getInstance().hostname;
                HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "Machine IP Address = "+sMachineAddress+"","Machine IP Address = "+sMachineAddress+"", "PASS");
            }
            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "Execution should start in '" + HTML.properties.getProperty("Region").toUpperCase() + "' environment","Execution started in '" + HTML.properties.getProperty("Region").toUpperCase() + "' environment", "PASS");
            logger.info("Thread = "+Thread.currentThread().getId() +" Driver = "+ManagerDriver.getInstance().getWebDriver());
            ManagerDriver.getInstance().getWebDriver().manage().timeouts().implicitlyWait(25, TimeUnit.SECONDS);
            if(HTML.properties.getProperty("Browser").equalsIgnoreCase("CH") && HTML.properties.getProperty("TypeOfAutomation").equalsIgnoreCase("HEAD"))
            {
                  logger.info("Browser maximized");
//                ManagerDriver.getInstance().getWebDriver().manage().window().maximize();
            }else
            {
                  ManagerDriver.getInstance().getWebDriver().manage().window().maximize();
                  logger.info("Browser maximized");
            }
            ManagerDriver.getInstance().getWebDriver().get(sURL);
            logger.info("Execution starting in '" + HTML.properties.getProperty("Region").toUpperCase() + "' environment and url '" + sURL + "'");
            Integer x = Integer.valueOf(HTML.properties.getProperty("VERYLONGWAIT"));
//          if(CommonManager.getInstance().getCommon().WaitUntilClickable(o.getObject("edtUserName"),  x))
            if(WaitUntilClickable(o.getObject("eleDeskTopAction"),  x))
            {
                logger.info("Application Opened Successfully");
                HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "Application should Open successfully","Application Opened successfully", "PASS");
                bStatus = true;
            }
            return bStatus;
    }
    
    /**
     * Method - Safe Method for User Select option from list menu, waits until the element is loaded and then selects an option from list menu
     * @param bylocator
     * @param sOptionToSelect
     * @param iWaitTime
     * @return true/false
     * @throws Exception
     * @throws Exception
     */
    public  boolean MouseHoverAction(By sMainMenu, By sSubMenu)
    {
        boolean status = false;
        try
         {
             String mouseOverScript = "if(document.createEvent){var evObj = document.createEvent('MouseEvents');evObj.initEvent('mouseover',true, false); arguments[0].dispatchEvent(evObj);} else if(document.createEventObject) { arguments[0].fireEvent('onmouseover');}";
             ((JavascriptExecutor) ManagerDriver.getInstance().getWebDriver()).executeScript(mouseOverScript, ManagerDriver.getInstance().getWebDriver().findElement(sMainMenu));
             Thread.sleep(1000);
             ((JavascriptExecutor)ManagerDriver.getInstance().getWebDriver()).executeScript("arguments[0].click();",ManagerDriver.getInstance().getWebDriver().findElement(sSubMenu));
             status = true;
        }catch(Exception e)
        {
             System.out.println("Element not found");
             status = false;
        }
        return status;
    }
    
    /**
     * @function Ability to get the text of the element
     * @param bylocator
     * @param iWaitTime
     * @return element
     * @throws Exception
    **/
    public String ReadElement(By bylocator, int iWaitTime) throws Exception
    {
        WaitUntilClickable(bylocator, iWaitTime);
        WebElement element = ManagerDriver.getInstance().getWebDriver().findElement(bylocator);
        return element.getText();
    }
    
    /**
    * @function Ability to get the text of the element which is having Un clickable field
    * @param bylocator
    * @param iWaitTime
    * @return element
    * @throws Exception
    **/
    public  String ReadElementforODS(By bylocator, int iWaitTime)
    {
        WebElement element = ManagerDriver.getInstance().getWebDriver().findElement(bylocator);
//      String sElementText = element.getText();
        return element.getText();
    }
    
    /**
     * @function Ability to get the text of the element which is having Attribute value
     * @param bylocator
     * @param iWaitTime
     * @return element
     * @throws Exception
     **/
     public  String ReadElementGetAttribute(By bylocator, String sAttributeValue, int iWaitTime) throws Exception
     {
            WebElement element = ManagerDriver.getInstance().getWebDriver().findElement(bylocator);
            return element.getAttribute(sAttributeValue);
     }

    public boolean ElementExist(By bylocator) throws Exception
    {
        WebElement element = ManagerDriver.getInstance().getWebDriver().findElement(bylocator);
        if(element.isDisplayed())
        {
            return true;
        }else
        {
            return false;
        }
    }
    
    /**
     * @function This will use to check whether the object is empty or not
     * @return
     */
    public  boolean ElementEmpty(By sLocator)
    {
        ManagerDriver.getInstance().getWebDriver().findElements(sLocator).isEmpty();
        return true;
    }
    
    /**
     * @function Check whether the element is dispalyed or not
     * @param sLocator
     * @return true/false
     */
    public  boolean ElementDisplayed(By sLocator)
    {
        boolean status = false;
        if(ManagerDriver.getInstance().getWebDriver().findElement(sLocator).isDisplayed())
        {
            status = true;
        }else{
            status = false;
        }
        return status;
    }
    
    /**
     * @function Compare two strings and populate the results in HTML
     * @param sCase
     * @param sExpectedValue
     * @param sAcutualValue
     * @return true/false
     * @throws Exception
     */
    public boolean CompareStringResult(String sCase, String sExpectedValue, String sAcutualValue) throws Exception
    {
        boolean status = true;
        if(sAcutualValue.contains(sExpectedValue))
        {
            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), ""+sCase+": "+sExpectedValue+" should match", ""+sCase+": "+sAcutualValue+" is matching", "PASS");
            status = true;
        }else{
            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), ""+sCase+": "+sExpectedValue+" should match", ""+sCase+": "+sAcutualValue+" is not matching", "FAIL");
            status = false;
        }
        return status;
    }
    
    /**
     * @function This function is used to compare the actual result with the exact expected result
     * @param sCase
     * @param sExpectedValue
     * @param sAcutualValue
     * @return
     * @throws Exception
     */
    public boolean CompareExactStringResult(String sCase, String sExpectedValue, String sAcutualValue) throws Exception
    {
            boolean status = true;
            if(sExpectedValue.equals(sAcutualValue))
            {
                   logger.info("Expected error text is matching with Actual Message Actual String:::'"+sAcutualValue+"' Expected String:::'" + sExpectedValue + "'");
                   HTML.fnInsertResult(HTML.properties.getProperty("testcasename"), HTML.properties.getProperty("methodName"), ""+sCase+": "+sExpectedValue+" should match", ""+sCase+": "+sAcutualValue+" is matching", "PASS");
                   status = true;
            }
            else
            {
                   logger.info("Expected error text is not matching with Actual Message Actual String:::'"+sAcutualValue+"' Expected String:::'" + sExpectedValue + "'");
                   HTML.fnInsertResult(HTML.properties.getProperty("testcasename"), HTML.properties.getProperty("methodName"), ""+sCase+": "+sExpectedValue+" should match", ""+sCase+": "+sAcutualValue+" is not matching", "FAIL");
                   status = false;
            }
            return status;
    }
    
    /**
     * @function Compare two strings and populate the results in HTML
     * @param sCase
     * @param sExpectedValue
     * @param sAcutualValue
     * @return true/false
     * @throws Exception
     */
     public boolean SpecialCompareResult(String sCase, String sExpectedValue, String sAcutualValue) throws Exception
     {
            boolean status = true;
            if(sExpectedValue.contains(sAcutualValue))
            {
            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), ""+sCase+": "+sExpectedValue+" should match", ""+sCase+": "+sAcutualValue+" is matching", "PASS");
                   status = true;
            }else{
            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), ""+sCase+": "+sExpectedValue+" should match", ""+sCase+": "+sAcutualValue+" is not matching", "FAIL");
                   status = false;
            }
            return status;
     }

    
    /**
     * @function Check whether the element is present in the applicaiton(element should not present) and populate the results
     * @param sCase
     * @param sExpectedValue
     * @param sAcutualValue
     * @param sId
     * @return staus
     * @throws Exception
     */
    public  boolean ElementExistOrNotFalse(String sCase, String sExpectedValue, String sAcutualValue, By sId) throws Exception
    {
        boolean status = ManagerDriver.getInstance().getWebDriver().findElements(sId).size()!=0;
        if(status == false)
        {
        HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), ""+sCase+": '"+sExpectedValue+"' should not present", ""+sCase+": '"+sAcutualValue+"' not present", "PASS");
            status = true;
        }else{
        HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), ""+sCase+": '"+sExpectedValue+"' should not present", ""+sCase+": '"+sAcutualValue+"' is present", "FAIL");
            status = false;
        }
        return status;
    }
    
    /**
     * @function Check whether the element is present in the applicaiton(element should present) and populate the results
     * @param sCase
     * @param sExpectedValue
     * @param sAcutualValue
     * @param sId
     * @return true/false
     * @throws Exception
     */
    public  boolean ElementExistOrNotTrue(String sCase, String sExpectedValue, String sAcutualValue, By sId) throws Exception
    {
        boolean status = ManagerDriver.getInstance().getWebDriver().findElements(sId).size()!=0;
        if(status == true)
        {
            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), ""+sCase+": '"+sExpectedValue+"'", ""+sCase+": '"+sAcutualValue+"'", "PASS");
            status = true;
        }else{
            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), ""+sCase+": '"+sExpectedValue+"'", ""+sCase+": '"+sAcutualValue+"' is not present", "FAIL");
            status = false;
        }
        return status;
    }
    
    /**
     * @function this function will quite the browser
     * @throws Exception
     */
    public void Terminate() throws Exception
    {
        String execution = HTML.properties.getProperty("TypeOfAutomation");
        if (execution.toUpperCase().contains("HEADLESS"))
        {
//          service.stop();
            if(ManagerPhantomJS.getInstance() != null  && ManagerPhantomJS.getInstance().getPhantomJSDrivrService() != null)
                ManagerPhantomJS.getInstance().getPhantomJSDrivrService().stop();
            logger.info("phantomjs service stoped");
        }
        else
        {
            if(ManagerDriver.getInstance() != null && ManagerDriver.getInstance().getWebDriver() != null)
                ManagerDriver.getInstance().getWebDriver().quit();
            logger.info("WebDriver Quit");
        }
        ScreenVideoCapture.stopVideoCapture(HTML.properties.getProperty("testcasename"));
    }
    
    /**
     * @function this function will quite the browser
     * @throws Exception
     *//*
    public void TerminateFail() throws Exception
    {
        String execution = HTML.properties.getProperty("TypeOfAutomation");
        if (execution.toUpperCase().contains("HEADLESS"))
        {
            service.stop();
            logger.info("phantomjs service stoped");
        }
        else
        {
            try{
                Alert al = ManagerDriver.getInstance().getWebDriver().switchTo().alert();
                al.dismiss();
            }
            catch(Exception e)
            {
                logger.info("Alert not found");
            }
            ManagerDriver.getInstance().getWebDriver().quit();
            logger.info("WebDriver Quit");
        }
        ScreenVideoCapture.stopVideoCapture(HTML.properties.getProperty("testcasename"));
    }*/
    
    
    /**
     * @function This function used to take the SS where its required
     * @throws IOException
     */
    public  void TerminationScreenShot() throws IOException
    {
           File reportFile;
           int number = 0;
           Date currDate = new Date();
           SimpleDateFormat dateFormat = new SimpleDateFormat("dd_MM_yyyy");
           String date = dateFormat.format(currDate);
           do {
                        reportFile = new File("Reports\\HTMLReports\\ScreenShot" + date +"_"+ number + ".png");
                        number++;
              }
                  while (reportFile.exists());
                  File screenshot = ((TakesScreenshot) ManagerDriver.getInstance().getWebDriver()).getScreenshotAs(OutputType.FILE);
                  FileUtils.copyFile(screenshot, reportFile);
    }
    
    /**
     *
     * @param intTestCaseID,intTestSetID,FLAG_EVALFAIL,strAttachmentFilePath,strAttachmentDesc,strUserName,strPassword,sQCURL,sQCURL,sProject
     * @function This function use to update the test results in ALM
     * @throws IOException
     */
    public void RunScript(String intTestCaseID, String intTestSetID, String FLAG_EVALFAIL, String strAttachmentFilePath, String strAttachmentDesc,String strUserName,String strPassword, String sQCURL, String sDomain, String sProject, String sDraftRun) throws IOException{
        //http://stackoverflow.com/questions/14711490/pass-arguments-to-vbs-from-java
        File directory = new File (".");
        String sConfigfilespath = directory.getCanonicalPath()+"\\VBScript\\UpdateALM.vbs";
            try
            {
                String[] parms = {"wscript", sConfigfilespath, intTestCaseID, intTestSetID, FLAG_EVALFAIL, strAttachmentFilePath, strAttachmentDesc,strUserName,strPassword,sQCURL,sDomain,sProject,sDraftRun};
//              Runtime.getRuntime().exec(parms);
                Process p = Runtime.getRuntime().exec(parms);
                if(!p.waitFor(2, TimeUnit.MINUTES)){
                       logger.info("Timed Out while sending email");
                       p.destroy();
                }

            } catch (IOException | InterruptedException e)
            {
                e.printStackTrace();
            }
    }
    
    /**
     * @function This function use to send a mail after the test with the results
     * @throws IOException
     */
    public void SendMail(String strMailTo,String strMailCC,String strSummaryFileName,String g_tSummaryEnd_Time,String g_tSummaryStart_Time, String strRelease,String strModuleName,String g_SummaryTotal_TC,String g_SummaryTotal_Pass,String g_SummaryTotal_Fail,String strEnvironment) throws IOException{
        //http://stackoverflow.com/questions/14711490/pass-arguments-to-vbs-from-java
        File directory = new File (".");
        String sConfigfilespath = directory.getCanonicalPath()+"\\VBScript\\SendMail.vbs";
            try
            {
                String[] parms = {"wscript", sConfigfilespath, strMailTo,strMailCC,strSummaryFileName,g_tSummaryEnd_Time,g_tSummaryStart_Time,strRelease,strModuleName, g_SummaryTotal_TC,g_SummaryTotal_Pass,g_SummaryTotal_Fail,strEnvironment};
//              Runtime.getRuntime().exec(parms);
                Process p = Runtime.getRuntime().exec(parms);
                if(!p.waitFor(2, TimeUnit.MINUTES)){
                       logger.info("Timed Out while sending email");
                       p.destroy();
                }
            } catch (IOException | InterruptedException e)
            {
                e.printStackTrace();
                logger.error("Thread ID = " + Thread.currentThread().getId() + "Error Occured =" +e.getMessage(), e);
            }
    }
    
    /**
     * @function this function use to find the element size
     * @param byLocater
     * @return true/false
     */
    public int ElementSize(By byLocater)
    {
        //boolean status = false;
        int size = ManagerDriver.getInstance().getWebDriver().findElements(byLocater).size();
        return size;
    }
    
    /**
     * @function This function use to update the data in the excel sheet
     * @param sFilename
     * @param sQuery
     * @return true/false
     * @throws Exception
     */
    public synchronized boolean UpdateQueryDeprecatedDoNotUse(String sFilename, String sQuery) throws Exception
    //public boolean UpdateQueryDeprecated(String sFilename, String sQuery) throws Exception
    {
        boolean status = false;
        try
        {
            Fillo fillo=new Fillo();
            File directory = new File (".");
            String sConfigfilespath = directory.getCanonicalPath()+"\\Data\\"+sFilename+".xlsm";
            Connection connection=fillo.getConnection(sConfigfilespath);
            String strQuery= sQuery;
            connection.executeUpdate(strQuery);
            connection.close();
            status = true;
        }
        catch(Exception e)
        {
            e.printStackTrace();
            logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
            status = false;
        }
        return status;
    }
    
    /**
     * @function Common function for all the screen class
     * @param Sheetname
     * @param o
     * @return true/false
     */
    public  Boolean ClassComponent(String Sheetname, Elements o)
    {
        //System.out.println("ClassComponent  Started = " + Thread.currentThread().getId() +" Driver = " + DriverManager.getInstance().getWebDriver());
        XlsxReader sXL;
        boolean tcAvailability = true;
        String sheetname = Sheetname;
        PCThreadCache.getInstance().setProperty(PCConstants.componentSheet, Sheetname);
        //sXL = new XlsxReader(   HTML.properties.getProperty("DataSheetName"));
        sXL = XlsxReader.getInstance();// new XlsxReader(   PCThreadCache.getInstance().getProperty("DataSheetName"));
        Boolean status = true;
        try
        {
            int rowcount = sXL.getRowCount(sheetname);
              for(int i = 2; i <= rowcount; i++)
              {
                 //if(sXL.getCellData(sheetname, "ID", i).equals(PCThreadCache.getInstance().getProperty("TCID")))
                  if(sXL.getCellData(sheetname, "ID", i).equals(PCThreadCache.getInstance().getProperty("TCID")))
                  {
                      tcAvailability = false;
                        int colcount = sXL.getColumnCount(sheetname);
                        for(int j = 2; j <= colcount; j++)
                        {
                                String ColName = sXL.getCellData(sheetname, j, 1);
                                if (!ColName.isEmpty())
                                {
                                        String value = sXL.getCellData(sheetname, j, i);
                                        String element = ColName.substring(0, 3);
                                        String sIteration = sXL.getCellData(sheetname, 1, i);
                                        PCThreadCache.getInstance().setProperty(PCConstants.Iteration, sIteration);
                                        if (element.contentEquals("mel") || element.contentEquals("fun") || element.contentEquals("cfu") || element.contentEquals("zed") || element.contentEquals("edt") || element.contentEquals("btn")  || element.contentEquals("ele") || element.contentEquals("lst") || element.contentEquals("pwd") || element.contentEquals("dbl") || element.contentEquals("scl") || element.contentEquals("rdo") || element.contentEquals("chk") || element.contentEquals("clr") || element.contentEquals("edj") || element.contentEquals("elj")  || element.contentEquals("ofu") || element.contentEquals("edw") || element.contentEquals("bli") || (element.contentEquals("lsj") || (element.contentEquals("drp") || element.contentEquals("sel") || element.contentEquals("val"))))
                                        {
                                            if ((!value.equals("")))
                                            {
                                                String ClassName  =null;
                                                if(element.toUpperCase().contains("FUN") || element.toUpperCase().contains("OFU") )
                                                {
                                                    ClassName  = "com.pc.screen." + sheetname;
                                                }
                                                if(element.toUpperCase().contains("CFU") || (element.toUpperCase().contains("VAL")))
                                                {
                                                    ClassName  = "com.pc.screen." + "SCRCommon";
                                                }
                                                if(element.toUpperCase().contains("FUN") || element.toUpperCase().contains("CFU") || element.toUpperCase().contains("OFU"))
                                                {
                                                        String methodName = ColName.substring(3);
                                                        if(value.toUpperCase().equals("YES"))
                                                        {
                                                            Class noparams[] = {};
                                                            Class cls = Class.forName(ClassName);
                                                            Object obj = cls.newInstance();
                                                            Method method = cls.getDeclaredMethod(methodName, noparams);
                                                            status = (Boolean)method.invoke(obj, null);
                                                        }
                                                        else
                                                        {
                                                            if(ColName.toUpperCase().endsWith("PAGE"))
                                                            {
                                                                   methodName = "ODSCfun";
                                                            }
                                                            Class[] paramString = new Class[1];
                                                            paramString[0] = String.class;
                                                            Class cls = Class.forName(ClassName);
                                                            Object obj = cls.newInstance();
                                                            Method method = cls.getDeclaredMethod(methodName, paramString);
                                                            status = (Boolean)method.invoke(obj, new String(value));
                                                        }
                                                }
                                                else if(element.toUpperCase().contains("VAL"))
                                                {
                                                    String methodName = "IconValidation";
                                                    Class[] paramString = new Class[2];
                                                    paramString[0] = String.class;
                                                    paramString[1] = String.class;
                                                    Class cls = Class.forName(ClassName);
                                                    Object obj = cls.newInstance();
//                                                  Method method = cls.getDeclaredMethod(methodName, new Class[]{String.class,String.class});
                                                    Method method = cls.getDeclaredMethod(methodName, paramString);
                                                    status = (Boolean)method.invoke(obj, new String(ColName),new String(value));
                                                }
                                                else
                                                {
                                                        status = SafeAction(o.getObject(ColName), value,ColName);
                                                }
                                                if(!status)
                                                {
                                                    status = handleUnknownAlert();
                                                    return false;
                                                }
                                            }
                                        }
                                }
                        }
                  }
              }
              if(tcAvailability)
              {
                  logger.info(PCThreadCache.getInstance().getProperty("TCID") + ":::"+PCThreadCache.getInstance().getProperty("TCID")+" not available in "+Sheetname+" Sheet");
                  status = false;
              }
        }
        catch(Exception e)
        {
            e.printStackTrace();
            logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
            status = false;
        }
        return status;
    }
    
        
    public String getElementByXPath(Document document,
            String path) throws Exception {
        XPath xPath =  XPathFactory.newInstance().newXPath();
        String value = xPath.compile(path).evaluate(document);
        return value;
         
    }
     
    public Document setElementByXPath(Document document,
            String path,String newContent) throws Exception {
        XPath xPath =  XPathFactory.newInstance().newXPath();
        Node n = (Node) xPath.compile(path).evaluate(document, XPathConstants.NODE);
        n.setTextContent(newContent);
        return document;
         
    }

    

    public Document setNodeValueByElementName(Document document,
            String elementName, String value) {
        Node elementByTag = getElementByTagName(document, elementName);
        if (elementByTag != null) {
            elementByTag.setTextContent(value);
        }
        return document;
    }

    public String getNodeValueByElementName(Document document,
            String elementName) {
        Node elementByTag = getElementByTagName(document, elementName);
        if (elementByTag != null) {
            return elementByTag.getTextContent();
        }
        return null;
    }

    public static org.w3c.dom.Node getElementByTagName(Document document,
            String elementName) {
        NodeList nodeList = document.getElementsByTagName("*");
        for (int i = 0; i < nodeList.getLength(); i++) {
            org.w3c.dom.Node node = nodeList.item(i);
            System.out.println(node.getNodeName());
            if (elementName.equalsIgnoreCase(node.getNodeName())) {
                return node;
            }
        }
        return null;
    }

    public  void changeNodeText(Node context, String xpath, String value)throws XPathExpressionException
            {
                XPathFactory xFactory = XPathFactory.newInstance();
                XPath xPath = xFactory.newXPath();
                XPathExpression expression = xPath.compile(xpath);
                NodeList nodes = (NodeList)expression.evaluate(context, XPathConstants.NODESET);
                for (int i = 0; i < nodes.getLength(); i++)
                {
                    Node node = nodes.item(i);
                    node.setTextContent(value);
                }
            }
    
    public Boolean coverageFunctionCall(String strFunctionName, String strFunctionValue)
    {
        Boolean blnStatus = false;
        String ClassName  =null;
        String strMethodName = strFunctionName.substring(3);
        String strFunctionType = strFunctionName.substring(0, 3);
        try{
            if(strFunctionType.toUpperCase().contains("FUN"))
            {
                ClassName  = "com.pc.screen." + PCThreadCache.getInstance().getProperty(PCConstants.componentSheet);
                if(strFunctionValue.toUpperCase().equals("YES"))
                {
                    Class noparams[] = {};
                    Class cls = Class.forName(ClassName);
                    Object obj = cls.newInstance();
                    Method method = cls.getDeclaredMethod(strMethodName, noparams);
                    blnStatus = (Boolean)method.invoke(obj, null);
                }
                else
                {
                    Class[] paramString = new Class[1];
                    paramString[0] = String.class;
                    Class cls = Class.forName(ClassName);
                    Object obj = cls.newInstance();
                    Method method = cls.getDeclaredMethod(strMethodName, paramString);
                    blnStatus = (Boolean)method.invoke(obj, new String(strFunctionValue));
                }
            }else if (strFunctionType.toUpperCase().contains("CFU"))
            {
                ClassName  = "com.pc.screen." + "SCRCommon";
                if(strFunctionValue.toUpperCase().equals("YES"))
                {
                    Class noparams[] = {};
                    Class cls = Class.forName(ClassName);
                    Object obj = cls.newInstance();
                    Method method = cls.getDeclaredMethod(strMethodName, noparams);
                    blnStatus = (Boolean)method.invoke(obj, null);
                }
                else
                {
                    Class[] paramString = new Class[1];
                    paramString[0] = String.class;
                    Class cls = Class.forName(ClassName);
                    Object obj = cls.newInstance();
                    Method method = cls.getDeclaredMethod(strMethodName, paramString);
                    blnStatus = (Boolean)method.invoke(obj, new String(strFunctionValue));
                }
            }
        }catch(Exception e)
        {
            e.printStackTrace();
        }finally{
            ClassName  =null;
            strMethodName = null;
            strFunctionType = null;
        }
        return blnStatus;
    }
    
    /**
     * @function This function use to Select the data from the table and click the element accordingly
     * @param obj,readCol,actionCol,strReadString,actionObjetName
     * @return status
     * @throws Exception
     */
    public Boolean ActionOnTable(By obj, int readCol, int actionCol, String strReadString, String actionObjetName, String sTagName) throws Exception
    {
      boolean Status=false;
      boolean SearchString=false;
      boolean ActionObject=false;
      WebElement mytable = ManagerDriver.getInstance().getWebDriver().findElement(obj);
      List<WebElement> allrows = mytable.findElements(By.tagName("tr"));
      for(int i = 0; i <= allrows.size()-1; i++)
      {
          List<WebElement> Cells = allrows.get(i).findElements(By.tagName("td"));
          String readText = Cells.get(readCol).getText();
          if (readText.contains(strReadString))
          {
              SearchString = true;
              List<WebElement> CellElements = Cells.get(actionCol).findElements(By.tagName(sTagName));
              for(WebElement element: CellElements)
              {
                  String objName = element.getText();
                  if(objName.contains(actionObjetName))
                    {
                        Status = true;
                        ActionObject = true;
                        element.click();
                        break;
                    }
              }
         }
         if(ActionObject == true)
         {
             break;
         }
      }
      if(SearchString)
      {
            logger.info("Search String available in the table. '" + strReadString + "'");
            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and srarch string is  '" + actionObjetName + "'", "PASS");
            if(ActionObject)
              {
                    logger.info("Search and click on object in the table cell and object name is '" + actionObjetName + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search object in the table cell and click on object. Object name is '" + actionObjetName + "'","System searched object in the table and clicked on object. object name is '" + actionObjetName + "'", "PASS");
                    Status = true;
              }
            else
              {
                    logger.info("Search and click on object in the table cell and object name is '" + actionObjetName + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search object in the table cell and click on object. Object name is '" + actionObjetName + "'","System searched object in the table and clicked on object. object name is '" + actionObjetName + "'", "FAIL");
                    Status = false;
              }
      }
      else
      {
            logger.info("Search String not available in the table. '" + strReadString + "'");
            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and srarch string is  '" + actionObjetName + "'", "FAIL");
            Status = false;
      }
      return Status;
    }
    
    /**
     * @function This function use to Select the data from the table and click the element accordingly
     * @param obj,readCol,actionCol,strReadString,actionObjetName
     * @return status
     * @throws Exception
     */
     public Boolean ActionOnTable(By obj, int readCol, int actionCol, String strReadString, String sTagName) throws Exception
     {
       boolean Status=false;
       boolean SearchString=false;
       boolean ActionObject=false;
       JavascriptExecutor js = (JavascriptExecutor) ManagerDriver.getInstance().getWebDriver();
       WebElement mytable = ManagerDriver.getInstance().getWebDriver().findElement(obj);
       List<WebElement> allrows = mytable.findElements(By.tagName("tr"));
       for(int i = 0; i <= allrows.size()-1; i++)
       {
              List<WebElement> Cells = allrows.get(i).findElements(By.tagName("td"));
              String readText = Cells.get(readCol).getText();
              if (readText.contains(strReadString))
              {
                     SearchString = true;
                     List<WebElement> CellElements = Cells.get(actionCol).findElements(By.tagName(sTagName));
                       for(WebElement element: CellElements)
                       {
        //                   String objName = element.getText();
        //                   if(objName.contains(actionObjetName))
        //                        {
//                               WebElement sElement = returnObject(By.id(readAttriID1));
                                 Status = true;
                                 ActionObject = true;
                                 element.click();
                                Status = SCRCommon.JavaScript(js);
                                 break;
//                 }
                       }
          }
            if(ActionObject == true)
            {
                   break;
            }
       }
       if(SearchString)
       {
                   logger.info("Search String available in the table. '" + strReadString + "'");
                   HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and srarch string is  '" + strReadString + "'", "PASS");
                   if(ActionObject)
                     {
                                logger.info("Search and click on object in the table cell and object name is '" + strReadString + "'");
                                HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search object in the table cell and click on object. Object name is '" + strReadString + "'","System searched object in the table and clicked on object. object name is '" + strReadString + "'", "PASS");
                                Status = true;
                     }
                   else
                     {
                                logger.info("Search and click on object in the table cell and object name is '" + strReadString + "'");
                                HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search object in the table cell and click on object. Object name is '" + strReadString + "'","System searched object in the table and clicked on object. object name is '" + strReadString + "'", "FAIL");
                                Status = false;
                     }
       }
       else
       {
                   logger.info("Search String not available in the table. '" + strReadString + "'");
                   HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and srarch string is  '" + strReadString + "'", "FAIL");
                   Status = false;
       }
       return Status;
     }
     /**
      * @function This function use to Select the data from the table and click the element accordingly
      * @param obj,readCol,actionCol,strReadString,actionObjetName
      * @return status
      * @throws Exception
      */
      public Boolean ActionOnTable_JS(By obj, int readCol, int actionCol, String strReadString, String sTagName) throws Exception
      {
        boolean Status=false;
        boolean SearchString=false;
        boolean ActionObject=false;
        WebElement mytable = ManagerDriver.getInstance().getWebDriver().findElement(obj);
        List<WebElement> allrows = mytable.findElements(By.tagName("tr"));
        for(int i = 0; i <= allrows.size()-1; i++)
        {
               List<WebElement> Cells = allrows.get(i).findElements(By.tagName("td"));
               String readText = Cells.get(readCol).getText();
               if (readText.contains(strReadString))
               {
                      SearchString = true;
                      List<WebElement> CellElements = Cells.get(actionCol).findElements(By.tagName(sTagName));
                       for(WebElement element: CellElements)
                       {
//
                                String readAttriID1 = allrows.get(i).getAttribute("id"); //clcikg on specifc row
                                WebElement sElement = returnObject(By.id(readAttriID1));
                                //sElement.click();
                                Status = SafeAction(By.id(readAttriID1), "elj", "elj");
                                ActionObject = true;
                                //status = true;
                                break;

                       }
           }
             if(ActionObject == true)
             {
                    break;
             }
        }
        if(SearchString)
        {
                    logger.info("Search String available in the table. '" + strReadString + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and srarch string is  '" + strReadString + "'", "PASS");
                    if(ActionObject)
                      {
                                 logger.info("Search and click on object in the table cell and object name is '" + strReadString + "'");
                                 HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search object in the table cell and click on object. Object name is '" + strReadString + "'","System searched object in the table and clicked on object. object name is '" + strReadString + "'", "PASS");
                                 Status = true;
                      }
                     else
                      {
                                 logger.info("Search and click on object in the table cell and object name is '" + strReadString + "'");
                                 HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search object in the table cell and click on object. Object name is '" + strReadString + "'","System searched object in the table and clicked on object. object name is '" + strReadString + "'", "FAIL");
                                 Status = false;
                      }
        }
        else
        {
                    logger.info("Search String not available in the table. '" + strReadString + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and srarch string is  '" + strReadString + "'", "FAIL");
                    Status = false;
        }
        return Status;
      }
    
    /**
     * @function this function use to get the text from the table according to the input and the column
     * @param obj,readTextCol,getTextCol,strReadString
     * @return String
     * @throws Exception
     */
    public  String GetTextFromTable1(By obj, int readTextCol, int getTextCol, String strReadString) throws Exception
    {
          String text = null;
          WebElement mytable = ManagerDriver.getInstance().getWebDriver().findElement(obj);
          List<WebElement> allrows = mytable.findElements(By.tagName("tr"));
          for(int i = 0; i <= allrows.size()-1; i++)
          {
              List<WebElement> Cells = allrows.get(i).findElements(By.tagName("td"));
              String readText = Cells.get(readTextCol).getText();
              if (readText.contains(strReadString))
              {
                  text = Cells.get(getTextCol).getText();
                  break;
              }
           }
          return text;
    }
    
    /**
     * @function this function use to get the text from the table according to the input and the column
     * @param obj,readTextCol,getTextCol,strReadString
     * @return String
     * @throws Exception
     */
     public String GetTextFromTable(By obj, int readTextCol, int getTextCol, String strReadString) throws Exception
     {
              String text = null;
              boolean SearchString = false;
              WebElement mytable = ManagerDriver.getInstance().getWebDriver().findElement(obj);
              List<WebElement> allrows = mytable.findElements(By.tagName("tr"));
              for(int i = 0; i <= allrows.size()-1; i++)
              {
                     List<WebElement> Cells = allrows.get(i).findElements(By.tagName("td"));
                     String readText = Cells.get(readTextCol).getText();
                     if (readText.contains(strReadString))
                     {
                           SearchString = true;
                           text = Cells.get(getTextCol).getText();
                           break;
                     }
               }
              if(SearchString)
              {
                          logger.info("Search String available in the table. '" + strReadString + "'");
                          HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and search string is  '" + strReadString + "'", "PASS");
              }
              else
              {
                         logger.info("Search String not available in the table. '" + strReadString + "'");
                         HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and search string is  '" + strReadString + "'", "FAIL");
                         SearchString = false;
              }
              return text;
     }
     
    /**
     * @function this function use to get the text the table
     * @param obj,getTextRow,getTextRow
     * @return String
     * @throws Exception
     */
    public  String GetTextFromTable(By obj, int getTextRow, int getTextCol) throws Exception
    {
          String text = null;
          WebElement mytable = ManagerDriver.getInstance().getWebDriver().findElement(obj);
          List<WebElement> allrows = mytable.findElements(By.tagName("tr"));
          List<WebElement> Cells = allrows.get(getTextRow).findElements(By.tagName("td"));
          text = Cells.get(getTextCol).getText();
          return text;
    }
    
    /**
     * @function This function use to get the text from the table according to the column
     * @param obj
     * @param readTextCol
     * @param strReadString
     * @return String
     * @throws Exception
     */
    public  String GetTextFromTable(By obj, int readTextCol, String strReadString) throws Exception
    {
          boolean SearchString = false;
          String readText = null;
          WebElement mytable = ManagerDriver.getInstance().getWebDriver().findElement(obj);
          List<WebElement> allrows = mytable.findElements(By.tagName("tr"));
          for(int i = 0; i <= allrows.size()-1; i++)
          {
              List<WebElement> Cells = allrows.get(i).findElements(By.tagName("td"));
              readText = Cells.get(readTextCol).getText();
              if(readText.contains(strReadString))
              {
                  SearchString = true;
                  break;
              }
           }
          if(SearchString)
          {
                //logger.info("Search String available in the table. '" + strReadString + "'");
              HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and search string is  '" + strReadString + "'", "PASS");
          }
          else
          {
                //logger.info("Search String not available in the table. '" + strReadString + "'");
            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and search string is  '" + strReadString + "'", "FAIL");
                SearchString = false;
          }
          return readText;
    }
    
    /**
     * @function This function use to get the text from the table according to the column
     * @param obj
     * @param readTextCol
     * @param strReadString
     * @return String
     * @throws Exception
     */
    public  String GetTextFromTable(By obj, int readTextCol, int getTextCol, String strReadTextString, String strGetTextString) throws Exception
    {
          boolean searchString = false;
          boolean readString = false;
          String readText = null;
          String getText = null;
          WebElement mytable = ManagerDriver.getInstance().getWebDriver().findElement(obj);
          List<WebElement> allrows = mytable.findElements(By.tagName("tr"));
          for(int i = 0; i <= allrows.size()-1; i++)
          {
              List<WebElement> Cells = allrows.get(i).findElements(By.tagName("td"));
              readText = Cells.get(readTextCol).getText();
              if(readText.contains(strReadTextString))
                  {
                      readString = true;
                      getText = Cells.get(getTextCol).getText();
                          if(getText.contains(strGetTextString))
                          {
                              searchString = true;
                              break;
                          }
                  }
           }
          if(searchString)
          {
                //logger.info("Search String available in the table. '" + strReadString + "'");
              HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadTextString + "'","System searched string in table and search string is  '" + strReadTextString + "'", "PASS");
                  if(readString)
                  {
                        //logger.info("Search String available in the table. '" + strReadString + "'");
                      HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strGetTextString + "'","System searched string in table and search string is  '" + strGetTextString + "'", "PASS");
                  }
                  else
                  {
                        //logger.info("Search String not available in the table. '" + strReadString + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strGetTextString + "'","System searched string in table and search string is  '" + strGetTextString + "'", "FAIL");
                    searchString = false;
                  }
          }
          else
          {
                //logger.info("Search String not available in the table. '" + strReadString + "'");
            HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadTextString + "'","System searched string in table and search string is  '" + strReadTextString + "'", "FAIL");
            searchString = false;
          }
          return getText;
    }
    
    /**
     * @function This function use to start the driver script
     * @param strRunMode
     * @param strTestCaseName
     * @param DataSheetName
     * @param Region
     * @throws Exception
     */
    public boolean RunTest_multipletest(String strRunMode, String strTestCaseName, String DataSheetName, String Region) throws Exception
    {
          //System.out.println("RunTest  Started = " + Thread.currentThread().getId());
          logger.debug("Thread ID = " + Thread.currentThread().getId() + " common = "+ CommonManager.getInstance().getCommon() +" driver = "+ManagerDriver.getInstance().getWebDriver());
          //fixed for test case status
          PCThreadCache.getInstance().resetProperties();
          //fixed for test case status
          Date d = new Date();
          SimpleDateFormat sdf = new SimpleDateFormat();
          System.out.println("Start Time--------------------------------------------" + d);
          boolean isTestCasePass = false;
//        boolean strYES = true;
          Boolean status = true;
          //DataSheetName = "Data";
          String strColumnName = null;
          String strCondition = null;
          String testCaseType = null;
          if(!Region.isEmpty())
          {
              HTML.properties.setProperty("Region",Region);
          }
          PCThreadCache.getInstance().setProperty("testCaseExecutionStartTime",sdf.format(d));
          if(strRunMode.contains("RunModeYes"))
          {
              strColumnName = "Execution";
              strCondition = "YES";
          }
          else if(strRunMode.contains("RunModeNo"))
          {
              strColumnName = "ID";
              strCondition = strTestCaseName;
          }
          sXL = XlsxReader.getInstance();
          String sheetname = "TestCase";
          int rowcount = sXL.getRowCount(sheetname);
          for(int i = 2; i <= rowcount; i++)
          {
              if(sXL.getCellData(sheetname, strColumnName, i).equalsIgnoreCase(strCondition)
                      && sXL.getCellData(sheetname, "Execution", i).equalsIgnoreCase("YES"))
              {
                  boolean ScriptLevelStatus = true;
//                strYES = false;
                  TCID  = sXL.getCellData(sheetname, "ID", i);
                  TestCaseID  = sXL.getCellData(sheetname, "TestCaseID", i);
                  TestSetID  = sXL.getCellData(sheetname, "TestSetID", i);
                  testCaseType = sXL.getCellData(sheetname, "TestCaseType", i);
                  //UpdateID = sXL.getCellData(sheetname, "UpdateID", i);
                  //PCThreadCache.getInstance().setProperty("UpdateID",UpdateID);
                  PCThreadCache.getInstance().setProperty("TCID",TCID);
                  PCThreadCache.getInstance().setProperty("Row",String.valueOf(i));
                  PCThreadCache.getInstance().setProperty("TestCaseID",TestCaseID);
                  PCThreadCache.getInstance().setProperty("TestSetID",TestSetID);
                  PCThreadCache.getInstance().setProperty("TestCaseType",testCaseType);
                  testcasename = sXL.getCellData(sheetname, "TestCaseName", i);
                  PCThreadCache.getInstance().setProperty("testcasename",testcasename);
                  HTML.fnInitilization(testcasename);
                  logger.info("Thread ID = " + Thread.currentThread().getId() + " -----------------STARTED RUNNING TEST CASE " + testcasename + " EXECUTION----------------- Thread = " +Thread.currentThread().getId());
                  //Commented for graph report
                  if(testCaseType != null && testCaseType.length() >0 && "Regression".equalsIgnoreCase(testCaseType) && HTML.properties.getProperty("DataBaseUpdate").equalsIgnoreCase("YES"))
                  {
                      ReportUtil.initBeginExecuction();
                      ReportUtil.updateDataFeed("IN_PROGRESS");
                  }
                    int colcount = sXL.getColumnCount(sheetname);
                    for(int j = 2; j <= colcount; j++)
                    {
                        try
                          {
                                String ColName = sXL.getCellData(sheetname, j, 1);
                                if(ColName.contains("Component"))
                                  {
                                        TCRow = i;
                                        methodName = sXL.getCellData(sheetname, j, i);
                                        //HTML.properties.setProperty("methodName",methodName);
                                        PCThreadCache.getInstance().setProperty("methodName",methodName);
                                        ////logger.info("methodName ======"+methodName + Thread.currentThread().getId());
                                        
                                        if (!methodName.isEmpty())
                                        {
                                            //no paramater
                                            /*Class noparams[] = {};
                                             //load the AppTest at runtime
                                            Class cls = Class.forName("com.pc.screen." + methodName);
                                            Object obj = cls.newInstance();
                                            HTML.fnInsertResult(testcasename, methodName, "Component should start execution","Started Executing " + methodName + " Component", "PASS", common);
                                            //call the printIt method
                                            Method method = cls.getDeclaredMethod("SCR" + methodName, noparams);*/
                                            if(methodName.contains("_"))
                                            {
                                                String[] methodName2 = methodName.split("_");
                                                String sMultipleComponentTCID= TCID.concat("_"+methodName2[1]);
//                                              PCThreadCache.getInstance().setProperty("methodName",methodName2[0]);
                                                PCThreadCache.getInstance().setProperty("TCID",sMultipleComponentTCID);
                                                logger.info("Thread ID = " + Thread.currentThread().getId() + "---------------Started Executing " + methodName + " function---------------");
                                                HTML.fnInsertResult(testcasename, methodName2[1], "Component execution should start","Started Executing " + methodName + " Component", "PASS");
                                                Class[] paramString = new Class[1];
                                                Class noparams[] = {};
                                                paramString[0] = String.class;
                                                Class cls = Class.forName("com.pc.screen." + methodName2[0]);
                                                Object obj = cls.newInstance();
                                                //Method method = cls.getDeclaredMethod("SCR" + methodName, Common.class);
                                                Method method = cls.getDeclaredMethod("SCR" + methodName2[0],noparams);
                                                //status = (Boolean)method.invoke(obj,  CommonManager.getInstance().getCommon());
                                                status = (Boolean)method.invoke(obj);
                                                PCThreadCache.getInstance().setProperty("TCID",TCID);
                                            }else
                                            {
                                                logger.info("Thread ID = " + Thread.currentThread().getId() +"---------------Started Executing " + methodName + " function---------------");
                                                HTML.fnInsertResult(testcasename, methodName, "Component execution should start","Started Executing " + methodName + " Component", "PASS");
                                                Class[] paramString = new Class[1];
                                                Class noparams[] = {};
                                                paramString[0] = String.class;
                                                Class cls = Class.forName("com.pc.screen." + methodName);
                                                Object obj = cls.newInstance();
                                                //Method method = cls.getDeclaredMethod("SCR" + methodName, Common.class);
                                                Method method = cls.getDeclaredMethod("SCR" + methodName,noparams);
                                                //status = (Boolean)method.invoke(obj,  CommonManager.getInstance().getCommon());
                                                status = (Boolean)method.invoke(obj);
                                            }
                                            //status = (Boolean)method.invoke(obj, common);
                                            if(status)
                                            {
                                                logger.info("Thread ID = " + Thread.currentThread().getId() + " ---------------Completed Executing " + methodName + " function---------------");
                                                //logger.info("methodName 333333333333======"+methodName + Thread.currentThread().getId());
                                                HTML.fnInsertResult(testcasename, methodName, "Component execution should end","Completed Executing " + methodName + " Component", "PASS");
                                            }
                                            else
                                            {
                                                status = handleUnknownAlert();
                                                ScriptLevelStatus = false;
                                                break;
                                            }
                                        }
                                }
                          }
                          catch(Exception e)
                          {
                                e.printStackTrace();
                                logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
                                ScriptLevelStatus = false;
                                break;
                          }
                    }
                    if(ScriptLevelStatus)
                    {
                        logger.info("Thread ID = " + Thread.currentThread().getId() + " -----------------ENDED RUNNING TEST CASE " + testcasename + " EXECUTION-----------------");
                        logger.info("'TestCaseID:' "+TCID+" 'Component:' "+methodName+"");
                        HTML.fnSummaryInsertTestCase();
                        CommonManager.getInstance().getCommon().Terminate();
                        isTestCasePass = true;
                    }
                    else
                    {
                        logger.info("Thread ID = " + Thread.currentThread().getId() + " ---------------Error in executing " + methodName + " function---------------");
                        logger.info("'TestCaseID:' "+TCID+" 'Component:' "+methodName+"");
                        HTML.fnInsertResult(testcasename, methodName, "Component should run properly", "Error in executing: '" + methodName + "'", "FAIL");
                        HTML.fnSummaryInsertTestCase();
                        status = handleUnknownAlert();
                        CommonManager.getInstance().getCommon().Terminate();
                        isTestCasePass = false;
                    }
              }
          }
         /* if(strYES)
          {
              //logger.info("No test case selected as 'YES' in Data sheet");
          }*/
          // Graph report code
          if(testCaseType != null && testCaseType.length() >0 && "Regression".equalsIgnoreCase(testCaseType) && HTML.properties.getProperty("DataBaseUpdate").equalsIgnoreCase("YES")){
              if(isTestCasePass){
                    ReportUtil.updateDataFeed("PASS");
                    ReportUtil.finalizeExec("Pass");
              } else{
                    ReportUtil.finalizeExec("Fail");
                    ReportUtil.updateDataFeed("FAIL");
              }
          }
          Date dd = new Date();
          System.out.println("End Time--------------------------------------------" + dd);
          return  isTestCasePass;
    }
    
    
    /**
     * @function This function use to start the driver script
     * @param strRunMode
     * @param strTestCaseName
     * @param DataSheetName
     * @param Region
     * @throws Exception
     */
    //E2E Framework integration start - modified return type from void to boolean
    public boolean RunTest(String strRunMode, String strTestCaseName, String DataSheetName, String Region) throws Exception
    //E2E Framework integration end
    {
          //System.out.println("RunTest  Started = " + Thread.currentThread().getId());
          logger.debug("Thread ID = " + Thread.currentThread().getId() + " common = "+ CommonManager.getInstance().getCommon() +" driver = "+ManagerDriver.getInstance().getWebDriver());
          PCThreadCache.getInstance().resetProperties();
          Date d = new Date();
          SimpleDateFormat sdf = new SimpleDateFormat();
          System.out.println("Start Time--------------------------------------------" + d);
          //logger.info("-----------------STARTED RUNNING TESTNG METHOD-----------------");
          boolean isTestCasePass = false;
          boolean strYES = true;
          boolean status = true;
          //DataSheetName = "Data";
          String strColumnName = null;
          String strCondition = null;
          String testCaseType = null;
          String ServiceType= null;
          if(!Region.isEmpty())
          {
              HTML.properties.setProperty("Region",Region);
          }
          PCThreadCache.getInstance().setProperty("testCaseExecutionStartTime",sdf.format(d));
          if(strRunMode.contains("RunModeYes"))
          {
              strColumnName = "Execution";
              strCondition = "YES";
          }
          else if(strRunMode.contains("RunModeNo"))
          {
              strColumnName = "ID";
              strCondition = strTestCaseName;
          }
          sXL = XlsxReader.getInstance(); //new XlsxReader(DataSheetName);
          String sheetname = "TestCase";
          int rowcount = sXL.getRowCount(sheetname);
          for(int i = 2; i <= rowcount; i++)
          {
              if(sXL.getCellData(sheetname, strColumnName, i).equalsIgnoreCase(strCondition)
                      && sXL.getCellData(sheetname, "Execution", i).equalsIgnoreCase("YES"))
              {
                  boolean ScriptLevelStatus = true;
                  strYES = false;
                  TCID  = sXL.getCellData(sheetname, "ID", i);
                  TestCaseID  = sXL.getCellData(sheetname, "TestCaseID", i);
                  TestSetID  = sXL.getCellData(sheetname, "TestSetID", i);
                  testCaseType = sXL.getCellData(sheetname, "TestCaseType", i);
                  ServiceType = sXL.getCellData(sheetname, "ServiceType", i);
                  
                  
                  //UpdateID = sXL.getCellData(sheetname, "UpdateID", i);
                  //PCThreadCache.getInstance().setProperty("UpdateID",UpdateID);
                  PCThreadCache.getInstance().setProperty("TCID",TCID);
                  PCThreadCache.getInstance().setProperty("Row",String.valueOf(i));
                  PCThreadCache.getInstance().setProperty("TestCaseID",TestCaseID);
                  PCThreadCache.getInstance().setProperty("TestSetID",TestSetID);
                  PCThreadCache.getInstance().setProperty("TestCaseType",testCaseType);
                  testcasename = sXL.getCellData(sheetname, "TestCaseName", i);
                  PCThreadCache.getInstance().setProperty("testcasename",testcasename);
                  HTML.fnInitilization(testcasename);
                  logger.info("Thread ID = " + Thread.currentThread().getId() + " -----------------STARTED RUNNING TEST CASE " + testcasename + " EXECUTION----------------- Thread = " +Thread.currentThread().getId());
                    int colcount = sXL.getColumnCount(sheetname);
                    for(int j = 2; j <= colcount; j++)
                    {
                        try
 {
                        String ColName = sXL.getCellData(sheetname, j, 1);
                        if (ColName.contains("Component")) {
                            TCRow = i;
                            methodName = sXL.getCellData(sheetname, j, i);
                            PCThreadCache.getInstance().setProperty(
                                    "methodName", methodName);
                            if (!methodName.isEmpty()) {
                                if (ServiceType.contains("SOAP")) {
                                        status = webServiceComponent(methodName , TCID);
                                }
                                if (status) {
                                    logger.info("Thread ID = "
                                            + Thread.currentThread().getId()
                                            + " ---------------Completed Executing "
                                            + methodName
                                            + " function---------------");
                                    // logger.info("methodName 333333333333======"+methodName
                                    // + Thread.currentThread().getId());
                                    HTML.fnInsertResult(testcasename,
                                            methodName,
                                            "Component execution should end",
                                            "Completed Executing " + methodName
                                                    + " Component", "PASS");
                                } else {
                                    status = handleUnknownAlert();
                                    ScriptLevelStatus = false;
                                    break;
                                }
                            }
                        }
                    }
                          catch(Exception e)
                          {
                                e.printStackTrace();
                                logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
                                ScriptLevelStatus = false;
                                break;
                          }
                    }
                    if(ScriptLevelStatus)
                    {
                        logger.info("Thread ID = " + Thread.currentThread().getId() + " -----------------ENDED RUNNING TEST CASE " + testcasename + " EXECUTION-----------------");
                        logger.info("'TestCaseID:' "+TCID+" 'Component:' "+methodName+"");
                        HTML.fnSummaryInsertTestCase();
                        CommonManager.getInstance().getCommon().Terminate();
                        isTestCasePass = true;
                    }
                    else
                    {
                        logger.info("Thread ID = " + Thread.currentThread().getId() + " ---------------Error in executing " + methodName + " function---------------");
                        logger.info("'TestCaseID:' "+TCID+" 'Component:' "+methodName+"");
                        HTML.fnInsertResult(testcasename, methodName, "Component should run properly", "Error in executing: '" + methodName + "'", "FAIL");
                        HTML.fnSummaryInsertTestCase();
                        status = handleUnknownAlert();
                        CommonManager.getInstance().getCommon().Terminate();
                        isTestCasePass = false;
                    }
              }
          }
         /* if(strYES)
          {
              //logger.info("No test case selected as 'YES' in Data sheet");
          }*/
          // Graph report code
          /*if(testCaseType != null && testCaseType.length() >0 && "Regression".equalsIgnoreCase(testCaseType) && HTML.properties.getProperty("DataBaseUpdate").equalsIgnoreCase("YES")){
              if(isTestCasePass){
                    ReportUtil.updateDataFeed("PASS");
                    ReportUtil.finalizeExec("Pass");
              } else{
                    ReportUtil.finalizeExec("Fail");
                    ReportUtil.updateDataFeed("FAIL");
              }
          }*/
          Date dd = new Date();
          System.out.println("End Time--------------------------------------------" + dd);
        //E2E Framework integration start
          return  isTestCasePass;
        //E2E Framework integration end
    }
    
    /**
     * @function This function use to start the driver script
     * @param strRunMode
     * @param strTestCaseName
     * @param DataSheetName
     * @param Region
     * @throws Exception
     */
    //E2E Framework integration start - modified return type from void to boolean
    public boolean RunTest_old(String strRunMode, String strTestCaseName, String DataSheetName, String Region) throws Exception
    //E2E Framework integration end
    {
          //System.out.println("RunTest  Started = " + Thread.currentThread().getId());
          logger.debug("Thread ID = " + Thread.currentThread().getId() + " common = "+ CommonManager.getInstance().getCommon() +" driver = "+ManagerDriver.getInstance().getWebDriver());
          PCThreadCache.getInstance().resetProperties();
          Date d = new Date();
          SimpleDateFormat sdf = new SimpleDateFormat();
          System.out.println("Start Time--------------------------------------------" + d);
          //PropertyConfigurator.configure("log4j.properties");
          //HTML.fnSummaryInitialization("Execution Summary Report");
          //logger.info("-----------------STARTED RUNNING TESTNG METHOD-----------------");
          boolean isTestCasePass = false;
          boolean strYES = true;
          Boolean status = true;
          //DataSheetName = "Data";
          String strColumnName = null;
          String strCondition = null;
          String testCaseType = null;
          if(!Region.isEmpty())
          {
              HTML.properties.setProperty("Region",Region);
          }
//        HTML.properties.setProperty("DataSheetName",DataSheetName);
//        PCThreadCache.getInstance().setProperty("DataSheetName",DataSheetName);
          PCThreadCache.getInstance().setProperty("testCaseExecutionStartTime",sdf.format(d));
          if(strRunMode.contains("RunModeYes"))
          {
              strColumnName = "Execution";
              strCondition = "YES";
          }
          else if(strRunMode.contains("RunModeNo"))
          {
              strColumnName = "ID";
              strCondition = strTestCaseName;
          }
          sXL = XlsxReader.getInstance(); //new XlsxReader(DataSheetName);
          String sheetname = "TestCase";
          int rowcount = sXL.getRowCount(sheetname);
          for(int i = 2; i <= rowcount; i++)
          {
              if(sXL.getCellData(sheetname, strColumnName, i).equalsIgnoreCase(strCondition)
                      && sXL.getCellData(sheetname, "Execution", i).equalsIgnoreCase("YES"))
              {
                  boolean ScriptLevelStatus = true;
                  strYES = false;
                  TCID  = sXL.getCellData(sheetname, "ID", i);
                  TestCaseID  = sXL.getCellData(sheetname, "TestCaseID", i);
                  TestSetID  = sXL.getCellData(sheetname, "TestSetID", i);
                  testCaseType = sXL.getCellData(sheetname, "TestCaseType", i);
                  //UpdateID = sXL.getCellData(sheetname, "UpdateID", i);
                  //PCThreadCache.getInstance().setProperty("UpdateID",UpdateID);
                  PCThreadCache.getInstance().setProperty("TCID",TCID);
                  PCThreadCache.getInstance().setProperty("Row",String.valueOf(i));
                  PCThreadCache.getInstance().setProperty("TestCaseID",TestCaseID);
                  PCThreadCache.getInstance().setProperty("TestSetID",TestSetID);
                  PCThreadCache.getInstance().setProperty("TestCaseType",testCaseType);
                  testcasename = sXL.getCellData(sheetname, "TestCaseName", i);
                  PCThreadCache.getInstance().setProperty("testcasename",testcasename);
                  HTML.fnInitilization(testcasename);
                  logger.info("Thread ID = " + Thread.currentThread().getId() + " -----------------STARTED RUNNING TEST CASE " + testcasename + " EXECUTION----------------- Thread = " +Thread.currentThread().getId());
                  //Commented for graph report
                 /* if(testCaseType != null && testCaseType.length() >0 && "Regression".equalsIgnoreCase(testCaseType) && HTML.properties.getProperty("DataBaseUpdate").equalsIgnoreCase("YES")){
                      ReportUtil.initBeginExecuction();
                      ReportUtil.updateDataFeed("IN_PROGRESS");
                  }*/
                    int colcount = sXL.getColumnCount(sheetname);
                    for(int j = 2; j <= colcount; j++)
                    {
                        try
                          {
                                String ColName = sXL.getCellData(sheetname, j, 1);
                                if(ColName.contains("Component"))
                                  {
                                        TCRow = i;
                                        methodName = sXL.getCellData(sheetname, j, i);
                                        //HTML.properties.setProperty("methodName",methodName);
                                        PCThreadCache.getInstance().setProperty("methodName",methodName);
                                        ////logger.info("methodName ======"+methodName + Thread.currentThread().getId());
                                        
                                        if (!methodName.isEmpty())
                                        {
                                            //no paramater
                                            /*Class noparams[] = {};
                                             //load the AppTest at runtime
                                            Class cls = Class.forName("com.pc.screen." + methodName);
                                            Object obj = cls.newInstance();
                                            HTML.fnInsertResult(testcasename, methodName, "Component should start execution","Started Executing " + methodName + " Component", "PASS", common);
                                            //call the printIt method
                                            Method method = cls.getDeclaredMethod("SCR" + methodName, noparams);*/
                                            logger.info("Thread ID = " + Thread.currentThread().getId() + "---------------Started Executing " + methodName + " function---------------");
                                            Class[] paramString = new Class[1];
                                            Class noparams[] = {};
                                            paramString[0] = String.class;
                                            Class cls = Class.forName("com.pc.screen." + methodName);
                                            Object obj = cls.newInstance();
                                            //Method method = cls.getDeclaredMethod("SCR" + methodName, Common.class);
                                            Method method = cls.getDeclaredMethod("SCR" + methodName,noparams);
                                            //status = (Boolean)method.invoke(obj,  CommonManager.getInstance().getCommon());
                                            status = (Boolean)method.invoke(obj);
                                            //status = (Boolean)method.invoke(obj, common);
                                            if(status)
                                            {
                                                logger.info("Thread ID = " + Thread.currentThread().getId() + " ---------------Completed Executing " + methodName + " function---------------");
                                                //logger.info("methodName 333333333333======"+methodName + Thread.currentThread().getId());
                                                HTML.fnInsertResult(testcasename, methodName, "Component execution should end","Completed Executing " + methodName + " Component", "PASS");
                                            }
                                            else
                                            {
                                                status = handleUnknownAlert();
                                                ScriptLevelStatus = false;
                                                break;
                                            }
                                        }
                                }
                          }
                          catch(Exception e)
                          {
                                e.printStackTrace();
                                logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
                                ScriptLevelStatus = false;
                                break;
                          }
                    }
                    if(ScriptLevelStatus)
                    {
                        logger.info("Thread ID = " + Thread.currentThread().getId() + " -----------------ENDED RUNNING TEST CASE " + testcasename + " EXECUTION-----------------");
                        logger.info("'TestCaseID:' "+TCID+" 'Component:' "+methodName+"");
                        HTML.fnSummaryInsertTestCase();
                        CommonManager.getInstance().getCommon().Terminate();
                        isTestCasePass = true;
                    }
                    else
                    {
                        logger.info("Thread ID = " + Thread.currentThread().getId() + " ---------------Error in executing " + methodName + " function---------------");
                        logger.info("'TestCaseID:' "+TCID+" 'Component:' "+methodName+"");
                        HTML.fnInsertResult(testcasename, methodName, "Component should run properly", "Error in executing: '" + methodName + "'", "FAIL");
                        HTML.fnSummaryInsertTestCase();
                        status = handleUnknownAlert();
                        CommonManager.getInstance().getCommon().Terminate();
                        isTestCasePass = false;
                    }
              }
          }
         /* if(strYES)
          {
              //logger.info("No test case selected as 'YES' in Data sheet");
          }*/
          // Graph report code
          /*if(testCaseType != null && testCaseType.length() >0 && "Regression".equalsIgnoreCase(testCaseType) && HTML.properties.getProperty("DataBaseUpdate").equalsIgnoreCase("YES")){
              if(isTestCasePass){
                    ReportUtil.updateDataFeed("PASS");
                    ReportUtil.finalizeExec("Pass");
              } else{
                    ReportUtil.finalizeExec("Fail");
                    ReportUtil.updateDataFeed("FAIL");
              }
          }*/
          Date dd = new Date();
          System.out.println("End Time--------------------------------------------" + dd);
        //E2E Framework integration start
          return  isTestCasePass;
        //E2E Framework integration end
    }
    
    /**
     * @function This function use to start the driver script
     * @param strRunMode
     * @param strTestCaseName
     * @param DataSheetName
     * @param Region
     * @throws Exception
     */
    public boolean RunTest8_7(String strRunMode, String strTestCaseName, String DataSheetName, String Region) throws Exception
    {
          //System.out.println("RunTest  Started = " + Thread.currentThread().getId());
          logger.debug("Thread ID = " + Thread.currentThread().getId() + " common = "+ CommonManager.getInstance().getCommon() +" driver = "+ManagerDriver.getInstance().getWebDriver());
          //fixed for test case status
          PCThreadCache.getInstance().resetProperties();
          //fixed for test case status
          Date d = new Date();
          SimpleDateFormat sdf = new SimpleDateFormat();
          System.out.println("Start Time--------------------------------------------" + d);
          boolean isTestCasePass = false;
//        boolean strYES = true;
          Boolean status = true;
          //DataSheetName = "Data";
          String strColumnName = null;
          String strCondition = null;
          String testCaseType = null;
          if(!Region.isEmpty())
          {
              HTML.properties.setProperty("Region",Region);
          }
          PCThreadCache.getInstance().setProperty("testCaseExecutionStartTime",sdf.format(d));
          if(strRunMode.contains("RunModeYes"))
          {
              strColumnName = "Execution";
              strCondition = "YES";
          }
          else if(strRunMode.contains("RunModeNo"))
          {
              strColumnName = "ID";
              strCondition = strTestCaseName;
          }
          sXL = XlsxReader.getInstance();
          String sheetname = "TestCase";
          int rowcount = sXL.getRowCount(sheetname);
          for(int i = 2; i <= rowcount; i++)
          {
              if(sXL.getCellData(sheetname, strColumnName, i).equalsIgnoreCase(strCondition)
                      && sXL.getCellData(sheetname, "Execution", i).equalsIgnoreCase("YES"))
              {
                  boolean ScriptLevelStatus = true;
//                strYES = false;
                  TCID  = sXL.getCellData(sheetname, "ID", i);
                  TestCaseID  = sXL.getCellData(sheetname, "TestCaseID", i);
                  TestSetID  = sXL.getCellData(sheetname, "TestSetID", i);
                  testCaseType = sXL.getCellData(sheetname, "TestCaseType", i);
                  //UpdateID = sXL.getCellData(sheetname, "UpdateID", i);
                  //PCThreadCache.getInstance().setProperty("UpdateID",UpdateID);
                  PCThreadCache.getInstance().setProperty("TCID",TCID);
                  PCThreadCache.getInstance().setProperty("Row",String.valueOf(i));
                  PCThreadCache.getInstance().setProperty("TestCaseID",TestCaseID);
                  PCThreadCache.getInstance().setProperty("TestSetID",TestSetID);
                  PCThreadCache.getInstance().setProperty("TestCaseType",testCaseType);
                  testcasename = sXL.getCellData(sheetname, "TestCaseName", i);
                  PCThreadCache.getInstance().setProperty("testcasename",testcasename);
                  HTML.fnInitilization(testcasename);
                  logger.info("Thread ID = " + Thread.currentThread().getId() + " -----------------STARTED RUNNING TEST CASE " + testcasename + " EXECUTION----------------- Thread = " +Thread.currentThread().getId());
                  //Commented for graph report
                  if(testCaseType != null && testCaseType.length() >0 && "Regression".equalsIgnoreCase(testCaseType) && HTML.properties.getProperty("DataBaseUpdate").equalsIgnoreCase("YES"))
                  {
                      ReportUtil.initBeginExecuction();
                      ReportUtil.updateDataFeed("IN_PROGRESS");
                  }
                    int colcount = sXL.getColumnCount(sheetname);
                    for(int j = 2; j <= colcount; j++)
                    {
                        try
                          {
                                String ColName = sXL.getCellData(sheetname, j, 1);
                                if(ColName.contains("Component"))
                                  {
                                        TCRow = i;
                                        methodName = sXL.getCellData(sheetname, j, i);
                                        //HTML.properties.setProperty("methodName",methodName);
                                        PCThreadCache.getInstance().setProperty("methodName",methodName);
                                        ////logger.info("methodName ======"+methodName + Thread.currentThread().getId());
                                        
                                        if (!methodName.isEmpty())
                                        {
                                            //no paramater
                                            /*Class noparams[] = {};
                                             //load the AppTest at runtime
                                            Class cls = Class.forName("com.pc.screen." + methodName);
                                            Object obj = cls.newInstance();
                                            HTML.fnInsertResult(testcasename, methodName, "Component should start execution","Started Executing " + methodName + " Component", "PASS", common);
                                            //call the printIt method
                                            Method method = cls.getDeclaredMethod("SCR" + methodName, noparams);*/
                                            if(methodName.contains("_"))
                                            {
                                                String[] methodName2 = methodName.split("_");
                                                String sMultipleComponentTCID= TCID.concat("_"+methodName2[1]);
//                                              PCThreadCache.getInstance().setProperty("methodName",methodName2[0]);
                                                PCThreadCache.getInstance().setProperty("TCID",sMultipleComponentTCID);
                                                logger.info("Thread ID = " + Thread.currentThread().getId() + "---------------Started Executing " + methodName + " function---------------");
                                                HTML.fnInsertResult(testcasename, methodName2[1], "Component execution should start","Started Executing " + methodName + " Component", "PASS");
                                                Class[] paramString = new Class[1];
                                                Class noparams[] = {};
                                                paramString[0] = String.class;
                                                Class cls = Class.forName("com.pc.screen." + methodName2[0]);
                                                Object obj = cls.newInstance();
                                                //Method method = cls.getDeclaredMethod("SCR" + methodName, Common.class);
                                                Method method = cls.getDeclaredMethod("SCR" + methodName2[0],noparams);
                                                //status = (Boolean)method.invoke(obj,  CommonManager.getInstance().getCommon());
                                                status = (Boolean)method.invoke(obj);
                                                PCThreadCache.getInstance().setProperty("TCID",TCID);
                                            }else
                                            {
                                                logger.info("Thread ID = " + Thread.currentThread().getId() +"---------------Started Executing " + methodName + " function---------------");
                                                HTML.fnInsertResult(testcasename, methodName, "Component execution should start","Started Executing " + methodName + " Component", "PASS");
                                                Class[] paramString = new Class[1];
                                                Class noparams[] = {};
                                                paramString[0] = String.class;
                                                Class cls = Class.forName("com.pc.screen." + methodName);
                                                Object obj = cls.newInstance();
                                                //Method method = cls.getDeclaredMethod("SCR" + methodName, Common.class);
                                                Method method = cls.getDeclaredMethod("SCR" + methodName,noparams);
                                                //status = (Boolean)method.invoke(obj,  CommonManager.getInstance().getCommon());
                                                status = (Boolean)method.invoke(obj);
                                            }
                                            //status = (Boolean)method.invoke(obj, common);
                                            if(status)
                                            {
                                                logger.info("Thread ID = " + Thread.currentThread().getId() + " ---------------Completed Executing " + methodName + " function---------------");
                                                //logger.info("methodName 333333333333======"+methodName + Thread.currentThread().getId());
                                                HTML.fnInsertResult(testcasename, methodName, "Component execution should end","Completed Executing " + methodName + " Component", "PASS");
                                            }
                                            else
                                            {
                                                status = handleUnknownAlert();
                                                ScriptLevelStatus = false;
                                                break;
                                            }
                                        }
                                }
                          }
                          catch(Exception e)
                          {
                                e.printStackTrace();
                                logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
                                ScriptLevelStatus = false;
                                break;
                          }
                    }
                    if(ScriptLevelStatus)
                    {
                        logger.info("Thread ID = " + Thread.currentThread().getId() + " -----------------ENDED RUNNING TEST CASE " + testcasename + " EXECUTION-----------------");
                        logger.info("'TestCaseID:' "+TCID+" 'Component:' "+methodName+"");
                        HTML.fnSummaryInsertTestCase();
                        CommonManager.getInstance().getCommon().Terminate();
                        isTestCasePass = true;
                    }
                    else
                    {
                        logger.info("Thread ID = " + Thread.currentThread().getId() + " ---------------Error in executing " + methodName + " function---------------");
                        logger.info("'TestCaseID:' "+TCID+" 'Component:' "+methodName+"");
                        HTML.fnInsertResult(testcasename, methodName, "Component should run properly", "Error in executing: '" + methodName + "'", "FAIL");
                        HTML.fnSummaryInsertTestCase();
                        status = handleUnknownAlert();
                        CommonManager.getInstance().getCommon().Terminate();
                        isTestCasePass = false;
                    }
              }
          }
         /* if(strYES)
          {
              //logger.info("No test case selected as 'YES' in Data sheet");
          }*/
          // Graph report code
          if(testCaseType != null && testCaseType.length() >0 && "Regression".equalsIgnoreCase(testCaseType) && HTML.properties.getProperty("DataBaseUpdate").equalsIgnoreCase("YES")){
              if(isTestCasePass){
                    ReportUtil.updateDataFeed("PASS");
                    ReportUtil.finalizeExec("Pass");
              } else{
                    ReportUtil.finalizeExec("Fail");
                    ReportUtil.updateDataFeed("FAIL");
              }
          }
          Date dd = new Date();
          System.out.println("End Time--------------------------------------------" + dd);
          return  isTestCasePass;
    }
    
    /**
     * @function use to handle the unknown alert
     * @return true/false
     */
    public boolean handleUnknownAlert()
    {
        boolean status = false;
        try{
                Alert al = ManagerDriver.getInstance().getWebDriver().switchTo().alert();
                al.dismiss();
                logger.info("Alert found now quiting the browser");
                status = false;
            }
        catch(Exception e)
            {
                status = false;
                logger.info("No Alert found");
            }
        return status;
        
    }
    /*//Deprecated Do Not Use
    public Recordset GetDataFromExcelDoNotUse(String strFileName, String strQuery1) throws FilloException
    {
           Fillo fillo=new Fillo();
           Connection connection=(Connection) fillo.getConnection(strFileName);
           Recordset recordset = connection.executeQuery(strQuery1);
           return recordset;
    }
  
    //Deprecated Do Not Use
    public void UpdateDataInExcelDoNotUse(String strFileName, String strQuery) throws FilloException
    {
         Fillo fillo=new Fillo();
         Connection connection=fillo.getConnection(strFileName);
         connection.executeUpdate(strQuery);
         connection.close();
    }
    
    //Deprecated Do Not Use
    public void UpdateStatusInExcelDoNotUse(String strFileName, String strQuery) throws FilloException
    {
         Fillo fillo=new Fillo();
         Connection connection=fillo.getConnection(strFileName);
         connection.executeUpdate(strQuery);
         connection.close();
    }*/
    
    /**
     * @function This function use to retrieve Product Select Shell / SOR
     * @return String
     * @throws Exception
     */
     public String getSpecifiedExcelValue(String strSheetName,String strProductSelection) throws Exception
     {
            String strProduct = "";
            XlsxReader sXL;
            boolean blnFlag = false;
            sXL = XlsxReader.getInstance();//new XlsxReader(HTML.properties.getProperty("DataSheetName"));
            int rowcount = sXL.getRowCount(strSheetName);
            try
            {
                   for(int i=2;i<=rowcount;i++)
                   {
                         String value = sXL.getCellData(strSheetName, 0, i);
                         if(!value.isEmpty())
                         {
                                if(PCThreadCache.getInstance().getProperty("TCID").equals(value))
                                {
                                       int colcount = sXL.getColumnCount(strSheetName);
                                       for(int j = 2; j <= colcount; j++)
                                       {
                                                 String ColName = sXL.getCellData(strSheetName, j, 1);
                                                 if(ColName.equals(strProductSelection))
                                                 {
                                                        strProduct = sXL.getCellData(strSheetName, j, i);
                                                        blnFlag = true;
                                                        break;
                                                 }
                                       }
                                }
                         }
                         if (blnFlag == true)
                         {
                                break;
                         }
                   }
            }
            catch (Exception e)
            {
                   blnFlag = false;
                   logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
                   e.printStackTrace();
            }
            return strProduct;
     }
     
     /**
      * @function Use to click the check box according to the check box label
      * @param obj
      * @param readCol
      * @param actionCol
      * @param strReadString
      * @param actionObjetName
      * @param sTagName
      * @return true/false
      * @throws Exception
      */
     public  Boolean SelectCheckBoxOnTable(By obj, int readCol, int actionCol, String strReadString, String actionObjetName, String sTagName) throws Exception
     {
       boolean Status=false;
       boolean SearchString=false;
       boolean ActionObject=false;
       WebElement mytable = ManagerDriver.getInstance().getWebDriver().findElement(obj);
       List<WebElement> allrows = mytable.findElements(By.tagName("tr"));
       for(int i = 0; i <= allrows.size()-1; i++)
       {
              List<WebElement> Cells = allrows.get(i).findElements(By.tagName("td"));
              String readText = Cells.get(readCol).getText();
              if (readText.contains(strReadString))
              {
                     SearchString = true;
                   
                     Cells.get(actionCol).click();
                     ActionObject = true;
              
          }
            if(ActionObject == true)
            {
                   break;
            }
       }
       if(SearchString)
       {
                   //logger.info("Search String available in the table. '" + strReadString + "'");
                   HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and srarch string is  '" + actionObjetName + "'", "PASS");
                   if(ActionObject)
                     {
                                //logger.info("Search and click on object in the table cell and object name is '" + actionObjetName + "'");
                                HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search object in the table cell and click on object. Object name is '" + actionObjetName + "'","System searched object in the table and clicked on object. object name is '" + actionObjetName + "'", "PASS");
                                Status = true;
                     }
                     else
                     {
                                //logger.info("Search and click on object in the table cell and object name is '" + actionObjetName + "'");
                                HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search object in the table cell and click on object. Object name is '" + actionObjetName + "'","System searched object in the table and clicked on object. object name is '" + actionObjetName + "'", "FAIL");
                                Status = false;
                     }
       }
       else
       {
                   //logger.info("Search String not available in the table. '" + strReadString + "'");
                   HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and srarch string is  '" + actionObjetName + "'", "FAIL");
                   Status = false;
       }
       return Status;
     }
     
     /**
      * @function This function use to Create Activity as per the input
      * @return String
      * @throws IOException
      * @throws Exception
      */
     public  boolean SelectActivity(String strValue) throws IOException
     {
               boolean Status = false;
//               By option = By.xpath("//span[starts-with(text(),'"+strValue+"')]");
               By option = By.xpath("//*[contains(@id,'NewActivityMenuItemSet')]//span[contains(text(), '"+strValue+"')]");
              try {
                  
                     Status = CommonManager.getInstance().getCommon().SafeAction(option, "scl","scl");
                     Status = CommonManager.getInstance().getCommon().SafeAction(option, "ele","ele");
              } catch (Exception e) {
                     e.printStackTrace();
                     logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
              }
              if(Status)
              {
                     //logger.info("Clicked on '" + option + "' element or button or link and element '"+ option + "'");
              HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should click on '" + option + "' element or button or link", "Clicked on '" + option + "' element or button or link", "PASS");
              }
              else
              {
                     //logger.info("Object not enabled or displayed or not clickable '"+ option + "'");
                     HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should display '" + option + "' in screen", "'" + option + "' not displayed in screen", "FAIL");
              }
       
     return Status;
     }
   
    /**
    * @function This function use to Select the data from the table and performs the action accordingly
    * @param obj
    * @param readCol
    * @param actionCol
    * @param strReadString
    * @param lobType
    * @param NoOfSubmissions
    * @param sTagName
    * @return true/false
    * @throws Exception
    */
    public  Boolean ActionOnTableSelect(By obj,int readCol, int actionCol, String strReadString, String lobType,String NoOfSubmissions, String sTagName) throws Exception
    {
      boolean Status=false;
      boolean SearchString=false;
      boolean ActionObject=false;
      WebElement selectObj = null;
      String readText = "";
      WebElement mytable = ManagerDriver.getInstance().getWebDriver().findElement(obj);
      List<WebElement> sElement = null;
      List<WebElement> Cells = null;
      List<WebElement> allrows = mytable.findElements(By.tagName("tr"));
      for(int i = 0; i <= allrows.size()-1; i++)
      {
          Cells = allrows.get(i).findElements(By.tagName("td"));
          readText = Cells.get(readCol).getText();
          if (readText.contains(strReadString))
          {
              SearchString = true;
              switch (lobType.toUpperCase())
              {
                  case "SHELL":
                      selectObj = Cells.get(0).findElement(By.tagName(sTagName));
                      break;
                  case "SOR":
                      selectObj = Cells.get(1).findElement(By.tagName(sTagName));
                      break;
              }
              // Click on the specified column
              selectObj.click();
              // Select specified item from the list
              sElement = ManagerDriver.getInstance().getWebDriver().findElements(By.tagName("li"));
              for(int j=0; j<=sElement.size()-1; j++)
              {
                 if (sElement.get(j).getText().contains(NoOfSubmissions))
                 {
                   sElement.get(j).click();
                   ActionObject = true;
                   break;
                 }
              }
              
         }
         if(ActionObject == true)
         {
             break;
         }
      }
      if(SearchString)
      {
            //logger.info("Search String available in the table. '" + strReadString + "'");
          HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and srarch string is  '" + readText + "'", "PASS");
            if(ActionObject)
              {
                    //logger.info("Search and click on object in the table cell and object name is '" + NoOfSubmissions + "'");
                  HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search object in the table cell and click on object. Object name is '" + NoOfSubmissions + "'","System searched object in the table and clicked on object. object name is '" + NoOfSubmissions + "'", "PASS");
                    Status = true;
              }
              else
              {
                    //logger.info("Search and click on object in the table cell and object name is '" + NoOfSubmissions + "'");
                  HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search object in the table cell and click on object. Object name is '" + NoOfSubmissions + "'","System searched object in the table is not available. object name is '" + NoOfSubmissions + "'", "FAIL");
                    Status = false;
              }
      }
      else
      {
            //logger.info("Search String not available in the table. '" + strReadString + "'");
        HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and srarch string is  '" + NoOfSubmissions + "'", "FAIL");
            Status = false;
      }
      return Status;
    }
    /**
     * @function This function use to verify the text from the table according to the column
     * @param obj
     * @param readTextCol,readTextRow
     * @param strReadString
     * @return true/false
     * @throws Exception
     */
    public boolean VerifyTextFromTable(By obj, int readTextRow, int readTextCol, String strReadString) throws Exception
    {
          boolean SearchString = false;
          String readText = null;
          WebElement mytable = ManagerDriver.getInstance().getWebDriver().findElement(obj);
          List<WebElement> allrows = mytable.findElements(By.tagName("tr"));
          List<WebElement> Cells = allrows.get(readTextRow).findElements(By.tagName("td"));
          readText = Cells.get(readTextCol).getText();
              if(readText.contains(strReadString))
              {
                  SearchString = true;
                  logger.info("Search String available in the table. '" + strReadString + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and search string is  '" + strReadString + "'", "PASS");
                  
              }
              else
              {
                    logger.info("Search String not available in the table. '" + strReadString + "'");
                    HTML.fnInsertResult(PCThreadCache.getInstance().getProperty("testcasename"), PCThreadCache.getInstance().getProperty("methodName"), "System should search string in table and Search string is '" + strReadString + "'","System searched string in table and search string is  '" + strReadString + "'", "FAIL");
              }
          
          return SearchString;
    }
    /**
     * @function Ability to get the text of the element
     * @param bylocator
     * @param iWaitTime
     * @return String
     * @throws Exception
    **/
    public String ReadElementFromListEditBox(By bylocator, int iWaitTime) throws Exception
    {
        WaitUntilClickable(bylocator, iWaitTime);
        WebElement element = ManagerDriver.getInstance().getWebDriver().findElement(bylocator);
        return element.getAttribute("value");
    }
    
    /**
     * @function used to get text from table
     * @param obj
     * @param getTextRow
     * @param getTextCol
     * @param tagName
     * @return value
     * @throws Exception
     */
     public  String GetTextFromTableTagName(By obj, int getTextRow, int getTextCol,String tagName) throws Exception
     {
              String text = null;
              WebElement mytable = ManagerDriver.getInstance().getWebDriver().findElement(obj);
              List<WebElement> allrows = mytable.findElements(By.tagName("tr"));
              List<WebElement> Cells = allrows.get(getTextRow).findElements(By.tagName("td"));
              List<WebElement> NewCells = Cells.get(getTextCol).findElements(By.tagName(tagName));
              text=NewCells.get(0).getText();
              return text;
     }
     
     
     
     
    public boolean webServiceComponent(String sheetNames, String tcId) throws InterruptedException {
        
        /*if (sheetNames.contains("SeqSvc460QuoteDownloadSC01")||sheetNames.contains("SeqSvc460IssueDownloadSC01")
                || sheetNames.contains("SeqSvc424EndDownloadSC01") || sheetNames.contains("SeqSvc427CanDownloadSC01")
                || sheetNames.contains("SeqSvc425ReinstDownloadSC01") || sheetNames.contains("SeqSvc426ChgRenewDownloadSC01")
                || sheetNames.contains("SeqSvc543AuditDownloadSC01"))
        {
            Thread.sleep(40000);
        }*/
        try {
            sXL = XlsxReader.getInstance();
            if (sheetNames != null) {
                for (String sheetName : sheetNames.split(",")) {
                    XSSFSheet sheet = sXL.getSheet(sheetName);
                    if (sheet == null)
                        return false;
                    logger.info("erasing previous error if any present for the testcase id in \"testCases\" sheet");
                    updateErrorInfo(tcId, ""); //erasing previous error if any present for the testcase id in "testCases" sheet.
                    XSSFRow firstRow = sheet.getRow(0);
                    int lastCellNum = firstRow.getLastCellNum();
                    Map<Integer, String> ipMap = new HashMap<>();
                    Map<Integer, String> opMap = new HashMap<>();

                    Map<Integer, String> requestFileMap = new HashMap<>();
                    Map<Integer, String> dynamicIpOpRequired = new HashMap<>();
                    readXLSHeaderAndCreateMaps(firstRow, lastCellNum, ipMap,
                            opMap, requestFileMap, dynamicIpOpRequired);
                    int rowCount = sheet.getLastRowNum() + 1;
                    // Start of the code for remove the random ip and op values
                    setIpOpEmptyValues(sheet,ipMap,opMap,tcId,dynamicIpOpRequired,requestFileMap);
                    // End of the code for removing the random ip and op values
                    Map<String, String> ipXpathMap = new HashMap<>();
                    Map<String, String> opXpathMap = new HashMap<>();
                    getXpathsFromXls(sheet.getRow(1), ipMap,opMap,ipXpathMap,opXpathMap);
                    
                    HashMap<String, Integer> curSheetHeaddersMap = new HashMap<>();
                    readXLSHeader(sheetName, curSheetHeaddersMap );
                    for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++) {
                        XSSFRow currentRow = sheet.getRow(rowIndex);
                        if (currentRow == null) {
                            break;
                        }
                        int headderIndex = curSheetHeaddersMap.get("ID") ==null?-1:curSheetHeaddersMap.get("ID");
                        if(headderIndex < 0){
                            logNoHeadderError(sheetName,"ID");
                            return false;
                        }
                        XSSFCell idCell = currentRow.getCell(headderIndex);
                        String testCaseID = sXL.getCellValue(idCell);
                        if (testCaseID == null || testCaseID.isEmpty() || !testCaseID.equalsIgnoreCase(tcId)) {
                            continue;
                        }
                        
                        String THREAD_WAIT_IND = "funWait";
                        headderIndex = curSheetHeaddersMap.get(THREAD_WAIT_IND) ==null?-1:curSheetHeaddersMap.get(THREAD_WAIT_IND);
                        if(headderIndex < 0){
                               logNoHeadderError(sheetName,THREAD_WAIT_IND);
                        } else {
                            long waitTime = 0;
                            XSSFCell waitCell = currentRow
                                    .getCell(headderIndex);
                            String waitIndicator = sXL.getCellValue(waitCell);
                            if (waitIndicator != null
                                    && !waitIndicator.trim().isEmpty()) {
                                try {
                                    waitTime=Integer.valueOf(waitIndicator);
                                } catch (Exception e) {
                                    if ("YES".equalsIgnoreCase(waitIndicator.trim())
                                            || "Y".equalsIgnoreCase(waitIndicator
                                                    .trim())) {
                                        //long waitTime = 0;
                                        String waittime = HTML.properties
                                                .getProperty("SERVICEWAIT");
                                        waitTime = waittime == null ? 0 : Integer
                                                .valueOf(waittime);
                                        
                                    }

                                }
                                if(waitTime>0)
                                    Thread.sleep(waitTime);
                            }
                            
                        }
                        
                        Document document = null;
                        String requestXmlPath = null;
                        String requestPayloadFile=null;
                        for (Integer index : requestFileMap.keySet()) {
                            XSSFCell cell = currentRow.getCell(index);
                            String cellValue = sXL.getCellValue(cell);
                            if (cellValue != null && !cellValue.isEmpty()) {
                                requestXmlPath = "Request/" + cellValue
                                        + ".xml";
                                requestPayloadFile = String.valueOf(cellValue).trim();
                                document = getRequestDocument(cellValue);
                                break;
                            }
                        }
                        if (document == null) {
                            logger.error("request file not Exist. Path:"
                                    + requestXmlPath);
                            updateErrorInfo(tcId, "request file not Exist. Path:"
                                    + requestXmlPath+" for the testcase : "+tcId + "in sheet "+sheetName);
                            return false;
                        }
                        if (document != null) {
                            boolean validateXpaths = validateXpaths(document , ipXpathMap);
                            if(!validateXpaths){
                                logger.error("given xpaths for Input and output parameters are wrong. Please correct it");
                                return validateXpaths;
                                    
                            }else{
                                logger.info("input parameters XPats are valid and proceding to hit the web service");
                            }
                            //check the available dynamic ip/op parameters configurations for the service in.
                            if (dynamicIpOpRequired != null
                                    && !dynamicIpOpRequired.isEmpty()) {
                                XSSFCell ipOpRequiredCell = currentRow
                                        .getCell(getKeyByValue(
                                                dynamicIpOpRequired,
                                                "DynamicIpOpRequired"));
                                String ipOpRequired = sXL
                                        .getCellValue(ipOpRequiredCell);
                                if ("Yes".equalsIgnoreCase(ipOpRequired.trim())
                                        || "Y".equalsIgnoreCase(ipOpRequired
                                                .trim())) {
                                    checkIpOpconfigAndUpdateSheet(ipMap, opMap,
                                            tcId, sheetName, rowIndex);
                                }
                            }else{
                                logger.info("dynamic updation/generating values for input and output parameters is not available for the test case id:"+tcId +" in sheet \""+sheetName+"\"");
                            }
                            
                            // Request parameters modification as per the excel
                            // sheet.
                            modifydocumentwithIPParams(ipMap, ipXpathMap , currentRow, document);
                            logRequest(document);
                            savePayload(document,requestPayloadFile,tcId,"Request/sent");
                            /*headderIndex = -1;
                            headderIndex = curSheetHeaddersMap.get("EndPointUrl") ==null?-1:curSheetHeaddersMap.get("EndPointUrl");
                            if(headderIndex < 0){
                                logNoHeadderError(sheetName,"EndPointUrl");
                                return false;
                            }
                            XSSFCell serviceUrlcell = currentRow.getCell(headderIndex);
                            String serviceUrlCellValue = sXL.getCellValue(serviceUrlcell);*/
                            String serviceUrlCellValue = getServiceEndUrl(sheetName);
                            if(serviceUrlCellValue == null || serviceUrlCellValue.isEmpty()){
                                logger.error("Service End URL is not present.PLease correct it");
                                return false;
                            }

                            // service call
                            String response = callWebService(requestXmlPath,
                                    document, serviceUrlCellValue, tcId);

                            // read the response
                            if(response == null){
                                logger.error("response from the main service is null. So terminating without checking the sub level services.");
                                updateErrorInfo(tcId, "error occured while consuming the service:"+sheetName+" and url: "+serviceUrlCellValue+" fot the testcase:"+tcId);
                                return false;
                            }else if (response != null) {
                                logger.info("the response from the service Url:" + serviceUrlCellValue+" is "+response);
                                Document responseDocument = convertStringToDocument(response);
                                savePayload(responseDocument,requestPayloadFile,tcId,"Response/received");
                                String faultXpath = "Envelope/Body/Fault";
                                expectedXPaths.add(faultXpath);
                                Node faultNode = getElementNodeByXPath(
                                        responseDocument, faultXpath);
                                if (faultNode == null) {

                                    // success scenario
                                    // no error - opmap iteration and populate
                                    // values into xls
                                    if(!validateXpaths(responseDocument , opXpathMap)){
                                        logger.error("Output parameters are invalid. So returning true with out updating output parameters in the sheet.");
                                        return true;
                                    }
                                    logger.info("output parameter values :");
                                    for (Integer index : opMap.keySet()) {
                                        String xPath = opXpathMap.get(opMap.get(index));
                                        String elementValueByXPath = getElementValueByXPath(
                                                responseDocument, xPath);
                                        logger.info(xPath +" : "+ elementValueByXPath);
                                        sXL.setCellData(sheet,
                                                elementValueByXPath, rowIndex,
                                                index);
                                    }
                                    return true;

                                } else {
                                    // error/failed scenario
                                    // error - iterate error map
                                    //SubLevelComponents sheet contains all sub level services.
                                    //look for the SubLevelComponents sheet to find the root cause of the fault by
                                    // executing sub level services if any.
                                    
                                    logger.info("fault codes found in the response and started checking the sub level components of testcase:"+testCaseID+" which are configured in the SubLevelComponents sheet .");
                                    updateXlsWithFaultInfo(responseDocument, tcId);
                                    String subLevelSheetComponent = "SubLevelComponents";
                                    XSSFSheet subLevelSheet = sXL.getSheet(subLevelSheetComponent);
                                    HashMap<String, Integer> slSheetHeaddersMap = new HashMap<>();
                                    readXLSHeader(subLevelSheetComponent, slSheetHeaddersMap );
                                    if (subLevelSheet == null){
                                        logger.error("SubLevelComponents sheet not found and returning false");
                                        updateErrorInfo(tcId, "SubLevelComponents sheet not found and returning false");
                                        return false;
                                    }
                                    int rowCountInSubLevel = subLevelSheet.getLastRowNum() + 1;
                                    for(int index=1;index <=rowCountInSubLevel;index++){
                                        XSSFRow subLevelCurrentRow = subLevelSheet.getRow(index);
                                        if (subLevelCurrentRow == null) {
                                            break;
                                        }
                                        int maxCellNum = subLevelCurrentRow.getLastCellNum();
                                        
                                        int slHeadderIndex = slSheetHeaddersMap.get("Service") ==null?-1:slSheetHeaddersMap.get("Service");
                                        if(slHeadderIndex < 0){
                                            logNoHeadderError(sheetName,"Service");
                                            return false;
                                        }
                                            XSSFCell mainServiceCell = subLevelCurrentRow.getCell(slHeadderIndex);
                                             String serviceName = sXL.getCellValue(mainServiceCell);
                                             serviceName = String.valueOf(serviceName).trim();
                                             if(serviceName != null && serviceName.equalsIgnoreCase(sheetName)){
                                                 for(int colIndex=1;colIndex<=maxCellNum;colIndex++){
                                                     XSSFCell subserviceCell = subLevelCurrentRow.getCell(colIndex); //Ser482WC
                                                     String subservice = sXL.getCellValue(subserviceCell);
                                                     boolean flag = callSublevelWebService(String.valueOf(subservice).trim(), currentRow,tcId);
                                                     if(!flag){
                                                         return false;
                                                     }
                                                 }
                                             }else{
                                                 continue;
                                             }
                                    }
                                }
                            }

                        }

                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return false;
    }

    private void updateXlsWithFaultInfo(Document document, String tcId) throws Exception {
        String errorMsg = "";
        String faultCodeXpath = "Envelope/Body/Fault/faultcode";
        expectedXPaths.add(faultCodeXpath);
        String faultMessageXpath = "Envelope/Body/Fault/faultstring";
        expectedXPaths.add(faultMessageXpath);
        errorMsg = "Fault Code: [" + getElementValueByXPath(document, faultCodeXpath) + "] - Fault Message: [" + getElementValueByXPath(document, faultMessageXpath) + "]" ;
        updateErrorInfo(tcId, errorMsg);
    }

    private void logRequest(Document document) {
        logger.info("the request payload is : "+trim(convertDocumentToString(document)));
    }
    public String trim(String input) {
        BufferedReader reader = new BufferedReader(new StringReader(input));
        StringBuffer result = new StringBuffer();
        try {
            String line;
            while ( (line = reader.readLine() ) != null)
                result.append(line.trim());
            return result.toString();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private String getServiceEndUrl(String serviceName) throws Exception {
        String endUrl = null;
        String endUrlSheet = "EndUrls";
        XSSFSheet sheet = sXL.getSheet(endUrlSheet);
        if(sheet == null)
            return null;
        HashMap<String, Integer> map = new HashMap<>();
        readXLSHeader(endUrlSheet, map );
        String region = null;
        for(int rowIndex = 0; rowIndex <= sheet.getLastRowNum() ; rowIndex++){
            XSSFRow currentRow = sheet.getRow(rowIndex);
            if(currentRow == null)
                continue;
            int index = map.get("Service") ==null?-1:map.get("Service");
            if(index < 0){
                logNoHeadderError(endUrlSheet,"Service");
                return null;
            }
            XSSFCell serviceCell = currentRow.getCell(index);
            String service = sXL.getCellValue(serviceCell);
            if (service == null || service.isEmpty() || !service.equalsIgnoreCase(serviceName)) {
                continue;
            }
            region = HTML.properties.getProperty("Region");
            index = -1;
            index = map.get(region) ==null?-1:map.get(region);
            if(index < 0){
                logNoHeadderError(serviceName,"region");
                return null;
            }
            XSSFCell endUrlCell = currentRow.getCell(index);
            endUrl = sXL.getCellValue(endUrlCell)==null?null:String.valueOf(sXL.getCellValue(endUrlCell)).trim();
            break;
        }
        logger.info("the service end URL for the service :\""+serviceName+"\" for the region- \""+region+"\" is "+endUrl);
        return endUrl;
    }

    private void updateErrorInfo(String tcId , String errorMsg) throws Exception {
        String previousError = "";
        XSSFSheet sheet = null;
        try{
            String sheetName = "TestCase";
            sheet =     sXL.getSheet(sheetName);
            int rowCount = sheet.getLastRowNum() + 1;
            HashMap<String, Integer> sheetHeaddersMap = new HashMap<>();
            readXLSHeader(sheetName, sheetHeaddersMap );
            for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++) {
                XSSFRow currentRow = sheet.getRow(rowIndex);
                int index = sheetHeaddersMap.get("ID") ==null?-1:sheetHeaddersMap.get("ID");
                if(index < 0){
                    logNoHeadderError(sheetName,"ID");
                    return;
                }
                XSSFCell idCell = currentRow.getCell(index);
                String id = sXL.getCellValue(idCell);
                if(id != null && id.equalsIgnoreCase(tcId)){
                    index =-1;
                    index = sheetHeaddersMap.get("Desc") ==null?-1:sheetHeaddersMap.get("Desc");
                    if(index < 0){
                        logNoHeadderError(sheetName, "Desc");
                        return;
                    }
                    XSSFCell errorCell = currentRow.getCell(index);
                    previousError = sXL.getCellValue(errorCell);
                    if(errorMsg == null || errorMsg.isEmpty())
                        previousError = "";
                    if(previousError != null & !previousError.trim().isEmpty()){
                        previousError = previousError.trim() + "\n";
                    }
                    logger.info("error info : \""+previousError+errorMsg+"\" is updated in TestCases sheet at position [row:"+currentRow.getRowNum()+" , col:"+index+"]");
                    sXL.setCellData(sheet, previousError+errorMsg, currentRow.getRowNum(), index);
                    break;
                }
            }
        
        }catch(Exception e){
            e.printStackTrace();
        }
        
        
    }
    

    private boolean callSublevelWebService(String subservices, XSSFRow currentRow, String tcId) {

        try {
            if (subservices != null) {
                for (String sheetName : subservices.split(",")) {
                    XSSFSheet sheet = sXL.getSheet(sheetName);
                    if (sheet == null)
                        return false;
                    XSSFRow firstRow = sheet.getRow(0);
                    int lastCellNum = firstRow.getLastCellNum();
                    Map<Integer, String> ipMap = new HashMap<>();
                    Map<Integer, String> opMap = new HashMap<>();
                    Map<Integer, String> requestFileMap = new HashMap<>();
                    Map<Integer, String> dynamicIpOpRequired = new HashMap<>();
                    readXLSHeaderAndCreateMaps(firstRow, lastCellNum, ipMap,
                            opMap, requestFileMap, dynamicIpOpRequired);
                    int rowCount = sheet.getLastRowNum() + 1;
                    Map<String, String> ipXpathMap = new HashMap<>();
                    Map<String, String> opXpathMap = new HashMap<>();
                    getXpathsFromXls(sheet.getRow(1), ipMap,opMap,ipXpathMap,opXpathMap);
                    HashMap<String, Integer> sheetHeaddersMap = new HashMap<>();
                    readXLSHeader(sheetName, sheetHeaddersMap );
                    for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++) {
                        XSSFRow sublevelserviceCurrentRow = sheet.getRow(rowIndex);
                        if (sublevelserviceCurrentRow == null) {
                            break;
                        }
                        int headderIndex = sheetHeaddersMap.get("ID") ==null?-1:sheetHeaddersMap.get("ID");
                        if(headderIndex < 0){
                            logNoHeadderError(sheetName,"ID");
                            return false;
                        }
                        XSSFCell idCell = sublevelserviceCurrentRow.getCell(headderIndex);
                        String testCaseID = sXL.getCellValue(idCell);
                        if (testCaseID == null || testCaseID.isEmpty() || !testCaseID.equalsIgnoreCase(tcId)) {
                            continue;
                        }
                        
                        Document document = null;
                        String requestXmlPath = null;
                        String requestPayloadFile=null;
                        for (Integer index : requestFileMap.keySet()) {
                            XSSFCell cell = sublevelserviceCurrentRow.getCell(index);
                            String cellValue = sXL.getCellValue(cell);
                            if (cellValue != null && !cellValue.isEmpty()) {
                                requestXmlPath = "Request/" + cellValue
                                        + ".xml";
                                requestPayloadFile = String.valueOf(cellValue).trim();
                                document = getRequestDocument(cellValue);
                                break;
                            }
                        }
                        if (document == null) {
                            logger.error("request file not Exist. Path:"
                                    + requestXmlPath);
                            updateErrorInfo(tcId,"request file not Exist for the test case : "+tcId+"in sheet :"+sheetName);
                            return false;
                        }
                        if (document != null) {
                            boolean validateXpaths = validateXpaths(document , ipXpathMap);
                            if(!validateXpaths){
                                updateErrorInfo(tcId,"given xpaths in "+sheetName+" for Input and output parameters are wrong for the test case"+tcId);
                                logger.error("given xpaths in "+sheetName+" for Input and output parameters are wrong. Please correct it");
                                return validateXpaths;
                                    
                            }else{
                                logger.info("input parameters XPats in "+sheetName +" are valid and proceding to hit the web service");
                            }
                            if (dynamicIpOpRequired != null
                                    && !dynamicIpOpRequired.isEmpty()) {
                                XSSFCell ipOpRequiredCell = currentRow
                                        .getCell(getKeyByValue(
                                                dynamicIpOpRequired,
                                                "DynamicIpOpRequired"));
                                String ipOpRequired = sXL
                                        .getCellValue(ipOpRequiredCell);
                                if ("Yes".equalsIgnoreCase(ipOpRequired.trim())
                                        || "Y".equalsIgnoreCase(ipOpRequired
                                                .trim())) {
                                    checkIpOpconfigAndUpdateSheet(ipMap, opMap,
                                            tcId, sheetName, rowIndex);
                                }
                            }else{
                                logger.info("dynamic updation/generating values for input and output parameters is not available for the test case id:"+tcId +" in sheet \""+sheetName+"\"");
                            }
                            // Request parameters modification as per the excel
                            // sheet.
                            modifydocumentwithIPParams(ipMap, ipXpathMap , sublevelserviceCurrentRow, document);
                            logRequest(document);
                            savePayload(document,requestPayloadFile,tcId,"Request/sent");
                            String serviceUrlCellValue = getServiceEndUrl(sheetName);
                            if(serviceUrlCellValue == null || serviceUrlCellValue.isEmpty()){
                                updateErrorInfo(tcId,"Service End URL in cell index - \""+headderIndex+"\" is not present for the test case"+tcId+" in sheet "+sheetName);
                                logger.error("Service End URL in cell index - \"3\" is not present.PLease correct it");
                                return false;
                            }

                            // service call
                            String response = callWebService(requestXmlPath,
                                    document, serviceUrlCellValue, tcId);

                            // read the response
                            if(response == null){
                                logger.error("response from the main service is null. So terminating without checking the sub level services.");
                                updateErrorInfo(tcId,"null response retrieved for the service url "+serviceUrlCellValue+" and the test case "+tcId+" in sheet "+sheetName);
                                return false;
                            }else if (response != null) {
                                logger.info("the response from the service Url:" + serviceUrlCellValue+" is "+response);
                                Document responseDocument = convertStringToDocument(response);
                                savePayload(responseDocument,requestPayloadFile,tcId,"Response/received");
                                String faultXpath = "Envelope/Body/Fault";
                                expectedXPaths.add(faultXpath);
                                Node faultNode = getElementNodeByXPath(
                                        responseDocument, faultXpath);
                                if (faultNode == null) {
                                    logger.info("sublevel service :"+serviceUrlCellValue+" is passed for testcase id :"+tcId+" in the sheet "+sheetName);
                                    logger.info("updated Excel with output parameter values.");
                                    return true;

                                } else {
                                    
                                    updateErrorInfo(tcId,"fault response for the service url "+serviceUrlCellValue+" and the test case "+tcId+" in sheet "+sheetName);
                                    updateXlsWithFaultInfo(responseDocument, tcId);
                                    return false;
                                }
                            }
                        }
                    }
                }
            }
        }catch(Exception e){
            try {
                updateErrorInfo(tcId,"exception occured for the test case "+tcId+" in sheet "+subservices+". cause : "+e.getMessage());
            } catch (Exception e1) {
                e1.printStackTrace();
            }
            return false;
        }
    
        return false;
    }

    public void logNoHeadderError(String sheetName,String headderName) {
        logger.error("no headder with name \""+headderName+"\" in the sheet: "+sheetName);
    }

    private void checkIpOpconfigAndUpdateSheet(Map<Integer, String> ipMap,
            Map<Integer, String> opMap, String tcId, String service, int destRowIndex) throws Exception {
        logger.info("- started");
        String sheetName = "DynamicIpOpConfig";
        XSSFSheet DynamicIpOpConfigSheet = sXL.getSheet(sheetName);
        if (DynamicIpOpConfigSheet == null){
            logger.error("DynamicIpOpConfigSheet sheet is not available in the given Excel file.");
            return;
        }
        HashMap<String, Integer> ipOpHeaddersMap = new HashMap<>();
        readXLSHeader(sheetName, ipOpHeaddersMap );
        int rows = DynamicIpOpConfigSheet.getLastRowNum();
        Map<String, HashMap<String, Integer>> sheetMap= new HashMap<String, HashMap<String, Integer>>();
        HashMap<String, Integer> sheetRowMap = new HashMap<>();
        for(int rowIndex=0;rowIndex<=rows;rowIndex++){
            XSSFRow currentRow = DynamicIpOpConfigSheet.getRow(rowIndex);
            if (currentRow == null) {
                continue;
            }
            int headderIndex = ipOpHeaddersMap.get("Service") ==null?-1:ipOpHeaddersMap.get("Service");
            if(headderIndex < 0){
                logNoHeadderError(sheetName,"Service");
                return;
            }
            XSSFCell serviceNameCell = currentRow.getCell(headderIndex);
            String serviceName = sXL.getCellValue(serviceNameCell);
            if (serviceName == null || serviceName.isEmpty() || !serviceName.equalsIgnoreCase(service)) {
                continue;
            }
            headderIndex = -1;
            headderIndex = ipOpHeaddersMap.get("Destination Col Name") ==null?-1:ipOpHeaddersMap.get("Destination Col Name");
            if(headderIndex < 0){
                logNoHeadderError(sheetName,"Destination Col Name");
                return;
            }
            XSSFCell destinationHeadderCell = currentRow.getCell(headderIndex);
            String destinationHeadders = sXL.getCellValue(destinationHeadderCell);
            destinationHeadders = destinationHeadders.trim();
            
            //random number generaion logic
            headderIndex = -1;
            headderIndex = ipOpHeaddersMap.get("Format") ==null?-1:ipOpHeaddersMap.get("Format");
            if(headderIndex < 0){
                logNoHeadderError(sheetName,"Format");
                return;
            }
            XSSFCell formatterCell = currentRow.getCell(headderIndex);
            String format = sXL.getCellValue(formatterCell);
            format = format.trim();
            XSSFSheet destinationSheet = sXL.getSheet(service);
            if (format != null && !format.isEmpty()) {
                String value = "";
////                date~format
//              //date~YYYY-MM-dd'T'HH:mm:ss
//              //date~YYYY-MM-dd
//              //date~dd-MM-YYYY
//
//              if (format.contains("date")) {
//                  String[] d = format.split("~");
//                  String dFormat = "YYYY-MM-dd'T'HH:mm:ss";
//                  if (d.length > 1) {
//                      dFormat = d[1];
//
//                  }
//                  value = new SimpleDateFormat(dFormat).format(new Date());
//              }
                 
                if ("date".equalsIgnoreCase(format)) {
                    value = new SimpleDateFormat("YYYY-MM-dd'T'HH:mm:ss")
                            .format(new Date());
                } else {
                    //length
                    String[] speFormat = null;
                    String splitter = null;
                    if(format.contains(" ")){
                        splitter=" ";
                        speFormat = format.split(" ");
                    }else if(format.contains("-")){
                        splitter="-";
                        speFormat = format.split("-");
                    }else{
                        speFormat = new String[1];
                        speFormat[0] =format;
                        splitter="";
                    }
                    for(String s: speFormat){
                        String type = "";
                        int length = 0;
                        if(s.length() == 4){
                            type = s.substring(2);
                            length = Integer.parseInt(s.substring(0,2));
                        }
                        switch (type) {
                            case "AL":  //Alphabets
                                if (!value.isEmpty())
                                    value = value + splitter;
                                value = value
                                        + RandomStringUtils
                                                .randomAlphabetic(length);
                                break;
                            case "NU": //Numbers
                                if (!value.isEmpty())
                                    value = value + splitter;
                                value = value
                                        + RandomStringUtils.randomNumeric(length);
                                break;
                            case "AN": //Alpha Numeric
                                if (!value.isEmpty())
                                    value = value + splitter;
                                value = value
                                        + RandomStringUtils
                                                .randomAlphanumeric(length);
                                break;
                                

                        }
                    }
                    
                }
                value = value.trim().toUpperCase();
                for (String destinationHeadder : destinationHeadders
                        .split("\n")) {
                    Integer colIndex = getKeyByValue(ipMap, destinationHeadder);
                    if (colIndex == null) {
                        colIndex = getKeyByValue(opMap, destinationHeadder);
                    }
                    if (colIndex == null) {
                        logger.error(destinationHeadder
                                + " is not found in the sheet \"" + service);
                        continue;
                    }
                    sXL.setCellData(destinationSheet, value,
                            destRowIndex, colIndex);
                    logger.info("\"" + value + "\" is generated in format:\""
                            + format + "\" and updated in sheet \"" + service
                            + "\" at position [row:" + destRowIndex + " , col:"
                            + colIndex + " ]");
                }
            }
            
            //end
            
            
            
            //get value from the source sheet
            headderIndex = -1;
            headderIndex = ipOpHeaddersMap.get("SourceSheet") ==null?-1:ipOpHeaddersMap.get("SourceSheet");
            if(headderIndex < 0){
                logNoHeadderError(sheetName,"SourceSheet");
                return;
            }
            
            XSSFCell sourceSheetCell = currentRow.getCell(headderIndex);
            String sourceSheet = sXL.getCellValue(sourceSheetCell);
            if(sourceSheet == null || sourceSheet.trim().isEmpty())
                continue;
            if(sheetMap.get(sourceSheet) == null){
                HashMap<String, Integer> headdderMap = new HashMap<>();
                readXLSHeader(sourceSheet, headdderMap);
                sheetMap.put(sourceSheet, headdderMap);
            }
            HashMap<String, Integer> map = sheetMap.get(sourceSheet);
            headderIndex = -1;
            headderIndex = ipOpHeaddersMap.get("Source Column Name") ==null?-1:ipOpHeaddersMap.get("Source Column Name");
            if(headderIndex < 0){
                logNoHeadderError(sheetName,"Source Column Name");
                return;
            }
            XSSFCell sourceHeadderCell = currentRow.getCell(headderIndex);
            String sourceHeadder = sXL.getCellValue(sourceHeadderCell);
            //retrive value from the index
            XSSFSheet srcSheet = sXL.getSheet(sourceSheet);
            HashMap<String, Integer> sourcesheetHeaddersMap = new HashMap<>();
            readXLSHeader(sourceSheet, sourcesheetHeaddersMap );
            int srcRows = srcSheet.getLastRowNum();
            if (sheetRowMap.get(sourceSheet) == null) {
                for (int srcRowIndex = 0; srcRowIndex <= srcRows; srcRowIndex++) {
                    XSSFRow srcCurrentRow = srcSheet.getRow(srcRowIndex);
                    if (currentRow == null) {
                        break;
                    }
                    int srcheadderIndex = sourcesheetHeaddersMap.get("ID") ==null?-1:sourcesheetHeaddersMap.get("ID");
                    if(headderIndex < 0){
                        logNoHeadderError(sheetName,"ID");
                        return;
                    }
                    XSSFCell tcIdCell = srcCurrentRow.getCell(srcheadderIndex);
                    String srcTcId = sXL.getCellValue(tcIdCell);
                    if (srcTcId == null || srcTcId.isEmpty()
                            || !srcTcId.trim().equalsIgnoreCase(tcId)) {
                        continue;
                    }
                    sheetRowMap.put(sourceSheet, srcRowIndex);
                    break;
                }
            }
            int sourceRowIndex = sheetRowMap.get(sourceSheet) == null? -1 : sheetRowMap.get(sourceSheet);
            if(sourceRowIndex>=0){
                XSSFCell newValueCell = srcSheet.getRow(sourceRowIndex).getCell(map
                        .get(sourceHeadder));
                String newValue = sXL.getCellValue(newValueCell);
                if(newValue != null && !newValue.trim().isEmpty()){
                    
                    
                    for (String destinationHeadder : destinationHeadders
                            .split("\n")) {
                        Integer colIndex = getKeyByValue(ipMap,
                                destinationHeadder);
                        if (colIndex == null) {
                            colIndex = getKeyByValue(opMap, destinationHeadder);
                        }
                        if (colIndex == null) {
                            logger.error(destinationHeadder+" is not found in the sheet \""+service);
                            continue;
                        }
                        sXL.setCellData(destinationSheet, newValue.trim(),
                                destRowIndex, colIndex);
                        logger.info("\"" + newValue + "\" is retrieved from \""
                                + sourceSheet + "\" sheet at position [row:"
                                + sourceRowIndex + " , col:"
                                + map.get(sourceHeadder)
                                + " ] and updated in sheet \"" + service
                                + "\" at position [row:" + destRowIndex
                                + " , col:" + colIndex + " ]");
                    }
                }
            }
        }
        logger.info("- compleated");
    }
    
    public <T, E> T getKeyByValue(Map<T, E> map, E value) {
        for (Entry<T, E> entry : map.entrySet()) {
            if (value.equals(entry.getValue())) {
                return entry.getKey();
            }
        }
        return null;
}

    private void savePayload(Document document, String requestPayloadFile,
            String tcId,String rootDir) {
        final File dir = new File(rootDir);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        String timeStamp = new SimpleDateFormat("yyyy-MM-dd_hh-mm-ss").format(new Date());
        String fileName = tcId+"_"+requestPayloadFile+"_"+timeStamp+".xml";
        File f=new File(rootDir+"/"+fileName);
        try {
            FileWriter f2 = new FileWriter(f, false);
            f2.write(convertDocumentToString(document));
            f2.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        logger.info("the pay load uploaded into "+f.getAbsolutePath());
    }

    private boolean validateXpaths(Document document, Map<String, String> xpathMap) {
        boolean isXpathsValid = true;
        XPath xPath = XPathFactory.newInstance().newXPath();
        for(String path : xpathMap.keySet()){
            String xPathName = xpathMap.get(path);
            try {
                Node node = (Node) xPath.compile(xPathName).evaluate(document,
                        XPathConstants.NODE);
                if(node == null){
                    isXpathsValid = false;
                    logger.error("Xpath is invalid for the headder: "+path+". Please correct it. "+path+" header Xpath is  "
                            + xPathName);
                }
            } catch (XPathExpressionException e) {
                logger.error("in Catch : Xpath is invalid for the headder: "+path+". Please correct it. "+path+" header Xpath is  "
                            + xPathName);
                e.printStackTrace();
                isXpathsValid = false;
            }
        }
        return isXpathsValid;
    }

    private void getXpathsFromXls(XSSFRow row, Map<Integer, String> ipMap,
            Map<Integer, String> opMap, Map<String, String> ipXpathMap,
            Map<String, String> opXpathMap) {
        
        for(int ipIndex : ipMap.keySet()){
            XSSFCell cell = row.getCell(ipIndex);
            String cellValue = sXL.getCellValue(cell);
            
            cellValue= formatXpath(String.valueOf(cellValue).trim());
            ipXpathMap.put(ipMap.get(ipIndex), cellValue);
        }
        for(int opIndex : opMap.keySet()){
            XSSFCell cell = row.getCell(opIndex);
            String cellValue = sXL.getCellValue(cell);
            cellValue= formatXpath(String.valueOf(cellValue).trim());
            opXpathMap.put(opMap.get(opIndex), cellValue);
        }
        
    }
    
    // To remove the random generated values and op values
    private void setIpOpEmptyValues(XSSFSheet sheet, Map<Integer, String> ipMap,
            Map<Integer, String> opMap,String tcId,Map<Integer, String> dynamicIpOpRequired,Map<Integer, String> requestFileMap) {
        try {
            XSSFSheet sheet1 = sXL.getSheet("DynamicIpOpConfig");
            int rowCount = sheet.getLastRowNum();
            int rowCount1 = sheet1.getLastRowNum();
            int colCount=sXL.getColumnCount(sheet.getSheetName());
            int colCount1=sXL.getColumnCount(sheet1.getSheetName());
            
            String sheetName=sheet.getSheetName();
        
        //for(int ipIndex : ipMap.keySet()){
            for(int x=1;x<=rowCount;x++)
            {
                XSSFRow row1=sheet.getRow(x);
                if (row1==null)
                    continue;
                
                XSSFCell cell1 = row1.getCell(0);
                String testCaseID = sXL.getCellValue(cell1);
            
                if (testCaseID.contains(tcId)){
                    if (dynamicIpOpRequired != null
                            && !dynamicIpOpRequired.isEmpty())
                    {
                        int col=getKeyByValue(dynamicIpOpRequired,"DynamicIpOpRequired");
                        XSSFRow row3=sheet.getRow(x);
                        XSSFCell cell2 = row3.getCell(col);
                        String value = sXL.getCellValue(cell2);
                        if (value.contains("YES") || value.contains("Yes") || value.contains("Y")){
                            //XSSFRow row4=sheet1.getRow(1);
                            //XSSFCell cell4 = row4.getCell(0);
                            //if (sXL.getCellValue(cell4).contains("Service")){
                                for(int y=1;y<=rowCount1;y++)
                                {
                                    XSSFRow row5=sheet1.getRow(y);
                                    XSSFCell cell5 = row5.getCell(0);
                                    if (sXL.getCellValue(cell5).contains(sheetName))
                                    {
                                        //XSSFRow row6=sheet1.getRow(1);
                                        //XSSFCell cell6 = row6.getCell(3);
                                        //if (sXL.getCellValue(cell6).contains("Destination"))
                                        //{
                                            //XSSFRow row7=sheet1.getRow(y);
                                            XSSFCell cell7 = row5.getCell(3);
                                            String value1=sXL.getCellValue(cell7);
                                            String colName[]=value1.split("\n");
                                            for (int colNameCount=0;colNameCount<colName.length;colNameCount++)
                                            {
                                                //for (int colIndex=0;colIndex<=colCount;colIndex++)
                                                //{
                                                    //XSSFRow row8=sheet1.getRow(1);
                                                for(int ipIndex : ipMap.keySet()){
                                                    XSSFRow row8=sheet.getRow(0);
                                                    XSSFCell cell8 = row8.getCell(ipIndex);
                                                    String cellValue=sXL.getCellValue(cell8);
                                                    String value2=colName[colNameCount];
                                                    //System.out.println(value2);
                                                    if (cellValue.contains(value2))
                                                    {
                                                        sXL.setCellData(sheet, "", x, ipIndex);
                                                        //System.out.println("Identified and removed the Ip value "+value2+" for "+tcId+" from "+sheet.getSheetName()+" sheet");
                                                        logger.info("Identified and removed the Ip value "+value2+" for "+tcId+" from "+sheet.getSheetName()+" sheet");
                                                    }
                                                }
                                                //}
                                            //}
                                        }
                                    }
                                //}
                            }
                            
                            
                        }
                    }
                        
                        
                }
            }
            
            
        //}
        for(int opIndex : opMap.keySet()){
            for(int x=1;x<=rowCount;x++)
            {
                XSSFRow row1=sheet.getRow(x);
                XSSFRow row7=sheet.getRow(0);
                if (row1==null)
                    continue;
                
                XSSFCell cell1 = row1.getCell(0);
                String testCaseID = sXL.getCellValue(cell1);
                String opValue=sXL.getCellValue(row7.getCell(opIndex));
            
                if (testCaseID.contains(tcId)){
                        sXL.setCellData(sheet, "", x, opIndex);
                        //System.out.println("Identified and removed the op value "+opValue+" for "+tcId+" from "+sheet.getSheetName()+" sheet");
                        logger.info("Identified and removed the op value "+opValue+" for "+tcId+" from "+sheet.getSheetName()+" sheet");
                }
            }
        }
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
    

    private static String formatXpath(String xPath) {
        if(xPath != null){
            if(xPath.contains(":")){
                String[] paths = xPath.split("/");
                StringBuffer sb = new StringBuffer();
                for(String path : paths){
                    if(path.contains(":")){
                        String[] pathWOSc = path.split(":");
                        if(sb.length() > 0){
                            sb.append("/");
                        }
                        sb.append(pathWOSc[1]);
                    }else{
                        if(sb.length() > 0){
                            sb.append("/");
                        }
                        sb.append(path);
                    }
                }
                return sb.toString();
            }else{
                return xPath;
            }
        }
        return null;
    }

    /**
     * @param firstRow
     * @param lastCellNum
     * @param ipMap
     * @param opMap
     * @param requestFileMap
     * @param dynamicIpOpRequired
     */
    private void readXLSHeaderAndCreateMaps(XSSFRow firstRow, int lastCellNum,
            Map<Integer, String> ipMap, Map<Integer, String> opMap,
            Map<Integer, String> requestFileMap, Map<Integer, String> dynamicIpOpRequired) {
        for (int index = 0; index <= lastCellNum; index++) {
            XSSFCell cell = firstRow.getCell(index);
            String cellValue = sXL.getCellValue(cell);
            if (cellValue != null) {
                if (cellValue.contains("ip_")) {
                    ipMap.put(cell.getColumnIndex(), String.valueOf(cellValue).trim());
                } else if (cellValue.contains("op_")) {
                    opMap.put(cell.getColumnIndex(), String.valueOf(cellValue).trim());
                } else if (cellValue.contains("funServiceXML")) {
                    requestFileMap.put(cell.getColumnIndex(),
                            "requestXml");
                }else if(cellValue.contains("DynamicIpOpRequired")) {
                    dynamicIpOpRequired.put(cell.getColumnIndex(),
                            String.valueOf(cellValue).trim());
                }
            }
        }
    }
    
    
    private void readXLSHeader(String sheetName , HashMap<String, Integer> map) throws Exception {
        XSSFSheet sheet = sXL.getSheet(sheetName);
        if(sheet == null){
            logger.error(sheetName+" is not present in excel file");
        }
            
        XSSFRow firstRow = sheet.getRow(0);
        int lastCellNum = firstRow.getLastCellNum();
        for (int index = 0; index < lastCellNum; index++) {
            XSSFCell cell = firstRow.getCell(index);
            if(cell == null)
                continue;
            String cellValue = sXL.getCellValue(cell);
            if (cellValue != null) {
                    map.put(String.valueOf(cellValue).trim() , cell.getColumnIndex());
            }
        }
    }

    private void modifydocumentwithIPParams(Map<Integer, String> ipMap, Map<String, String> ipXpathMap,
             XSSFRow currentRow, Document document)
            throws Exception {
        for (Integer index : ipMap.keySet()) {
            XSSFCell cell = currentRow.getCell(index);
            String cellValue = sXL.getCellValue(cell);
            if (cellValue != null) {
                String xPath = ipXpathMap.get(String.valueOf(ipMap.get(index)).trim());
                if (xPath != null) {
                    setElementValueByXPath(document, xPath, String.valueOf(cellValue).trim());
                    logger.info("value updated in the document for the xPath :"+xPath+" and the values is "+cellValue);
                }else{
                    logger.error("xPath not found for the "+ipMap.get(index));
                }
            }
        }
    }

    private String callWebService(String requestXmlPath, Document document, String strURL, String tcId)
            throws FileNotFoundException, IOException, ClientProtocolException {
        String documentToString = convertDocumentToString(document);
        File input = convertStringToFile(documentToString,
                requestXmlPath);
        HttpPost post = null;
        HttpClient httpclient = null;
        String response = null;
        // Execute request
        try {
            post = new HttpPost(strURL);
            post.setEntity(new InputStreamEntity(
                    new FileInputStream(input), input.length()));
            post.setHeader("Content-type",
                    "text/xml; charset=ISO-8859-1");
            httpclient = HttpClientBuilder.create().build();
            HttpResponse httpResponse = httpclient.execute(post);
            if (httpResponse != null) {
                // Display status code
                StatusLine statusLine = httpResponse.getStatusLine();
                if(statusLine != null){
                    logger.info("Response status code: "+ httpResponse.getStatusLine());
                    if(httpResponse.getStatusLine().getStatusCode() != 200){
                        updateErrorInfo(tcId, "Response Status code:"+statusLine.getStatusCode() + " and Reason Phrase: "+statusLine.getReasonPhrase());
                    }
                }

                BufferedReader rd = new BufferedReader(
                        new InputStreamReader(httpResponse.getEntity().getContent()));
                StringBuffer result = new StringBuffer();
                String line = "";
                while ((line = rd.readLine()) != null) {
                    result.append(line);
                }
                response = result.toString();
            }
        } catch(Exception e ){
            response = null;
            logger.error("exception occured while consuming the service : "+strURL);
            e.printStackTrace();
        }finally {
            if (post != null)
                post.releaseConnection();
        }
        return response;
    }

    

    public Document getRequestDocument(String requestXmlName)
            throws Exception {
        Scanner scanner = null;
        try {
            if(requestXmlName == null || requestXmlName.isEmpty()){
                logger.error("Service request xml file : "+requestXmlName+" is not present.PLease correct it");
                return null;
            }
            logger.info("reuest payload created by " +"Request/" + requestXmlName.trim() + ".xml");
            scanner = new Scanner(
                    new File("Request/" + requestXmlName.trim() + ".xml"));
            String xml = scanner.useDelimiter("\\Z").next();
            return convertStringToDocument(xml);
        } catch(Exception e){
            logger.error("error occured while converting to document." + e.getMessage());
            return null;
        }finally {
            if (scanner != null)
                scanner.close();
        }
    }

    public Document convertStringToDocument(String xmlStr) {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder;
        try {
            builder = factory.newDocumentBuilder();
            Document doc = builder.parse(new InputSource(new StringReader(
                    xmlStr)));
            return doc;
        } catch (Exception e) {
            logger.error("Converstion failed from string to Document");
            e.printStackTrace();
        }
        return null;
    }

    

    
    
    public Node getElementNodeByXPath(Document document, String path)
            throws Exception {
        XPath xPath = XPathFactory.newInstance().newXPath();
        Node node = (Node) xPath.compile(path).evaluate(document,
                XPathConstants.NODE);
        if (node == null) {
            if(!expectedXPaths.contains(path)){
                logger.error("given Xpath is invalid. Please correct it. given Xpath is  "
                        + path);
                throw new Exception("Invalid Xpath : " + path);
            }
        }
        return node;

    }
    
    public String getElementValueByXPath(Document document, String path) throws Exception {
        Node node = getElementNodeByXPath(document,path);
        return node == null ? null : node.getTextContent();
    }
    

    public Document setElementValueByXPath(Document document, String path,
            String newContent) throws Exception {
        Node node = getElementNodeByXPath(document,path);
        if(node != null)
            node.setTextContent(newContent);
        return document;

    }

    public String convertDocumentToString(Document doc) {
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer transformer;
        try {
            transformer = tf.newTransformer();
            StringWriter writer = new StringWriter();
            transformer.transform(new DOMSource(doc), new StreamResult(writer));
            String output = writer.getBuffer().toString();
            return output;
        } catch (TransformerException e) {
            logger.error("converstion fail from document to String");
            e.printStackTrace();
        }

        return null;
    }

    public File convertStringToFile(String fileContent,
            String sFilePath) {
        File fold = new File(sFilePath);
        fold.delete();
        File fnew = new File(sFilePath);
//      logger.info(fileContent);

        try {
            FileWriter f2 = new FileWriter(fnew, false);
            f2.write(fileContent);
            f2.close();
        } catch (IOException e) {
            logger.error("converstion fail from String to File");
            e.printStackTrace();
        }
        return fnew;
    }
    
    

}
