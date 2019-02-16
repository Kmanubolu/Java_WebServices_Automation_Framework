package com.pc.screen;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import org.apache.commons.io.FileUtils;
import org.apache.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import com.pc.constants.PCConstants;
import com.pc.utilities.Common;
import com.pc.utilities.CommonManager;
import com.pc.utilities.HTML;
import com.pc.utilities.ManagerDriver;
import com.pc.utilities.PCThreadCache;
import com.pc.utilities.XlsxReader;

public class SCRCommon {

    public static String sheetname = "SCRCommon";
    static Logger logger =Logger.getLogger(sheetname);
    public static String Path;

    public static boolean JavaScriptDynamicWait(WebElement sElement, JavascriptExecutor js) throws Exception
    {
        boolean status = false;
            for (int i = 1; i <= Integer.parseInt(HTML.properties.getProperty("VERYLONGWAIT")); i++) {
                logger.info("Document Ajax State = "
                              + js.executeScript(
                                           "return Ext.Ajax.isLoading();")
                                           .toString());
                Boolean isAjaxRunning = Boolean.valueOf(js
                              .executeScript(
                                           "return Ext.Ajax.isLoading();") //returns true if ajax call is currently in progress
                              .toString());
                if (!isAjaxRunning.booleanValue()) {
                    status = true;
                       break;
                }
                Thread.sleep(1000);//wait for one secnod then check if ajax is completed
            }
         WebDriverWait wait = new WebDriverWait(ManagerDriver.getInstance().getWebDriver(),
                       Integer.parseInt(HTML.properties
                                     .getProperty("VERYLONGWAIT")));
         wait.until(ExpectedConditions.stalenessOf(sElement));// (By.id(readAttriID1)));
//       logger.info("End Wait....2");
     return status;
    }
    
    public static boolean JavaScript(JavascriptExecutor js) throws Exception
    {
        boolean status = false;
            for (int i = 1; i <= Integer.parseInt(HTML.properties.getProperty("VERYLONGWAIT")); i++) {
                logger.info("Document Ajax State = "
                              + js.executeScript(
                                           "return Ext.Ajax.isLoading();")
                                           .toString());
                Boolean isAjaxRunning = Boolean.valueOf(js
                              .executeScript(
                                           "return Ext.Ajax.isLoading();") //returns true if ajax call is currently in progress
                              .toString());
                if (!isAjaxRunning.booleanValue()) {
                    status = true;
                       break;
                }
                Thread.sleep(1000);//wait for one secnod then check if ajax is completed
            }
        return status;
    }
    
    public static boolean waitForJQueryProcessing(int timeOutInSeconds) {
        boolean jQcondition = false;
        try {
            new WebDriverWait(ManagerDriver.getInstance().getWebDriver(), timeOutInSeconds) {
            }.until(new ExpectedCondition<Boolean>() {

                @Override
                public Boolean apply(WebDriver driverObject) {
                    return (Boolean) ((JavascriptExecutor) driverObject)
                            .executeScript("return !!window.jQuery && window.jQuery.active == 0");
                }
            });
            jQcondition = (Boolean) ((JavascriptExecutor) ManagerDriver.getInstance().getWebDriver())
                    .executeScript("return window.jQuery != undefined && jQuery.active === 0");
            return jQcondition;
        } catch (Exception e) {
            e.printStackTrace();
           logger.error("Thread ID = " + Thread.currentThread().getId() + " Error Occured =" +e.getMessage(), e);
        }
        return jQcondition;
    }

}
