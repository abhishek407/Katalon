package newpackage

import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject

import java.text.DateFormat
import com.relevantcodes.extentreports.ExtentReports
import com.relevantcodes.extentreports.ExtentTest
import com.relevantcodes.extentreports.LogStatus

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.checkpoint.CheckpointFactory
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testcase.TestCaseFactory
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testdata.TestDataFactory
import com.kms.katalon.core.testobject.ObjectRepository
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords

import java.text.SimpleDateFormat

import internal.GlobalVariable

import MobileBuiltInKeywords as Mobile
import WSBuiltInKeywords as WS
import WebUiBuiltInKeywords as WebUI

public class newkeyword 
{
   
   public static String filepath;
   public static ExtentReports extent;
   public static ExtentTest test;
   
	
	@Keyword
	def keywordName() {
	  Date date = new Date();
	  DateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
	  DateFormat timeFormat = new SimpleDateFormat("HH-mm-ss");
	  String onlydate= dateFormat.format(date);
	  String onlytime= timeFormat.format(date);
	  filepath = "D:\\Katalon_CAPPM\\" +"Results\\"+onlydate+"\\"+onlytime;
	  System.setProperty("logfilename", filepath);
	  extent = new ExtentReports (filepath+"\\ExtendsReport.html",true);
	  extent
        .addSystemInfo("Host Name", "Excers Inc")
        .addSystemInfo("Environment", "CAPPM Adapater POC")
        .addSystemInfo("User Name", "Excers");
	  extent.loadConfig(new File("D:\\Katalon_CAPPM"+"\\extent-config.xml"));
	  /*test = extent.startTest("Test of Create Client");
	  test.log(LogStatus.INFO, "Browser Launched");
	  test.log(LogStatus.INFO, "Navigated to www.techbeamers.com");
	  test.log(LogStatus.INFO, "Browser closed");
	  extent.endTest(test);
	  extent.flush();*/
	}
}
