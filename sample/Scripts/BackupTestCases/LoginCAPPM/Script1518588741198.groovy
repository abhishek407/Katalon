import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths
import java.nio.file.StandardCopyOption
import java.text.DecimalFormat
import java.text.SimpleDateFormat

import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.checkpoint.CheckpointFactory as CheckpointFactory
import com.kms.katalon.core.logging.KeywordLogger as KeywordLogger
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as MobileBuiltInKeywords
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testcase.TestCaseFactory as TestCaseFactory
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testdata.TestDataFactory as TestDataFactory
import com.kms.katalon.core.testobject.ObjectRepository as ObjectRepository
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WSBuiltInKeywords
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUiBuiltInKeywords
import com.kms.katalon.core.webui.keyword.internal.WebUIAbstractKeyword as WebUIAbstractKeyword
import com.relevantcodes.extentreports.ExtentTest
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import internal.GlobalVariable as GlobalVariable
import newpackage.newkeyword

import org.openqa.selenium.Keys as Keys
import org.openqa.selenium.OutputType
import org.openqa.selenium.TakesScreenshot
import org.stringtemplate.v4.compiler.STParser.namedArg_return as namedArg_return
import com.kms.katalon.core.testdata.ExcelData as ExcelData
import com.kms.katalon.core.testdata.InternalData as InternalData
import com.kms.katalon.core.webui.common.WebUiCommonHelper as WebUiCommonHelper
import javax.swing.JOptionPane as JOptionPane

import org.apache.commons.io.FileUtils
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.openxml4j.exceptions.InvalidFormatException
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.junit.After as After
import org.openqa.selenium.By as By
import org.openqa.selenium.JavascriptExecutor
import org.openqa.selenium.WebDriver as WebDriver
import org.openqa.selenium.WebElement as WebElement
import com.kms.katalon.core.webui.driver.DriverFactory as DriverFactory
import com.relevantcodes.extentreports.ExtentReports
import com.relevantcodes.extentreports.ExtentTest
import com.relevantcodes.extentreports.LogStatus

WebUI.callTestCase(findTestCase('Main Test Cases/ConvertCSVtoXls'), [:])
WebUI.callTestCase(findTestCase('Test Cases/Main Test Cases/GetDataFromExcel'), [:])

CustomKeywords.'newpackage.newkeyword.keywordName'()

ExtentReports extent1 = newkeyword.extent;
ExtentTest extentTest1 = newkeyword.test;
   
extentTest1 = extent1.startTest("LoginPage TestCase");

WebUI.openBrowser('', FailureHandling.CONTINUE_ON_FAILURE)

WebUI.maximizeWindow(FailureHandling.CONTINUE_ON_FAILURE)

WebUI.navigateToUrl(GlobalVariable.open_url, FailureHandling.CONTINUE_ON_FAILURE)

extentTest1.log(LogStatus.INFO, "Browser Launched");

WebElement  username = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Login/input_userName'), 3)

applyhighlight(username);
WebUI.setText(findTestObject('CAPPM/Page_CA PPM  Login/input_userName'), GlobalVariable.input_username, FailureHandling.CONTINUE_ON_FAILURE)
extentTest1.log(LogStatus.INFO, "Username is entered");
WebElement  password = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Login/input_passWord'), 3)

applyhighlight(password);

WebUI.setText(findTestObject('CAPPM/Page_CA PPM  Login/input_passWord'), GlobalVariable.input_password, FailureHandling.CONTINUE_ON_FAILURE)
extentTest1.log(LogStatus.INFO, "Password is entered");
WebElement  buttonclick = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Login/input_ppm_login_button'), 3)


String screenShotPathLogin = capture("LoginPage");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathLogin));
applyhighlight(buttonclick);

WebUI.delay(2)
WebUI.click(findTestObject('CAPPM/Page_CA PPM  Login/input_ppm_login_button'), FailureHandling.CONTINUE_ON_FAILURE)


WebUI.delay(3)
WebElement home_page1;
try
{
 home_page1= WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/span_Home'),1)
}
catch (Exception e) 
{
	extentTest1.log(LogStatus.FAIL, "Login is Failed");
	extent1.endTest(extentTest1);
	extent1.flush();
	WebUI.closeBrowser()
}
boolean login_sucess = home_page1.isEnabled() && home_page1.isDisplayed();
println login_sucess
if(login_sucess)
{
extentTest1.log(LogStatus.PASS, "Login is Successfull");
}
else
{
extentTest1.log(LogStatus.FAIL, "Login is Failed");
}
extent1.endTest(extentTest1);
extent1.flush();
public void applyhighlight(WebElement element)
{
	WebDriver driver = DriverFactory.getWebDriver()
	JavascriptExecutor js=(JavascriptExecutor)driver;
	js.executeScript("arguments[0].style.border='3px solid red'", element);
}

public String capture(String screenShotName) throws IOException
{
	WebDriver driver = DriverFactory.getWebDriver()
	TakesScreenshot ts = (TakesScreenshot)driver;
	File source = ts.getScreenshotAs(OutputType.FILE);
	String dest = "D:\\Katalon_CAPPM\\Results\\screenshots\\"+screenShotName+".png";
	File destination = new File(dest);
	FileUtils.copyFile(source, destination);
				 
	return dest;
}