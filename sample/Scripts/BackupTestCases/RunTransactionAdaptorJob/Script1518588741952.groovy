import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject

import java.io.IOException
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



ExtentReports extent1 = newkeyword.extent;
ExtentTest extentTest1 = newkeyword.test;

extentTest1 = extent1.startTest("RunTransactionAdaptor Job TestCase");
try {

String click_path = GlobalVariable.click_path;

String Job_name = GlobalVariable.Job_name;

String screenshot_path = GlobalVariable.screenshot_path;

WebElement home_page = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/span_Home'), 
    3)

applyhighlight(home_page)

WebUI.click(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/span_Home'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.delay(2, FailureHandling.CONTINUE_ON_FAILURE)

WebElement jobs_page_click = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Overview General/a_Jobs'), 
    3)

applyhighlight(jobs_page_click)


WebUI.delay(2, FailureHandling.CONTINUE_ON_FAILURE)
removehighlight(home_page)
WebUI.scrollToElement(findTestObject('CAPPM/Page_CA PPM  Overview General/a_Jobs'), 0, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('CAPPM/Page_CA PPM  Overview General/a_Jobs'), FailureHandling.CONTINUE_ON_FAILURE)


WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)


removehighlight(jobs_page_click)

extentTest1.log(LogStatus.INFO, "Entered into Jobs Page");
String screenShotPathTxn = capture("JobsIntialPage_home");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathTxn));

WebElement input_job_name = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  Jobs Available Jobs/input_job_name'), 
    3)

applyhighlight(input_job_name)


WebUI.setText(findTestObject('Page_CA PPM  Jobs Available Jobs/input_job_name'), click_path, FailureHandling.CONTINUE_ON_FAILURE)

WebElement click_on_filter = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  Jobs Available Jobs/button_Filter'), 
    3)

applyhighlight(click_on_filter)

WebUI.click(findTestObject('Page_CA PPM  Jobs Available Jobs/button_Filter'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.takeScreenshot(screenshot_path + '\\'+click_path+'1.png')
WebElement Job_click_paths = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/Page_CA PPM  Jobs Available Jobs/Job_click_paths'), 
    3)

applyhighlight(Job_click_paths)

WebUI.click(findTestObject('Object Repository/Page_CA PPM  Jobs Available Jobs/Job_click_paths'), FailureHandling.CONTINUE_ON_FAILURE)

extentTest1.log(LogStatus.INFO, "JobName is Selected");
String screenShotPathTxn_1 = capture("JobisSelected");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathTxn_1));
WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.delay(2, FailureHandling.CONTINUE_ON_FAILURE)

ts = String.valueOf(System.currentTimeMillis())

WebElement new_job_name = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/input_job_name'), 
    3)

applyhighlight(new_job_name)

WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/input_job_name'), 0).clear()

WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/input_job_name'), 0).sendKeys(
    Job_name + ts)

WebElement input_INSTANCE_CODE = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/input_INSTANCE_CODE'), 
    3)

applyhighlight(input_INSTANCE_CODE)

WebUI.setText(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/input_INSTANCE_CODE'), 'Transaction', FailureHandling.CONTINUE_ON_FAILURE)

WebUI.delay(2, FailureHandling.CONTINUE_ON_FAILURE)

WebElement button_Submit = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/button_Submit'), 
    3)

applyhighlight(button_Submit)

WebUI.scrollToElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/button_Submit'), 0, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/button_Submit'), FailureHandling.CONTINUE_ON_FAILURE)

extentTest1.log(LogStatus.INFO, "Job is Invoked");
String screenShotPathTxn_2 = capture("JobisInvoked");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathTxn_2));

WebElement fliter_new_job_name = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Jobs Scheduled Jobs/input_job_name'), 
    3)

applyhighlight(fliter_new_job_name)

WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Jobs Scheduled Jobs/input_job_name'), 30).sendKeys(Job_name + 
    ts)

name = true

while (name) {
    WebUI.delay(5)

    WebElement button_Filter_new = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Jobs Scheduled Jobs/button_Filter'), 
        3)

    applyhighlight(button_Filter_new)

    WebUI.click(findTestObject('CAPPM/Page_CA PPM  Jobs Scheduled Jobs/button_Filter'))

    WebElement GetState = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/GetState'), 
        3)

    applyhighlight(GetState)

    var = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/GetState'), 
        3).getText()

    if (var.contentEquals('Completed')) 
	{
		extentTest1.log(LogStatus.INFO, "Job is Completed");
		String screenShotPathTxn_3 = capture("JobisCompleted");
		extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathTxn_3));
        name = false
    }
	else if(var.contentEquals('Failed')) 
	{
		extentTest1.log(LogStatus.FAIL, "Job got failed");
		String screenShotPathTxn_fail = capture("Jobisfailed");
		extentTest1.log(LogStatus.FAIL, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathTxn_fail));
		name = false
		throw new Exception();
	}
}
}
catch(Exception e)
{
	WebUI.closeBrowser();
}
finally
{
extent1.endTest(extentTest1);
extent1.flush();
}
void applyhighlight(WebElement element) {
    WebDriver driver = DriverFactory.getWebDriver()

    JavascriptExecutor js = ((driver) as JavascriptExecutor)

    js.executeScript('arguments[0].style.border=\'3px solid red\'', element)
}

void removehighlight(WebElement element) {
    WebDriver driver = DriverFactory.getWebDriver()

    JavascriptExecutor js = ((driver) as JavascriptExecutor)

    js.executeScript('arguments[0].style.border=\'\'', element)
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