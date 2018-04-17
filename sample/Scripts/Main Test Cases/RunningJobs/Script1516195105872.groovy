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

if(GlobalVariable.RunJob)
{
extentTest1 = extent1.startTest("Running Jobs TestCase");
String screenshot_path = GlobalVariable.screenshot_path;
WebElement home_page_1 = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/span_Home'), 
    3)

applyhighlight(home_page_1)

WebUI.click(findTestObject('Page_CA PPM  Overview General/span_Home'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Page_CA PPM  Overview General/span_ppm_nav_app_menu'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

WebElement jobs_page_click = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  Overview General (1)/a_Jobs'), 
    3)

applyhighlight(jobs_page_click)
WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Page_CA PPM  Overview General (1)/a_Jobs'), FailureHandling.CONTINUE_ON_FAILURE)



//WebUI.takeScreenshot(screenshot_path + '\\Jobs_Page.png')
removehighlight(home_page_1)

removehighlight(jobs_page_click)

String names_of_jobs = GlobalVariable.all_job_names;

ts = String.valueOf(System.currentTimeMillis())

String[] all_job_name = names_of_jobs.split(',')

List<String> all_job_names_list = new ArrayList<String>()

for (String a : all_job_name) {
    all_job_names_list.add(a)
}

println(all_job_names_list)
WebUI.delay(2)
extentTest1.log(LogStatus.INFO, "Entered into Jobs Page");
String screenShotPathjob = capture("JobsIntialPage1_homepage");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathjob));

for (def job_names : all_job_names_list) {
    WebUI.delay(2)

    WebElement input_job_name = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  Jobs Available Jobs/input_job_name'), 
        3)

    applyhighlight(input_job_name)

    WebUI.setText(findTestObject('Page_CA PPM  Jobs Available Jobs/input_job_name'), job_names)

    WebUI.delay(2)

    WebElement button_Filter = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  Jobs Available Jobs/button_Filter'), 
        3)

    applyhighlight(button_Filter)

    WebUI.click(findTestObject('Page_CA PPM  Jobs Available Jobs/button_Filter'))

    WebUI.takeScreenshot(screenshot_path + '\\'+job_names+'1.png')
	
	extentTest1.log(LogStatus.INFO, "JobName is Selected "+job_names);
	String screenShotPathJob_1 = capture("JobisSelected "+job_names);
	extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathJob_1));
    WebUI.delay(2)
	

    WebElement Job_click_paths = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/Page_CA PPM  Jobs Available Jobs/Job_click_paths'), 
        3)

    applyhighlight(Job_click_paths)

    WebUI.click(findTestObject('Object Repository/Page_CA PPM  Jobs Available Jobs/Job_click_paths'))

    WebUI.delay(2)

    WebElement input_job_name_apply = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/input_job_name'), 
        3)

    applyhighlight(input_job_name_apply)

    WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/input_job_name'), 5).clear()

    WebUI.delay(2)

    WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/input_job_name'), 5).sendKeys(
        job_names + ts)

    WebElement button_Submit_1 = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  Job Type Post Transact/button_Submit'), 
        3)

    applyhighlight(button_Submit_1)

    WebUI.delay(2)

    WebUI.click(findTestObject('Page_CA PPM  Job Type Post Transact/button_Submit'))
	extentTest1.log(LogStatus.INFO, "Job is Invoked "+job_names);
	String screenShotPathJob_2 = capture("JobisInvoked "+job_names);
	extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathJob_2));

    WebUI.takeScreenshot(screenshot_path + '\\'+job_names+'2.png')
    WebUI.delay(2)

    WebElement input_job_name_high = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/input_job_name'), 
        3)

    applyhighlight(input_job_name_high)

    WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/input_job_name'), 5).clear()

    WebUI.delay(2)

    WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/input_job_name'), 5).sendKeys(
        job_names + ts)

    WebUI.takeScreenshot(screenshot_path + '\\'+job_names+'3.png')
    name = true
	long startTime = System.currentTimeMillis();
    while (name && (System.currentTimeMillis()-startTime)<600000) 
	{
        WebUI.delay(5)

        WebElement button_Filter_high = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Jobs Scheduled Jobs/button_Filter'), 
            3)

        applyhighlight(button_Filter_high)

        WebUI.click(findTestObject('CAPPM/Page_CA PPM  Jobs Scheduled Jobs/button_Filter'))

        WebElement GetState = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/GetState'), 
            3)

        applyhighlight(GetState)

        var = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/GetState'), 
            3).getText()

        if (var.contentEquals('Completed')) {
			extentTest1.log(LogStatus.INFO, "Job is Completed "+job_names);
			String screenShotPathJob_4 = capture("JobisCompleted "+job_names);
			extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathJob_4));
            name = false
        }
    }
    
    WebUI.delay(1)

    WebElement a_Jobs_apply = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  Jobs Log/a_Jobs'), 3)

    applyhighlight(a_Jobs_apply)

    WebUI.click(findTestObject('Page_CA PPM  Jobs Log/a_Jobs'))

    WebUI.delay(1)

    WebElement a_Available = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  Jobs Log/a_Available Jobs'), 
        3)

    applyhighlight(a_Available)

    WebUI.click(findTestObject('Page_CA PPM  Jobs Log/a_Available Jobs'))

    WebUI.delay(2)
}
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