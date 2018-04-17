import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import java.nio.file.Files as Files
import java.nio.file.Path as Path
import java.nio.file.Paths as Paths
import java.nio.file.StandardCopyOption as StandardCopyOption
import java.text.DecimalFormat as DecimalFormat
import java.text.SimpleDateFormat as SimpleDateFormat
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
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import org.stringtemplate.v4.compiler.STParser.namedArg_return as namedArg_return
import com.kms.katalon.core.testdata.ExcelData as ExcelData
import com.kms.katalon.core.testdata.InternalData as InternalData
import com.kms.katalon.core.webui.common.WebUiCommonHelper as WebUiCommonHelper
import javax.swing.JOptionPane as JOptionPane
import org.apache.poi.hssf.usermodel.HSSFSheet as HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook as HSSFWorkbook
import org.apache.poi.openxml4j.exceptions.InvalidFormatException as InvalidFormatException
import org.apache.poi.ss.usermodel.Cell as Cell
import org.apache.poi.ss.usermodel.Row as Row
import org.apache.poi.ss.usermodel.Sheet as Sheet
import org.apache.poi.ss.usermodel.Workbook as Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory as WorkbookFactory
import org.junit.After as After
import org.openqa.selenium.By as By
import org.openqa.selenium.JavascriptExecutor as JavascriptExecutor
import org.openqa.selenium.WebDriver as WebDriver
import org.openqa.selenium.WebElement as WebElement
import com.kms.katalon.core.webui.driver.DriverFactory as DriverFactory

if(GlobalVariable.RunJob)
{
String screenshot_path = GlobalVariable.screenshot_path;

WebUI.click(findTestObject('Page_CA PPM  Overview General/span_Home'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Page_CA PPM  Overview General/span_ppm_nav_app_menu'), FailureHandling.CONTINUE_ON_FAILURE)

//WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Page_CA PPM  Overview General (1)/a_Jobs'), FailureHandling.CONTINUE_ON_FAILURE)

//WebUI.takeScreenshot(screenshot_path + '\\Jobs_Page.png')

String names_of_jobs = GlobalVariable.all_job_names;

ts = String.valueOf(System.currentTimeMillis())

String[] all_job_name = names_of_jobs.split(',')

List<String> all_job_names_list = new ArrayList<String>()

for (String a : all_job_name) {
    all_job_names_list.add(a)
}

println(all_job_names_list)

for (def job_names : all_job_names_list) {
	
    WebUI.delay(2)

    WebUI.setText(findTestObject('Page_CA PPM  Jobs Available Jobs/input_job_name'), job_names)

    //WebUI.delay(2)

    WebUI.click(findTestObject('Page_CA PPM  Jobs Available Jobs/button_Filter'))

    WebUI.takeScreenshot(screenshot_path + '\\'+job_names+'1.png')
    //WebUI.delay(2)

    WebUI.click(findTestObject('Object Repository/Page_CA PPM  Jobs Available Jobs/Job_click_paths'))

    WebUI.delay(2)


    WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/input_job_name'), 5).clear()

    WebUI.delay(2)

    WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/input_job_name'), 5).sendKeys(
        job_names + ts)


    //WebUI.delay(2)

    WebUI.click(findTestObject('Page_CA PPM  Job Type Post Transact/button_Submit'))

    WebUI.takeScreenshot(screenshot_path + '\\'+job_names+'2.png')
   // WebUI.delay(2)

    WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/input_job_name'), 5).clear()

    WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/input_job_name'), 5).sendKeys(
        job_names + ts)

    WebUI.takeScreenshot(screenshot_path + '\\'+job_names+'3.png')
    name = true

    while (name) {
        //WebUI.delay(5)

        WebUI.click(findTestObject('CAPPM/Page_CA PPM  Jobs Scheduled Jobs/button_Filter'))

        var = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/GetState'), 
            3).getText()

        if (var.contentEquals('Completed')) {
            name = false
        }
    }
    
    //WebUI.delay(1)

    WebUI.click(findTestObject('Page_CA PPM  Jobs Log/a_Jobs'))

    //WebUI.delay(1)

    WebUI.click(findTestObject('Page_CA PPM  Jobs Log/a_Available Jobs'))

    //WebUI.delay(2)
}

}

