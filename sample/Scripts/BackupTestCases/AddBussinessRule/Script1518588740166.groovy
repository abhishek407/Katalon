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

extentTest1 = extent1.startTest("Add BussinessRules TestCase");

KeywordLogger log = new KeywordLogger()


/*WebElement home_page_add = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/span_Home'),
	3)

applyhighlight(home_page_add)

WebUI.click(findTestObject('Page_CA PPM  Overview General/span_Home'))

WebUI.delay(2)

WebElement CA_PPM_Adapter_Config_List_rule = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  Overview General/a_CA PPM Adapter Config List'),3)

applyhighlight(CA_PPM_Adapter_Config_List_rule)
WebUI.delay(1)
WebUI.click(findTestObject('Page_CA PPM  Overview General/a_CA PPM Adapter Config List'))
removehighlight(home_page_add)
WebUI.delay(1)
extentTest1.log(LogStatus.INFO, "Entered into GD Adapter Config List");
String screenShotPathFields_1 = capture("ConfigListValidation1");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathFields_1));

removehighlight(CA_PPM_Adapter_Config_List_rule)

WebElement a_Transaction_fields_1 = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config/a_Transaction'),3)

applyhighlight(a_Transaction_fields_1)
WebUI.delay(1)
WebUI.click(findTestObject('Page_CA PPM  CA PPM Adapter Config/a_Transaction'))
WebUI.delay(1)
extentTest1.log(LogStatus.INFO, "Clicked on the Transaction Button");
String screenShotPathFields_Txn_1 = capture("Transaction_button_1");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathFields_Txn_1));


WebUI.delay(3)*/

WebElement FileProperties_rule = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  Overview General/FileProperties'),3)

applyhighlight(FileProperties_rule)
WebUI.delay(1)
WebUI.click(findTestObject('Page_CA PPM  Overview General/FileProperties'))

WebUI.delay(2)

WebElement bussiness_Rules_add = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/a_Business Rules'),3)

applyhighlight(bussiness_Rules_add)

WebUI.delay(2)
WebUI.click(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/a_Business Rules'))
WebUI.delay(1)
extentTest1.log(LogStatus.INFO, "Entered into Bussiness Rule Page");
String screenShotPathbussiness_rule_page = capture("Bussiness_Rule_Page");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathbussiness_rule_page));

GlobalVariable.AddedBusinessRule = true;

extentTest1.log(LogStatus.INFO, "Adding New Bussiness Rule");


WebElement button_New_add = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/button_New'),3)
	
applyhighlight(button_New_add)
WebUI.delay(2)
WebUI.click(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/button_New'))

WebUI.delay(2)

WebElement input_code_add = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/input_code'),3)

applyhighlight(input_code_add)
WebUI.setText(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/input_code'), GlobalVariable.input_code)

WebElement input_name_add = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/input_name'),3)

applyhighlight(input_name_add)
WebUI.setText(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/input_name'), GlobalVariable.input_name)

WebElement input_orderid_add = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/input_orderid'),3)

applyhighlight(input_orderid_add)
WebUI.setText(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/input_orderid'), GlobalVariable.input_orderid)

if(GlobalVariable.input_isactive.equalsIgnoreCase("true"))
{
WebElement input_isactive_add = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/input_isactive'),3)

applyhighlight(input_isactive_add)
WebUI.click(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/input_isactive'))
}

WebElement textarea_sqlval_add = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/textarea_sqlval'),3)

applyhighlight(textarea_sqlval_add)
WebUI.setText(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/textarea_sqlval'), GlobalVariable.sqlvalue_new)

WebElement button_Save_And_Return_add = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/button_Save And Return'),3)

applyhighlight(button_Save_And_Return_add)
extentTest1.log(LogStatus.INFO, "Bussiness Rule Addition Page");
String screenShot_rule_add_page = capture("Bussiness_Rule_Add_Page");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShot_rule_add_page));
WebUI.delay(1)
WebUI.click(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/button_Save And Return'))

WebUI.delay(2)




extentTest1.log(LogStatus.INFO, "Adding New Bussiness Rule without required fields");

WebElement button_New_add_fail = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/button_New'),3)
	
applyhighlight(button_New_add_fail)
WebUI.click(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/button_New'))

WebUI.delay(2)

WebElement input_code_add_fail = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/input_code'),3)

applyhighlight(input_code_add_fail)
WebUI.setText(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/input_code'), GlobalVariable.input_code)

WebElement input_name_add_fail = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/input_name'),3)

applyhighlight(input_name_add_fail)
WebUI.setText(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/input_name'), GlobalVariable.input_name)

WebElement button_Save_And_Return_add_fail = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/button_Save And Return'),3)

applyhighlight(button_Save_And_Return_add_fail)
WebUI.click(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/button_Save And Return'))
WebUI.delay(2)

String name = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config (1)/errormessagebusinessrule'),0).getText()
println name
if(name.equalsIgnoreCase('ErrorAll required fields need to be filled out.'))
{	
	println "Test case is failed"
	extentTest1.log(LogStatus.FAIL, "Test Case failed as unable to create new business rule");
	WebUI.delay(2)
	WebElement buttonReturn_add_fail = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/button_Return'),3)
	applyhighlight(buttonReturn_add_fail)
	extentTest1.log(LogStatus.FAIL, "Bussiness Rule Addition Page without Required Fields");
	String screenShot_rule_add_fail_page = capture("Bussiness_Rule_Add_fail_Page");
	extentTest1.log(LogStatus.FAIL, "Snapshot below: " + extentTest1.addScreenCapture(screenShot_rule_add_fail_page));
	WebUI.delay(1)
	WebUI.click(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/button_Return'))
}
else
{
	println "Test case is Passed"
	extentTest1.log(LogStatus.PASS, "Test Case passed as we are able to create new business rule without requried fields");
}


WebUI.delay(2)



extentTest1.log(LogStatus.INFO, "Editing the existing Bussiness Rule");

WebDriver driver = DriverFactory.getWebDriver()
List<WebElement> var = driver.findElements(By.xpath('//a[text()="Order ID"]'));
for (a in var)
{
	 visiable = a.displayed;
	 visiable1 = a.enabled;
	 println visiable
	 println visiable1
	 if(visiable && visiable1)
	 {
		 applyhighlight(a)
		 WebUI.delay(2)
		 a.click();
	 
	 }
}
WebUI.delay(2)
driver.findElement(By.xpath('//tr[1]/td[2]/a[@id="odf.gd_business_rulesProperties"]')).click();
WebUI.delay(2)
WebElement textarea_sqlval_edit = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/textarea_sqlval'),3)

applyhighlight(textarea_sqlval_edit)
WebUI.setText(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/textarea_sqlval'), GlobalVariable.sqlval_modified)

WebElement button_Save_And_Return_edit = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/button_Save And Return'),3)

applyhighlight(button_Save_And_Return_edit)
WebUI.delay(2)
extentTest1.log(LogStatus.INFO, "Bussiness Rule Editing Page");
String screenShot_rule_edit_page = capture("Bussiness_Rule_edit_Page");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShot_rule_edit_page));

WebUI.click(findTestObject('Page_CA PPM  CA PPM Adapter Config (1)/button_Save And Return'))

WebUI.delay(2)


extent1.endTest(extentTest1);
extent1.flush();

public void applyhighlight(WebElement element)
{
	WebDriver driver = DriverFactory.getWebDriver()
	JavascriptExecutor js=(JavascriptExecutor)driver;
	js.executeScript("arguments[0].style.border='3px solid red'", element);
}

public void applyhighlightelements(WebElement element)
{
	WebDriver driver = DriverFactory.getWebDriver()
	JavascriptExecutor js=(JavascriptExecutor)driver;
	js.executeScript("arguments[0].style.border='3px solid green'", element);
}

public void removehighlight(WebElement element)
{
	WebDriver driver = DriverFactory.getWebDriver()
	JavascriptExecutor js=(JavascriptExecutor)driver;
	js.executeScript("arguments[0].style.border=''", element);
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