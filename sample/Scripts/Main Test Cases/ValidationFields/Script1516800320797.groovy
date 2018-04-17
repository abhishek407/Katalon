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
import org.stringtemplate.v4.compiler.STParser.element_return
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
import org.openqa.selenium.By.ByXPath
import org.openqa.selenium.JavascriptExecutor
import org.openqa.selenium.WebDriver as WebDriver
import org.openqa.selenium.WebElement as WebElement
import org.openqa.selenium.support.ui.Select

import com.kms.katalon.core.webui.driver.DriverFactory as DriverFactory
import com.relevantcodes.extentreports.ExtentReports
import com.relevantcodes.extentreports.ExtentTest
import com.relevantcodes.extentreports.LogStatus

ExtentReports extent1 = newkeyword.extent;
ExtentTest extentTest1 = newkeyword.test;

extentTest1 = extent1.startTest("Validation Fields TestCase");

KeywordLogger log = new KeywordLogger()

ExcelData myData = ((findTestData('FieldsValidation')) as ExcelData)
println(myData.getAllData())
row = myData.getRowNumbers()

List<String> list_first_column = new ArrayList<String>()
List<String> list_second_column = new ArrayList<String>()
Map<String, String> excel_field_names = new LinkedHashMap<String, String>()
for (int i = 1; i <= row; i++) 
{
	list_first_column.add(myData.getValue(1, i))
	list_second_column.add(myData.getValue(2, i))
	excel_field_names.put(myData.getValue(2, i), myData.getValue(3, i))
}
println(list_first_column)
println(list_second_column)
println(excel_field_names)

WebElement home_page = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/span_Home'),
	3)

applyhighlight(home_page)

WebUI.click(findTestObject('Page_CA PPM  Overview General/span_Home'))

WebUI.delay(2)

WebElement CA_PPM_Adapter_Config_List = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  Overview General/a_CA PPM Adapter Config List'),3)

applyhighlight(CA_PPM_Adapter_Config_List)
WebUI.delay(1)
WebUI.click(findTestObject('Page_CA PPM  Overview General/a_CA PPM Adapter Config List'))
removehighlight(home_page)
WebUI.delay(1)
extentTest1.log(LogStatus.INFO, "Entered into GD Adapter Config List");
String screenShotPathFields = capture("ConfigListValidation");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathFields));

removehighlight(CA_PPM_Adapter_Config_List)

WebElement a_Transaction_fields = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config/a_Transaction'),3)

applyhighlight(a_Transaction_fields)
WebUI.delay(1)
WebUI.click(findTestObject('Page_CA PPM  CA PPM Adapter Config/a_Transaction'))
WebUI.delay(1)
extentTest1.log(LogStatus.INFO, "Clicked on the Transaction Button");
String screenShotPathFields_Txn = capture("Transaction_button");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathFields_Txn));

WebElement FileProperties = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  Overview General/FileProperties'),3)

applyhighlight(FileProperties)
WebUI.delay(1)
WebUI.click(findTestObject('Page_CA PPM  Overview General/FileProperties'))

WebElement a_File_Config = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  CA PPM Adapter Config/a_File Config'),3)

applyhighlight(a_File_Config)
WebUI.delay(1)

WebUI.click(findTestObject('Page_CA PPM  CA PPM Adapter Config/a_File Config'))
WebUI.delay(1)
extentTest1.log(LogStatus.INFO, "Entered into File Configuration Page");
String screenShotPathFields_page = capture("File_Configuration_page");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathFields_page));

InputFilePath = WebUI.getAttribute(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config/InputFilePath'),'value')
InputFileName = WebUI.getAttribute(findTestObject('Page_CA PPM  CA PPM Adapter Config/InputFileName'), 'value')
ArchiveFilePath = WebUI.getAttribute(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config/ArchiveFilePath'), 'value')
ErrorFilePath = WebUI.getAttribute(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config/ErrorFilePath'), 'value')
FileDataFormat = WebUI.getAttribute(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config/FileDataFormat'), 'value')
FileAppendTimestamp = WebUI.getAttribute(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config/FileAppendTimestamp'), 'checked')
FileDelimiter = WebUI.getAttribute(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config/FileDelimiter'), 'value')
DataRowStartsAt = WebUI.getAttribute(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config/DataRowStartsAt'), 'value')
NoofColumnsInFile = WebUI.getAttribute(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config/NoofColumnsInFile'), 'value')
Select select = new Select(WebUiCommonHelper.findWebElement(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config/ProcessFile'),3))
WebElement option = select.getFirstSelectedOption()

String ProcessFile = option.getText()

Exception_Email_Subject = WebUI.getAttribute(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config (1)/Exception_Email_Subject'),'value')
Exception_Template = WebUI.getAttribute(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config (1)/Exception_Template'),'value')
Generate_Error_File = WebUI.getAttribute(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config (1)/Generate_Error_File'), 'checked')
Validate_File_Columns = WebUI.getAttribute(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config (1)/Validate_File_Columns'), 'checked')
Ignore_File_On_Validation = WebUI.getAttribute(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config (1)/Ignore_File_On_Validation'), 'checked')
No_of_Trailer_Records= WebUI.getAttribute(findTestObject('Object Repository/Page_CA PPM  CA PPM Adapter Config (1)/No_of_Trailer_Records'),'value')


Generate_Error_File = processfiledata(Generate_Error_File)
Validate_File_Columns = processfiledata(Validate_File_Columns)
Ignore_File_On_Validation = processfiledata(Ignore_File_On_Validation)
FileAppendTimestamp = processfiledata(FileAppendTimestamp)

println InputFilePath
println InputFileName
println ArchiveFilePath
println ErrorFilePath
println FileDataFormat
println FileAppendTimestamp
println FileDelimiter
println DataRowStartsAt
println NoofColumnsInFile
println ProcessFile
println Exception_Email_Subject
println Exception_Template
println Generate_Error_File
println Validate_File_Columns
println Ignore_File_On_Validation
println No_of_Trailer_Records


public String processfiledata(String data)
{
	if(data == null)
	{
		return "false";
	}
	else
	{
		return data;
	}
}

Map<String, String> UI_field_names = new LinkedHashMap<String, String>()
UI_field_names.put("InputFilePath",InputFilePath);
UI_field_names.put("InputFileName",InputFileName);
UI_field_names.put("ArchiveFilePath",ArchiveFilePath);
UI_field_names.put("ErrorFilePath",ErrorFilePath);
UI_field_names.put("FileDataFormat",FileDataFormat);
UI_field_names.put("FileAppendTimestamp",FileAppendTimestamp);
UI_field_names.put("FileDelimiter",FileDelimiter);
UI_field_names.put("DataRowStartsAt",DataRowStartsAt);
UI_field_names.put("NoofColumnsInFile",NoofColumnsInFile);
UI_field_names.put("ProcessFile",ProcessFile);
UI_field_names.put("Exception_Email_Subject",Exception_Email_Subject);
UI_field_names.put("Exception_Template",Exception_Template);
UI_field_names.put("Generate_Error_File",Generate_Error_File);
UI_field_names.put("Validate_File_Columns",Validate_File_Columns);
UI_field_names.put("Ignore_File_On_Validation",Ignore_File_On_Validation);
UI_field_names.put("No_of_Trailer_Records",No_of_Trailer_Records);


WebDriver driver = DriverFactory.getWebDriver()
String path1 = "//label[text()='"; 
String path2 =  "']";

for(int i=0;i<list_first_column.size();i++)
{
String finalpath1 = path1+list_first_column.get(i)+path2;
WebElement element;
try 
{
		element = driver.findElement(By.xpath(finalpath1));
}
catch(Exception e)
	{
	println "Element is not found : "+list_first_column.get(i);
	extentTest1.log(LogStatus.FAIL, "Element is not found in the page : '"+list_first_column.get(i)+"'");
	String screenShotPath_element_not_found_page = capture("element_not_found_page"+i);
	extentTest1.log(LogStatus.FAIL, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath_element_not_found_page));
	continue;
	}
applyhighlightelements(element);
boolean value = element.displayed;
if(value)
{
	println "Element is displayed : "+list_first_column.get(i);
	extentTest1.log(LogStatus.PASS, "Element is displayed in the page : '"+list_first_column.get(i)+"'");
	String screenShotPathelement_page = capture("element_displayed_page"+i);
	extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathelement_page));
	String name = list_second_column.get(i);
	if(excel_field_names.get(name).equalsIgnoreCase(UI_field_names.get(name)))	
		{
		println "Value in the field is matched : "+excel_field_names.get(name);
		extentTest1.log(LogStatus.PASS, "Value in the field is matched : '"+excel_field_names.get(name)+"'");
		String screenShotPathvalue_page = capture("element_value_page"+i);
		extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathvalue_page));
		}
	else
		{
		println "Value in the field is not matched : "+excel_field_names.get(name);
		extentTest1.log(LogStatus.FAIL, "Value in the field is not matched : '"+excel_field_names.get(name)+"'");
		String screenShotPath_not_value_page = capture("element_not_value_page"+i);
		extentTest1.log(LogStatus.FAIL, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath_not_value_page));
		}
}
else
{
	println "Element is not displayed : "+list_first_column.get(i);
	extentTest1.log(LogStatus.FAIL, "Element is not displayed in the page : '"+list_first_column.get(i)+"'");
	String screenShotPath_not_element_page = capture("element_not_displayed_page"+i);
	extentTest1.log(LogStatus.FAIL, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath_not_element_page));
}
}

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