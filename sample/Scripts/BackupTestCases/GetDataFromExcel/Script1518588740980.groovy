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
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys
import org.stringtemplate.v4.compiler.STParser.namedArg_return as namedArg_return
import com.kms.katalon.core.testdata.ExcelData as ExcelData
import com.kms.katalon.core.testdata.InternalData as InternalData
import com.kms.katalon.core.webui.common.WebUiCommonHelper as WebUiCommonHelper
import javax.swing.JOptionPane as JOptionPane

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
import internal.GlobalVariable




ExcelData myData = ((findTestData('Test_Case_Validation')) as ExcelData)

myData.changeSheet('Clarity_Test_Data_Setup')

println(myData.getAllData())

excel_column = myData.getColumnNumbers()
excel_row = myData.getRowNumbers()

Map<String, String> excel_names = new LinkedHashMap<String, String>()

for (int i = 1; i <= excel_row; i++)
{
	excel_names.put(myData.getValue(1, i), myData.getValue(2, i))
}

println(excel_names)

GlobalVariable.open_url = excel_names.get('url')

GlobalVariable.input_username = excel_names.get('username')

GlobalVariable.input_password = excel_names.get('password')

GlobalVariable.Job_name = excel_names.get('jobname')

GlobalVariable.screenshot_path = excel_names.get('screenshotpath')

GlobalVariable.click_path = excel_names.get('clickpath')

GlobalVariable.all_job_names = excel_names.get('alljobnames')

GlobalVariable.output_file_path = excel_names.get('outfilepath')

GlobalVariable.start_date = excel_names.get('startDate')

GlobalVariable.end_date = excel_names.get('endDate')

GlobalVariable.input_code = excel_names.get('inputcode')

GlobalVariable.input_name = excel_names.get('inputname')

GlobalVariable.input_orderid = excel_names.get('inputorderid')

GlobalVariable.input_isactive = excel_names.get('inputisactive')

GlobalVariable.sqlvalue_new = excel_names.get('textareasqlval_new')

GlobalVariable.sqlval_modified = excel_names.get('textareasqlval_modified')


Path src2 = Paths.get('D:\\Katalon_CAPPM\\Standard_file\\Output_file.xls')

Path dest2 = Paths.get(GlobalVariable.output_file_path)

Files.copy(src2, dest2, StandardCopyOption.REPLACE_EXISTING)


