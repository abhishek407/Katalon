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
import oracle.net.ano.a

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

if(GlobalVariable.RunJob)
{
KeywordLogger log = new KeywordLogger()
String screenshot_path = GlobalVariable.screenshot_path+ "ValidationTransactions\\";

String start_date_ui = GlobalVariable.start_date;
String end_date_ui = GlobalVariable.end_date;

WebUI.click(findTestObject('Page_CA PPM  Overview General/span_Home'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Page_CA PPM  Overview General/span_ppm_nav_app_menu'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.scrollToElement(findTestObject('Object Repository/CAPPM/Page_CA PPM  Job Type CA PPM Adapte/Transactions Page'), 0, 
    FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Object Repository/CAPPM/Page_CA PPM  Job Type CA PPM Adapte/Transactions Page'), FailureHandling.CONTINUE_ON_FAILURE)
//WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)
WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Transaction List/img_Remove'), FailureHandling.CONTINUE_ON_FAILURE)
WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)
WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Transaction List/img_Browse'), FailureHandling.CONTINUE_ON_FAILURE)


WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Transaction List/img_x-form-trigger x-form-date'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.setText(findTestObject('CAPPM12/Page_CA PPM  Transaction List/span_1'), start_date_ui, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Transaction List/img_x-form-trigger x-form-date_1'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.setText(findTestObject('CAPPM12/Page_CA PPM  Transaction List/span_31'), end_date_ui, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)
WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Transaction List/button_Filter'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Transaction List/input_objectInstanceId'), FailureHandling.CONTINUE_ON_FAILURE)
//WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Transaction List/button_Add'), FailureHandling.CONTINUE_ON_FAILURE)

//WebUI.takeScreenshot(screenshot_path + '\\TransactionsPage.png', FailureHandling.CONTINUE_ON_FAILURE)


ExcelData myData4 = ((findTestData('Output_File')) as ExcelData)

rows = myData4.getRowNumbers()

coloumns = myData4.getColumnNumbers()

println(myData4.getAllData())

for (int i = 1; i <= rows; i++) {
    String project_name = myData4.getValue(10, i)

    String resource_id = myData4.getValue(5, i)

    String excel_date = myData4.getValue(1, i)
	String project_id = myData4.getValue(2, i)
    String excel_cost = myData4.getValue(8, i)
    String excel_units = myData4.getValue(7, i)
    String input_type_code = myData4.getValue(11, i)
	String project_cost_type = myData4.getValue(12, i)
	String task_cost_type = myData4.getValue(13, i)
	String final_cost_type
    println(input_type_code)
	println task_cost_type
	println project_cost_type
    if(input_type_code == "")
	{
    input_type_code= "BIllable";
	}
	println(input_type_code)
	if(task_cost_type == "")
	{
		final_cost_type = project_cost_type;
	}
	else
	{
		final_cost_type = task_cost_type;
	}
    double cost_updated = Double.parseDouble(excel_cost)

    double units_updated = Double.parseDouble(excel_units)

    double price_updated = cost_updated / units_updated

    double a_price = round(price_updated, 2)

    a_price = round(a_price * units_updated, 2)

    //String a_price_updated = a_price.toString();
    DecimalFormat format = new DecimalFormat('0.00')

    String a_price_updated = format.format(a_price)

    Date date1 = new SimpleDateFormat('dd/MM/yy').parse(excel_date)

    SimpleDateFormat formatter = new SimpleDateFormat('MM/dd/yy')

    String format_date = formatter.format(date1)

    println(project_name)

    println(resource_id)

    println(a_price_updated)

    println(format_date)

    println(input_type_code)
	println final_cost_type

   // WebUI.delay(2)
	WebUI.click(findTestObject('Object Repository/Page_CA PPM  Transaction List (1)/Removeinvestment'))

	WebUI.click(findTestObject('Object Repository/Page_CA PPM  Transaction List (1)/Addinvestment'))
	

	WebUI.setText(findTestObject('Object Repository/Page_CA PPM  Transaction List (1)/InvestmentIdName'),project_id)
	

	WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Transaction List/button_Filter'), FailureHandling.CONTINUE_ON_FAILURE)
	
	//WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)
	

	WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Transaction List/input_objectInstanceId'), FailureHandling.CONTINUE_ON_FAILURE)
	//WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

	WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Transaction List/button_Add'), FailureHandling.CONTINUE_ON_FAILURE)

	WebUI.delay(1)
	//resource addition
	WebUI.click(findTestObject('Object Repository/Page_CA PPM  Transaction List (1)/Removeresource'))

	WebUI.click(findTestObject('Object Repository/Page_CA PPM  Transaction List (1)/AddResource'))
	
	WebUI.setText(findTestObject('Object Repository/Page_CA PPM  Transaction List (1)/uniqueresourcename'),resource_id)
	
	WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Transaction List/button_Filter'), FailureHandling.CONTINUE_ON_FAILURE)
	
	WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Transaction List/input_objectInstanceId'), FailureHandling.CONTINUE_ON_FAILURE)
	
	WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)
	
	WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Transaction List/button_Add'), FailureHandling.CONTINUE_ON_FAILURE)
	
	
	WebUI.delay(2)
	WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Transaction List/button_Filter_1'))
	WebDriver driver = DriverFactory.getWebDriver()
	WebUI.delay(2)
	List<WebElement> var = driver.findElements(By.xpath('//table/thead/tr/th[3]/a'));
	for (a in var) 
	{
		 visiable = a.displayed;
		 visiable1 = a.enabled;
		 println visiable
		 println visiable1
		 if(visiable && visiable1)
		 {
			 WebUI.delay(2)
			 a.click();
		 
		 }
	}
	///driver.findElements(By.xpath('//*[@id="ppm_page_contents"]/div[2]/table/thead/tr/th[3]/a')).get(var-1).click();
	
    WebUI.delay(2)
	//WebDriver driver = DriverFactory.getWebDriver()
    WebElement Table = driver.findElement(By.xpath('//table[@class ="ppm_grid"]/tbody'))

    List<String> Rows = Table.findElements(By.tagName('tr'))
	

    println('No. of rows: ' + Rows.size())

    int counter_data = 0

    int j = 0

    for (j; j < Rows.size(); j++) {
        List<String> Cols = Rows.get(j).findElements(By.tagName('td'))

        investment = Cols.get(1).getText()

        date = Cols.get(2).getText()

        Date date2 = new SimpleDateFormat('MM/dd/yy').parse(date)

        SimpleDateFormat formatter2 = new SimpleDateFormat('MM/dd/yy')

        String format_date2 = formatter2.format(date2)

        resource = Cols.get(3).getText()

        cost = Cols.get(10).getText().replaceAll('USD','')

        cost = cost.replace(',','').trim()

        input_type = Cols.get(9).getText()
		cost_type = Cols.get(7).getText()

        println(investment)

        println (format_date2)

        println(resource)

        println(cost)

        println(input_type)
		println(cost_type)
        if (investment.equalsIgnoreCase(project_name) && resource.equalsIgnoreCase(resource_id) && cost.equalsIgnoreCase(
            a_price_updated) && format_date2.equalsIgnoreCase(format_date) && input_type.equalsIgnoreCase(input_type_code)
			&& cost_type.equalsIgnoreCase(final_cost_type))
		{
            counter_data++;
			break;
        }
        
        if (j == 19) {
            j = 0;

            WebUI.click(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/Next_page_option'))

            WebUI.delay(3)

            Table = driver.findElement(By.xpath('//table[@class ="ppm_grid"]/tbody'))

            Rows = Table.findElements(By.tagName('tr'))
            WebUI.delay(3)
        }
    }
    
    if (counter_data == 1) {
        println('Passed')

        log.logPassed('Test case validtion Passed')

        WebUI.takeScreenshot(((screenshot_path + '\\TransactionPassScreesnshot') + i) + '.png')
    } else {
        println('failed')

        log.logFailed('Test case validtion failed has the UI content is not matching with expected result')

        WebUI.takeScreenshot(((screenshot_path + '\\TransactionErrorScreenshot') + i) + '.png')
    }
}

}
public static double round(double value, int places)
{

	long factor = (long) Math.pow(10, places);
	value = value * factor;
	long tmp = Math.round(value);
	return (double) tmp / factor;
}

WebUI.callTestCase(findTestCase('Test Cases/Main Test Cases/CloserSteps'), [:])