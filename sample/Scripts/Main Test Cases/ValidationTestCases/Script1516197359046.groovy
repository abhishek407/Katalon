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

extentTest1 = extent1.startTest("Validation ErrorMessages TestCase");

KeywordLogger log = new KeywordLogger()

ExcelData myData = ((findTestData('Test_Case_Validation')) as ExcelData)

myData.changeSheet('Error_Messages')

println(myData.getAllData())

row = myData.getRowNumbers()

List<String> list = new ArrayList<String>()

for (int i = 1; i <= row; i++) {
    list.add(myData.getValue(1, i))
}

println(list)

myData.changeSheet('Error_Message_Validation')

println(myData.getAllData())

row1 = myData.getRowNumbers()

List<String> list1 = new ArrayList<String>()

List<String> list2 = new ArrayList<String>()

for (int i = 1; i <= row1; i++) {
    list1.add(myData.getValue(1, i))

    list2.add(myData.getValue(2, i))
}

println(list1)

println(list2)

ExcelData myData1 = ((findTestData('Input_File')) as ExcelData)

input_rows = myData1.getRowNumbers()

input_coloumns = myData1.getColumnNumbers()

WebElement home_page = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/span_Home'), 
    3)

applyhighlight(home_page)

WebUI.click(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/span_Home'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

WebElement GD_Adapter_Portlets_Summary = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/a_GD Adapter Portlets Summary'), 
    3)

applyhighlight(GD_Adapter_Portlets_Summary)

WebUI.scrollToElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/a_GD Adapter Portlets Summary'), 0, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/a_GD Adapter Portlets Summary'), FailureHandling.CONTINUE_ON_FAILURE)

removehighlight(home_page)

removehighlight(GD_Adapter_Portlets_Summary)

WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)
extentTest1.log(LogStatus.INFO, "Entered into GD Transaction Adapter Page");
String screenShotPath = capture("JobsIntialPage");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath));

WebElement GetAdapter_Status = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/GetAdapter Status'), 
    0)

applyhighlight(GetAdapter_Status)

status = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/GetAdapter Status'), 
    0).getText()

println(status)

WebUI.delay(2, FailureHandling.CONTINUE_ON_FAILURE)

String excelFilePath =  GlobalVariable.output_file_path;

String screenshot_path =  GlobalVariable.screenshot_path + "ValidationTestCases\\";

String totalrows = '0'

String sucessrows = '0'

String errorrows = '0'

WebDriver driver = DriverFactory.getWebDriver()

try {
    FileInputStream file = new FileInputStream(new File(excelFilePath))

    HSSFWorkbook workbook = WorkbookFactory.create(file)

    HSSFSheet sheet = workbook.getSheetAt(0)

    int rowCount = sheet.getLastRowNum()

    if (status.contentEquals('COMPLETED')) {
        println('Job is Completed Sucessfully')

        log.logPassed('Job is Completed Sucessfully')

        WebUI.takeScreenshot(screenshot_path + '\\JobCompleted.png')
		
		extentTest1.log(LogStatus.INFO, "Job is in Completed Status");
		String screenShotPath1 = capture("CompletedStatus");
		extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath1));
		

        totalrows = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/GetTotalRows'), 
            0).getText()

        sucessrows = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/GetProcessedRows'), 
            0).getText()

        errorrows = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/GetErrorRows'), 
            0).getText()
		if(totalrows.equalsIgnoreCase(errorrows))
		{
			GlobalVariable.RunJob = false;
		}

        WebUI.delay(2)

        WebElement ClickJob = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/ClickJob'), 
            3)

        applyhighlight(ClickJob)
		
		

        WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/ClickJob'), 
            3).click()

        WebUI.delay(2)
		
		extentTest1.log(LogStatus.INFO, "Entered into Validation page");
		String screenShotPath2 = capture("Validationpage");
		extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath2));

        WebElement Table = driver.findElement(By.xpath('//table[@class ="ppm_grid"]/tbody'))

        'To locate rows of table it will Capture all the rows available in the table '
        List<WebElement> Rows = Table.findElements(By.tagName('tr'))

        println('No. of rows: ' + Rows.size())

        'Find a matching text in a table and performing action'

        'Loop will execute for all the rows of the table'
        for (int i = 0; i < Rows.size(); i++) {
           
			 int a = i + 1
			


            'To locate columns(cells) of that specific row'
            List<WebElement> Cols = Rows.get(i).findElements(By.tagName('td'))
			
			applyhighlightelements(Cols.get(0))
			WebUI.delay(1)
			removehighlight(Cols.get(0))

            error_count = Cols.get(6).getText()

            error_message = Cols.get(7).getText()

            int firstindex = error_message.indexOf('(')

            int lastindex = error_message.indexOf(')')

            while ((firstindex != -1) && (lastindex != -1)) {
                String var = error_message.substring(firstindex, lastindex + 2)

                error_message = error_message.replace(var, '')

                System.out.println(error_message)

                firstindex = error_message.indexOf('(')

                lastindex = error_message.indexOf(')')
            }
            
            if (error_count.equalsIgnoreCase('0')) {
                if (list1.get(i).equalsIgnoreCase(error_count)) {
                    List<String> excel_data = new ArrayList<String>()

                    for (int local = 1; local <= input_coloumns; local++) {
                        excel_data.add(myData1.getValue(local, i + 1))
                    }
                    
                    loc_PostingDate = (excel_data[0])

                    loc_OracleProjectID = (excel_data[1])

                    loc_OracleTaskID = (excel_data[2])

                    loc_POID = (excel_data[3])

                    loc_ResourceID = (excel_data[4])

                    loc_TransClass = (excel_data[5])

                    loc_Units = (excel_data[6])

                    loc_Cost = (excel_data[7])

                    loc_CurrencyCode = (excel_data[8])
                    
                    UI_OracleProjectID = Cols.get(11).getText()

                    UI_OracleTaskID = Cols.get(12).getText()

                    UI_ResourceID = Cols.get(14).getText()

                    UI_Transaction_type = Cols.get(19).getText()

                    UI_Cost = Cols.get(21).getText()

                    UI_Input_type_code = Cols.get(22).getText()

                    DecimalFormat df = new DecimalFormat('#.####')

                    double cost = Double.parseDouble(loc_Cost)
					
                    double units = Double.parseDouble(loc_Units)

                    double price = cost / units
					
					if(GlobalVariable.AddedBusinessRule)
					{
						price = price + 1;
					}

                    String final_price = df.format(price)

                    UI_Cost = UI_Cost.replace(',', '')

                    double UI_Cost_converted = Double.parseDouble(UI_Cost)

                    String UI_Cost = df.format(UI_Cost_converted)

                    println(loc_OracleProjectID)

                    println(loc_OracleTaskID)

                    println(loc_ResourceID)

                    println(final_price)

                    println(UI_OracleProjectID)

                    println(UI_OracleTaskID)

                    println(UI_ResourceID)

                    println(UI_Cost)

                    if ((loc_OracleProjectID.equalsIgnoreCase(UI_OracleProjectID) && loc_ResourceID.equalsIgnoreCase(UI_ResourceID)) && 
                    final_price.equalsIgnoreCase(UI_Cost)) {
                        int columnCount = 0
                        Row row = sheet.createRow(++rowCount)

                        for (def var : excel_data) {
                            Cell cell = row.createCell(columnCount++)

                            cell.setCellValue(var)
                        }
                        
                        log.logPassed('Test case validtion Passed for error count equal to 0 & input data is also matching')

                        WebUI.takeScreenshot(((screenshot_path + '\\PassScreesnshot') + i) + '.png')
						
						extentTest1.log(LogStatus.PASS, "Test case validtion passeed for as the expected data and actual data is matching.Row No :"+a);
						String screenShotPath_1 = capture("ValidationTestCase_error"+a);
						extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath_1));
						
                    } 
					else {
                        log.logFailed('Test case validtion failed for error count 0 as the expected data and actual data is not matching')

                        WebUI.takeScreenshot(((screenshot_path + '\\ErrorScrreenhot') + i) + '.png')
						
						extentTest1.log(LogStatus.FAIL, "Test case validtion failed for as the expected data and actual data is not matching.Row No :"+a);
						String screenShotPath_2 = capture("ValidationTestCase_error"+a);
						extentTest1.log(LogStatus.FAIL, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath_2));
                    }
                } //println(status)
                else {
                    println('Test case validtion failed for error count 0')

                    log.logFailed('Test case validtion failed for error count equal to 0')

                    WebUI.takeScreenshot(((screenshot_path + '\\ErrorScrreenhot') + i) + '.png')
					
					extentTest1.log(LogStatus.FAIL, "Test case validtion failed for as the expected error count and actual error count is not matching.Row No :"+a);
					String screenShotPath_3 = capture("ValidationTestCase_error"+a);
					extentTest1.log(LogStatus.FAIL, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath_3));
                }
            } else if (error_count.equalsIgnoreCase('1')) {
                error_message = error_message.replaceAll('##', '').trim()

                println(error_message)

                if (list.contains(error_message)) {
                    if (list1.get(i).equalsIgnoreCase(error_count)) {
                        if (list2.get(i).equalsIgnoreCase(error_message)) {
                            println('Test case validtion passed for error count 1 & error message is also same')

                            log.logPassed('Test case validtion passed for error count 1 & error message is also same')
							
							extentTest1.log(LogStatus.PASS, "Test case validtion passeed for as the expected data and actual data is matching.Row No :"+a);
							String screenShotPath_4 = capture("ValidationTestCase_error"+a);
							extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath_4));

                            WebUI.takeScreenshot(((screenshot_path + '\\PassScrreenhot') + i) + '.png')
                        } else {
                            println('Test case validtion passed for error count 1 but failed matching error message')

                            log.logFailed('Test case validtion failed as the actual error message is not matching with expected error messsage.Row No :' + 
                                a)

                            WebUI.takeScreenshot(((screenshot_path + '\\ErrorScrreenhot') + i) + '.png')
							
							extentTest1.log(LogStatus.FAIL, "Test case validtion failed as the actual error message is not matching with expected error messsage.Row No :"+a);
							String screenShotPath_5 = capture("ValidationTestCase_error"+a);
							extentTest1.log(LogStatus.FAIL, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath_5));
                        }
                    } else {
                        println('Test case validtion failed for error count 1')

                        log.logFailed('Test case validtion failed as the expected error count and actual error count is not matching.Row No :' + 
                            a)

                        WebUI.takeScreenshot(((screenshot_path + '\\ErrorScrreenhot') + i) + '.png')
						
						extentTest1.log(LogStatus.FAIL, "Test case validtion failed as the expected error count and actual error count is not matching.Row No :"+a);
						String screenShotPath_6 = capture("ValidationTestCase_error"+a);
						extentTest1.log(LogStatus.FAIL, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath_6));
                    }
                } else {
                    println('Test case validtion failed as the error message not exist for error count 1')

                    log.logFailed('Test case validtion failed as the error message not exist in excel sheet.Row No :' + 
                        a)

                    WebUI.takeScreenshot(((screenshot_path + '\\ErrorScrreenhot') + i) + '.png')
					
					extentTest1.log(LogStatus.FAIL, "Test case validtion failed as the error message thrown in application not exist.Row No :"+a);
					String screenShotPath_7 = capture("ValidationTestCase_error"+a);
					extentTest1.log(LogStatus.FAIL, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath_7));
					
                }
            } else {
                String[] words = error_message.split('##')

                int value = words.length

                int counter = 0

                println(words)

                var = true

                for (String w : words) {
                    message = w.trim()

                    if (list.contains(message)) {
                        println(message)

                        if (list1.get(i).equalsIgnoreCase(error_count)) {
                            String[] words2 = list2.get(i).split('\\*')

                            for (String w1 : words2) {
                                if (w1.equalsIgnoreCase(message)) {
                                    counter++
                                }
                            }
                        } else {
                            println('Test case validation failed for error count 2 or more')

                            log.logFailed('Test case validtion failed for error count 2 or more as the expected error count and application error count is not matching.Row No :' + 
                                a)

                            WebUI.takeScreenshot(((screenshot_path + '\\ErrorScrreenhot') + i) + '.png')
							
							extentTest1.log(LogStatus.FAIL, "Test case validtion failed as the expected error count and actual error count is not matching.Row No :"+a);
							String screenShotPath_8 = capture("ValidationTestCase_error"+a);
							extentTest1.log(LogStatus.FAIL, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath_8));
                        }
                    } else {
                        println(message)

                        println('Test case validtion failed as the error message not exist in excel for error count 2 or more')

                        log.logFailed('Test case validtion failed as the error message not exist in excel for error count 2 or more,Looks like new error please check the application.Row No :' + 
                            a)

                        WebUI.takeScreenshot(((screenshot_path + '\\ErrorScrreenhot') + i) + '.png')
						
						extentTest1.log(LogStatus.FAIL, "Test case validtion failed as the error message thrown in application not exist.Row No :"+a);
						String screenShotPath_9 = capture("ValidationTestCase_error"+a);
						extentTest1.log(LogStatus.FAIL, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath_9));
                        var = false

                        break
                    }
                }
                
                if ((value == counter) && var) {
                    println('Test case validtion passed for error count 2 or more & error message is also same')

                    log.logPassed('Test case validtion passed for error count 2 or more & error messages is also same as the expected error messaage and UI message is matching')

                    WebUI.takeScreenshot(((screenshot_path + '\\PassScrreenhot') + i) + '.png')
					
					extentTest1.log(LogStatus.PASS, "Test case validtion passeed for as the expected data and actual data is matching.Row No :"+a);
					String screenShotPath_10 = capture("ValidationTestCase_error"+a);
					extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath_10));
                } else if ((value != counter) && var) {
                    println('Test case validtion passed for error count 2 or more but failed matching error message')

                    log.logFailed('Test case validtion passed for error count 2 or more but failed matching the error messages as expected error messaage and UI message is not matching.Row No :' + 
                        a)

                    WebUI.takeScreenshot(((screenshot_path + '\\ErrorScrreenhot') + i) + '.png')
					
					extentTest1.log(LogStatus.FAIL, "Test case validtion failed as the actual error message is not matching with expected error messsage.Row No :"+a);
					String screenShotPath_11 = capture("ValidationTestCase_error"+a);
					extentTest1.log(LogStatus.FAIL, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath_11));
                }
            }
        }
    } else {
        println('Test Failed because of reason:' + status)

        log.logFailed('Test Failed because of reason:' + status)
		
		extentTest1.log(LogStatus.FAIL, "Job Got Failed with reason :"+ status);
		String screenShotPath3 = capture("JobFailed");
		extentTest1.log(LogStatus.FAIL, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath3));

        WebUI.takeScreenshot(screenshot_path + '\\JobNotCompleted.png')
    }
    
    FileOutputStream outputStream = new FileOutputStream(excelFilePath)

    workbook.write(outputStream)

    outputStream.close()
}
catch (IOException e) {
    e.printStackTrace()
}

extentTest1.log(LogStatus.INFO, "All Validations Completed");
String screenShotPath4 = capture("ValidationsCompleted");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPath4));
//WebDriver driver = DriverFactory.getWebDriver()

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