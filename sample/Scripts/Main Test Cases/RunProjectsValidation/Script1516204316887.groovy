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
import org.openqa.selenium.support.ui.Select

import com.kms.katalon.core.webui.driver.DriverFactory as DriverFactory
import com.relevantcodes.extentreports.ExtentReports
import com.relevantcodes.extentreports.ExtentTest
import com.relevantcodes.extentreports.LogStatus


ExtentReports extent1 = newkeyword.extent;
ExtentTest extentTest1 = newkeyword.test;

if(GlobalVariable.RunJob)
{
extentTest1 = extent1.startTest("Run ProjectsValidation TestCase");
String excelFilePath = GlobalVariable.output_file_path;

WebElement home_page_2 = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/span_Home'), 
    3)

applyhighlight(home_page_2)

WebUI.click(findTestObject('Page_CA PPM  Overview General/span_Home'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Page_CA PPM  Overview General/span_ppm_nav_app_menu'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

WebElement a_Projects_2 = WebUiCommonHelper.findWebElement(findTestObject('CAPPM12/Page_CA PPM  Overview General/a_Projects'), 
    3)

applyhighlight(a_Projects_2)

WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Overview General/a_Projects'), FailureHandling.CONTINUE_ON_FAILURE)

removehighlight(home_page_2)

removehighlight(a_Projects_2)


WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)
extentTest1.log(LogStatus.INFO, "Entered into Projects Page");
String screenShotPathPj = capture("ProjectsPages_home");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathPj));

WebElement img_page_projmgr = WebUiCommonHelper.findWebElement(findTestObject('CAPPM12/Page_CA PPM  Projects/img_page_projmgr.projectList_o'), 
    3)

applyhighlight(img_page_projmgr)

WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Projects/img_page_projmgr.projectList_o'), FailureHandling.CONTINUE_ON_FAILURE)

removehighlight(img_page_projmgr)

WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

//WebUI.takeScreenshot(screenshot_path + '\\ProjectsPage.png')
FileInputStream file1 = new FileInputStream(new File(excelFilePath))

try {
    HSSFWorkbook workbook1 = WorkbookFactory.create(file1)

    HSSFSheet sheet1 = workbook1.getSheetAt(0)

    int rowCount1 = sheet1.getLastRowNum()

    WebDriver driver = DriverFactory.getWebDriver()

    for (int i = 1; i <= rowCount1; i++) {
        String var = sheet1.getRow(i).getCell(1)

        WebElement input_unique_code = WebUiCommonHelper.findWebElement(findTestObject('CAPPM12/Page_CA PPM  Projects/input_unique_code'), 
            3)

        applyhighlight(input_unique_code)

        WebUI.setText(findTestObject('CAPPM12/Page_CA PPM  Projects/input_unique_code'), var)

        WebElement button_Filter_high = WebUiCommonHelper.findWebElement(findTestObject('CAPPM12/Page_CA PPM  Projects/button_Filter'), 
            3)

        applyhighlight(button_Filter_high)

        WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Projects/button_Filter'))

        WebElement Project_Name_path = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM12/Page_CA PPM  Projects/Project_Name_path'), 
            3)

        applyhighlight(Project_Name_path)
		
		extentTest1.log(LogStatus.INFO, "Searching with project Id : "+ var);
		String screenShotPathProjectid = capture("EachProjectPage"+var);
		extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathProjectid));

        String Project_Name = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM12/Page_CA PPM  Projects/Project_Name_path'), 
            3).getText()

        println(Project_Name)

        //WebUI.takeScreenshot(screenshot_path + '\\ProjectsPage'+var+'.png')
        Cell cell1 = sheet1.getRow(i).createCell(9)

        cell1.setCellValue(Project_Name)

        WebUI.click(findTestObject('Object Repository/CAPPM12/Page_CA PPM  Projects/Resource_Name_click'))

        WebUI.delay(2)

        WebElement ClickProperties = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickProperties'), 
            3)

        applyhighlight(ClickProperties)

        WebUI.click(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickProperties'))

        WebUI.delay(2)

        WebElement ClickFinanacial = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickFinanacial'), 
            3)

        applyhighlight(ClickFinanacial)

        WebUI.click(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickFinanacial'))
		
		extentTest1.log(LogStatus.INFO, "Clicked on finanacial Tab");
		String screenShotPathFinanacial = capture("FinanacialTab"+var);
		extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathProjectid));

        WebUI.delay(2)

        WebElement Project_cost_type_apply = driver.findElement(By.xpath('//*[@name=\'cost_type\']'))

        applyhighlight(Project_cost_type_apply)

        Select select = new Select(driver.findElement(By.xpath('//*[@name=\'cost_type\']')))

        WebElement option = select.getFirstSelectedOption()

        String Project_cost_type = option.getText()

        System.out.println(Project_cost_type)

        cell1 = sheet1.getRow(i).createCell(11)

        cell1.setCellValue(Project_cost_type)

        WebElement ClickTask = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickTask'), 
            3)

        applyhighlight(ClickTask)

        WebUI.click(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickTask'))
		WebUI.delay(1)
        //WebUI.click(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickTask'))

       // WebUI.scrollToElement(findTestObject('Object Repository/Page_CA PPM  Jobs Log/scrollcosttype'), 3)
		extentTest1.log(LogStatus.INFO, "Clicked on Task Tab");
		String screenShotPathtask = capture("TaskTab"+var);
		extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathtask));
		WebUI.delay(3)

        WebElement Table = driver.findElement(By.xpath('//table[@class ="ppm_grid"]/tbody'))

        List<WebElement> Rows = Table.findElements(By.tagName('tr'))

        println('No. of rows: ' + Rows.size())

        int counter_data = 0

        int j = 0

        for (j; j < Rows.size(); j++) {
            List<WebElement> Cols = Rows.get(j).findElements(By.tagName('td'))

            String UI_taskid = Cols.get(1).getText()

            String taskid = sheet1.getRow(i).getCell(2)

            println(UI_taskid)

            println(taskid)

            if (UI_taskid.equalsIgnoreCase(taskid)) {
                String cost_type = Cols.get(3).getText()

                println(cost_type)
				applyhighlight(Cols.get(1));
				extentTest1.log(LogStatus.INFO, "Seraching the cost type for taksid :"+taskid);
				String screenShotPathtaskid = capture("taskidvalid"+var);
				extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathtaskid));
				applyhighlight(Cols.get(3));
				WebUI.delay(2);
                cell1 = sheet1.getRow(i).createCell(12)

                cell1.setCellValue(cost_type)

                break;
            }
        }
        
        WebElement ClickProperties_apply_high = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickProperties'), 
            3)

        applyhighlight(ClickProperties_apply_high)

        WebUI.click(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickProperties'))

        WebUI.delay(1)

        WebUI.click(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickProperties'))

        WebUI.delay(2)
		WebUI.scrollToPosition(0, 5000)
		WebUI.delay(3)
        WebElement button_Return_apply = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  Project acdv - Propert/button_Return'), 
            3)

        applyhighlight(button_Return_apply)
     
        WebUI.click(findTestObject('Page_CA PPM  Project acdv - Propert/button_Return'))

        WebUI.delay(3) // WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM12/Page_CA PPM  Projects/ReturnB'),0).click()
    }
    
    FileOutputStream outputStream1 = new FileOutputStream(excelFilePath)

    workbook1.write(outputStream1)

    outputStream1.close()
}
catch (IOException e) {
    e.printStackTrace()
} 

WebElement home_page = WebUiCommonHelper.findWebElement(findTestObject('CAPPM/Page_CA PPM  Job Type CA PPM Adapte/span_Home'), 
    3)

applyhighlight(home_page)

WebUI.click(findTestObject('Page_CA PPM  Overview General/span_Home'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Page_CA PPM  Overview General/span_ppm_nav_app_menu'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

WebElement Resource_home_click = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/Page_CA PPM  Overview General (1)/Resource_home_click'), 
    3)

applyhighlight(Resource_home_click)

WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Object Repository/Page_CA PPM  Overview General (1)/Resource_home_click'), FailureHandling.CONTINUE_ON_FAILURE)

removehighlight(home_page)

removehighlight(Resource_home_click)

WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

extentTest1.log(LogStatus.INFO, "Entered into Resources Page");
String screenShotPathResource = capture("ResourcesPage_home");
extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathResource));

WebElement projmgr_getResources_apply = WebUiCommonHelper.findWebElement(findTestObject('Page_CA PPM  Resource List/img_page_projmgr.getResources_'), 
    3)

applyhighlight(projmgr_getResources_apply)

WebUI.click(findTestObject('Page_CA PPM  Resource List/img_page_projmgr.getResources_'), FailureHandling.CONTINUE_ON_FAILURE)

FileInputStream file = new FileInputStream(new File(excelFilePath))

try {
    HSSFWorkbook workbook = WorkbookFactory.create(file)

    HSSFSheet sheet = workbook.getSheetAt(0)

    int rowCount = sheet.getLastRowNum()

    for (int i = 1; i <= rowCount; i++) {
        String var = sheet.getRow(i).getCell(4)

        WebElement Resource_name = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/Resource_name'), 
            3)

        applyhighlight(Resource_name)

        WebUI.setText(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/Resource_name'), var)

        WebElement button_Filter_again = WebUiCommonHelper.findWebElement(findTestObject('CAPPM12/Page_CA PPM  Projects/button_Filter'), 
            3)

        applyhighlight(button_Filter_again)

        WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Projects/button_Filter'))

        WebElement Resource_Name_click = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM12/Page_CA PPM  Projects/Resource_Name_click'), 
            3)

        applyhighlight(Resource_Name_click)
		
		extentTest1.log(LogStatus.INFO, "Searching for Resource id :"+ var);
		String screenShotPathResourceid = capture("Resource_id"+var);
		extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathResourceid));

        WebUI.click(findTestObject('Object Repository/CAPPM12/Page_CA PPM  Projects/Resource_Name_click'))

        WebUI.delay(3)

        WebElement Input_Type_Code_path_apply = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM12/Page_CA PPM  Projects/Input_Type_Code_path'), 
            3)

        applyhighlight(Input_Type_Code_path_apply)

        String type_name = WebUI.getAttribute(findTestObject('Object Repository/CAPPM12/Page_CA PPM  Projects/Input_Type_Code_path'), 
            'value')

        println(type_name)
		
		extentTest1.log(LogStatus.INFO, "Validating the Type code for :"+ var);
		String screenShotPathTypeCode = capture("Type_code"+var);
		extentTest1.log(LogStatus.PASS, "Snapshot below: " + extentTest1.addScreenCapture(screenShotPathTypeCode));

        Cell cell = sheet.getRow(i).createCell(10)

        cell.setCellValue(type_name)

        WebUI.delay(2)

        WebElement Return_button_resource = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM12/Page_CA PPM  Projects/Return_button_resource'), 
            3)

        applyhighlight(Return_button_resource)

        WebUI.click(findTestObject('Object Repository/CAPPM12/Page_CA PPM  Projects/Return_button_resource'))
    }
    
    FileOutputStream outputStream = new FileOutputStream(excelFilePath)

    workbook.write(outputStream)

    outputStream.close()
}
catch (IOException e) {
    e.printStackTrace()
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