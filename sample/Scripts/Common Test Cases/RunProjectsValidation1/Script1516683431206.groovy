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
import org.openqa.selenium.support.ui.Select as Select
import com.kms.katalon.core.webui.driver.DriverFactory as DriverFactory

if(GlobalVariable.RunJob)
{
String excelFilePath = GlobalVariable.output_file_path;


WebUI.click(findTestObject('Page_CA PPM  Overview General/span_Home'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Page_CA PPM  Overview General/span_ppm_nav_app_menu'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.delay(2, FailureHandling.CONTINUE_ON_FAILURE)


WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Overview General/a_Projects'), FailureHandling.CONTINUE_ON_FAILURE)


WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Projects/img_page_projmgr.projectList_o'), FailureHandling.CONTINUE_ON_FAILURE)


//WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)

//WebUI.takeScreenshot(screenshot_path + '\\ProjectsPage.png')
FileInputStream file1 = new FileInputStream(new File(excelFilePath))

try {
    HSSFWorkbook workbook1 = WorkbookFactory.create(file1)

    HSSFSheet sheet1 = workbook1.getSheetAt(0)

    int rowCount1 = sheet1.getLastRowNum()

    WebDriver driver = DriverFactory.getWebDriver()

    for (int i = 1; i <= rowCount1; i++) {
        String var = sheet1.getRow(i).getCell(1)

        WebUI.setText(findTestObject('CAPPM12/Page_CA PPM  Projects/input_unique_code'), var)


        WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Projects/button_Filter'))


        String Project_Name = WebUiCommonHelper.findWebElement(findTestObject('Object Repository/CAPPM12/Page_CA PPM  Projects/Project_Name_path'), 
            3).getText()

        println(Project_Name)

        //WebUI.takeScreenshot(screenshot_path + '\\ProjectsPage'+var+'.png')
        Cell cell1 = sheet1.getRow(i).createCell(9)

        cell1.setCellValue(Project_Name)

        WebUI.click(findTestObject('Object Repository/CAPPM12/Page_CA PPM  Projects/Resource_Name_click'))

       // WebUI.delay(2)

        WebUI.click(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickProperties'))

        //WebUI.delay(2)

        WebUI.click(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickFinanacial'))

        WebUI.delay(2)

        Select select = new Select(driver.findElement(By.xpath('//*[@name=\'cost_type\']')))

        WebElement option = select.getFirstSelectedOption()

        String Project_cost_type = option.getText()

        System.out.println(Project_cost_type)

        cell1 = sheet1.getRow(i).createCell(11)

        cell1.setCellValue(Project_cost_type)


        WebUI.click(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickTask'))
		//WebUI.delay(1)
        WebUI.click(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickTask'))

       // WebUI.scrollToElement(findTestObject('Object Repository/Page_CA PPM  Jobs Log/scrollcosttype'), 3)

		WebUI.delay(3)

        WebElement Table = driver.findElement(By.xpath('//table[@class ="ppm_grid"]/tbody'))

        List<String> Rows = Table.findElements(By.tagName('tr'))

        println('No. of rows: ' + Rows.size())

        int counter_data = 0

        int j = 0

        for (j; j < Rows.size(); j++) {
            List<String> Cols = Rows.get(j).findElements(By.tagName('td'))

            String UI_taskid = Cols.get(1).getText()

            String taskid = sheet1.getRow(i).getCell(2)

            println(UI_taskid)

            println(taskid)

            if (UI_taskid.equalsIgnoreCase(taskid)) {
                String cost_type = Cols.get(3).getText()

                println(cost_type)
                cell1 = sheet1.getRow(i).createCell(12)

                cell1.setCellValue(cost_type)

                break;
            }
        }

        WebUI.click(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickProperties'))

        ///WebUI.delay(2)

        WebUI.click(findTestObject('Object Repository/Page_CA PPM  Jobs Log/ClickProperties'))

        WebUI.delay(2)
		WebUI.scrollToPosition(0, 5000)
		//WebUI.delay(3)
     
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

WebUI.click(findTestObject('Page_CA PPM  Overview General/span_Home'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Page_CA PPM  Overview General/span_ppm_nav_app_menu'), FailureHandling.CONTINUE_ON_FAILURE)

WebUI.delay(1, FailureHandling.CONTINUE_ON_FAILURE)


WebUI.delay(2, FailureHandling.CONTINUE_ON_FAILURE)

WebUI.click(findTestObject('Object Repository/Page_CA PPM  Overview General (1)/Resource_home_click'), FailureHandling.CONTINUE_ON_FAILURE)


WebUI.delay(3, FailureHandling.CONTINUE_ON_FAILURE)


WebUI.click(findTestObject('Page_CA PPM  Resource List/img_page_projmgr.getResources_'), FailureHandling.CONTINUE_ON_FAILURE)

FileInputStream file = new FileInputStream(new File(excelFilePath))

try {
    HSSFWorkbook workbook = WorkbookFactory.create(file)

    HSSFSheet sheet = workbook.getSheetAt(0)

    int rowCount = sheet.getLastRowNum()

    for (int i = 1; i <= rowCount; i++) {
        String var = sheet.getRow(i).getCell(4)

        WebUI.setText(findTestObject('Object Repository/CAPPM/Page_CA PPM  Overview General/Resource_name'), var)

        WebUI.click(findTestObject('CAPPM12/Page_CA PPM  Projects/button_Filter'))

        WebUI.click(findTestObject('Object Repository/CAPPM12/Page_CA PPM  Projects/Resource_Name_click'))

        //WebUI.delay(3)

        String type_name = WebUI.getAttribute(findTestObject('Object Repository/CAPPM12/Page_CA PPM  Projects/Input_Type_Code_path'), 
            'value')

        println(type_name)

        Cell cell = sheet.getRow(i).createCell(10)

        cell.setCellValue(type_name)

       WebUI.delay(2)

        WebUI.click(findTestObject('Object Repository/CAPPM12/Page_CA PPM  Projects/Return_button_resource'))
    }
    
    FileOutputStream outputStream = new FileOutputStream(excelFilePath)

    workbook.write(outputStream)

    outputStream.close()
}
catch (IOException e) {
    e.printStackTrace()
} 

}