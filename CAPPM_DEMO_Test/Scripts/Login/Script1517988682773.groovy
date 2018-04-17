import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.checkpoint.CheckpointFactory as CheckpointFactory
import com.kms.katalon.core.logging.KeywordLogger
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
import com.kms.katalon.core.webui.driver.DriverFactory
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUiBuiltInKeywords
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import internal.GlobalVariable as GlobalVariable

import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.openqa.selenium.By
import org.openqa.selenium.Keys as Keys
import org.openqa.selenium.WebDriver
import org.openqa.selenium.WebElement

KeywordLogger log = new KeywordLogger()
WebUI.openBrowser('')

WebUI.maximizeWindow()

WebUI.navigateToUrl('https://cppm9242-dev.ondemand.ca.com/niku/nu')

log.logFailed("Failed")

WebUI.setText(findTestObject('Page_CA PPM State Street Developmen/input_userName'), 'Ramyasree.SriNidadavolu@contractor.ca.com')

WebUI.setText(findTestObject('Page_CA PPM State Street Developmen/input_passWord'), 'Excers_2010')

WebUI.sendKeys(findTestObject('Page_CA PPM State Street Developmen/input_passWord'), Keys.chord(Keys.ENTER))

WebUI.delay(3)

WebUI.click(findTestObject('Page_CA PPM State Street Developmen/span_Home'))


WebUI.delay(2)

WebUI.click(findTestObject('Page_CA PPM State Street Developmen/a_Projects'))

WebUI.click(findTestObject('Page_CA PPM State Street Developmen/img_page_projmgr.projectList_o'))

WebUI.setText(findTestObject('Page_CA PPM State Street Developmen/input_unique_code'), 'PRJ00026')

WebUI.click(findTestObject('Page_CA PPM State Street Developmen/button_Filter'))


WebUI.delay(2)

WebUI.click(findTestObject('Page_CA PPM State Street Developmen/a_Beacon - Demo'))


WebUI.delay(1)
WebUI.click(findTestObject('Page_CA PPM State Street Developmen/a_Financial Plans'))

WebUI.delay(2)

WebUI.click(findTestObject('Page_CA PPM State Street Developmen/a_Cost Plan 2'))

WebUI.delay(3)

WebUI.click(findTestObject('Page_CA PPM State Street Developmen (1)/img_caui-workspaceWorkspaceHea'))

WebUI.delay(2)

WebUI.click(findTestObject('Page_CA PPM State Street Developmen (1)/a_Time-scaled Value'))

WebUI.delay(2)

WebUI.click(findTestObject('Page_CA PPM State Street Developmen (1)/input_startDateType'))

WebUI.setText(findTestObject('Page_CA PPM State Street Developmen (1)/input_startDate'), '1/1/2016')

WebUI.click(findTestObject('Page_CA PPM State Street Developmen (1)/input_timeScaleType'))

WebUI.selectOptionByValue(findTestObject('Page_CA PPM State Street Developmen (1)/select_timeScale'), 'quarter', true)

WebUI.setText(findTestObject('Page_CA PPM State Street Developmen (1)/input_numberOfTimePeriods'), '20')

WebUI.delay(3)

WebUI.click(findTestObject('Page_CA PPM State Street Developmen (1)/button_Save And Return'))

try {
String filename = "D:\\letzNav\\output_file.xls" ;
HSSFWorkbook workbook = new HSSFWorkbook();
HSSFSheet sheet = workbook.createSheet("FirstSheet")
int rowCount = 0
	println rowCount
WebUI.delay(5)
WebDriver driver = DriverFactory.getWebDriver()

WebElement row_count = driver.findElement(By.xpath('//div[@class="ppm_gridcontent"]//thead//tr[2]'))

List<WebElement> Rows = row_count.findElements(By.tagName('th'));
List<String> years_data = new ArrayList<String>();

for(int i=0;i<Rows.size();i++)
{
	years_data.add(Rows.get(i).getText());
}

println years_data
HSSFRow row = sheet.createRow(rowCount++)
int columnCount = 0
row.createCell(columnCount++).setCellValue("Field_Names")
for (def var : years_data) {
    HSSFCell cell = row.createCell(columnCount++)
	cell.setCellValue(var)
}
List<WebElement> cost_row_count = driver.findElements(By.xpath('//div[@class="ppm_gridcontent"]//tbody/tr[@class="ppm_aggregation"]'))
for(int k=0;k<cost_row_count.size();k++)
{
List<WebElement> cost_Rows = cost_row_count.get(k).findElements(By.tagName('td'));

List<String> cost_data = new ArrayList<String>();
HSSFRow row1 = sheet.createRow(rowCount++)
int columnCount1 = 0
row1.createCell(columnCount1++).setCellValue(cost_Rows.get(0).getText())
//println cost_Rows.size()
int a = cost_Rows.size()-20;
//println a
for(int j=a;j<cost_Rows.size();j++)
{
	cost_data.add(cost_Rows.get(j).getText());
}
println cost_data

for (def var1 : cost_data) {
	HSSFCell cell = row1.createCell(columnCount1++)
	cell.setCellValue(var1)
}
}
FileOutputStream outputStream = new FileOutputStream(filename)
workbook.write(outputStream)
outputStream.close()
System.out.println("Your excel file has been generated!");
}
catch (IOException e) {
	e.printStackTrace()
}
