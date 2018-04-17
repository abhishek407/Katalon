import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject

import java.awt.List

import org.openqa.selenium.By
import org.openqa.selenium.Keys
import org.openqa.selenium.WebDriver
import org.openqa.selenium.Keys
import org.openqa.selenium.WebElement

import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.checkpoint.CheckpointFactory as CheckpointFactory
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as MobileBuiltInKeywords
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testcase.TestCaseFactory as TestCaseFactory
import com.kms.katalon.core.testdata.ExcelData
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

ExcelData data = (ExcelData) findTestData("Data Files/New Test Data");

println data.getAllData()
int row_count = data.getRowNumbers()


List<String> data_list = new ArrayList<String>();

Map<String, String> data_map = new LinkedHashMap<String, String>();

for(int i=1;i<=row_count;i++)
{
	data_map.put(data.getValue(1, i),data.getValue(2, i))
}

nav_url = data_map.get("AdminUrl");
user_name = data_map.get("username");
nav_password =  data_map.get("password");


String downloadPath = "D:\\letzNav";

WebUI.openBrowser("");

WebUI.maximizeWindow()

WebUI.navigateToUrl(nav_url)


WebUI.delay(3)

WebUI.setText(findTestObject('Object Repository/Admin_Letznav_username'), user_name)
WebUI.setText(findTestObject('Object Repository/Admin_Letznav_password'), nav_password)

WebUI.submit(findTestObject('Object Repository/Admin_LetzNav_Login'))

WebUI.delay(3)

WebUI.click(findTestObject('Object Repository/DownloadEditor'))
/*WebUI.click(findTestObject('Object Repository/DownloadPlayer'))*/
WebUI.delay(10)

WebDriver driver = DriverFactory.getWebDriver()
driver.findElement(By.className("manage-page-header")).sendKeys(Keys.chord(Keys.CONTROL, 'l'))
driver.findElement(By.className("manage-page-header")).sendKeys(Keys.chord(Keys.TAB))
driver.findElement(By.className("manage-page-header")).sendKeys(Keys.chord(Keys.TAB))
driver.findElement(By.className("manage-page-header")).sendKeys(Keys.chord(Keys.ENTER))

/*autoit_prj = 'D:\\letzNav\\AutoIt_continue.exe'

Runtime.getRuntime().exec(autoit_prj)
Thread.sleep(3000);*/
/*WebUI.waitForAlert(10)

alertText = WebUI.getAlertText()
println alertText
WebUI.acceptAlert();*/

