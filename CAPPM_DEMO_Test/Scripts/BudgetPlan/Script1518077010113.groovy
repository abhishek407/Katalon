import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.checkpoint.CheckpointFactory as CheckpointFactory
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
import org.apache.poi.hssf.usermodel.HSSFFont
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Color
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFFont
import org.junit.After
import org.openqa.selenium.By
import org.openqa.selenium.Keys as Keys
import org.openqa.selenium.WebDriver
import org.openqa.selenium.WebElement


WebUI.openBrowser('')

WebUI.maximizeWindow()

WebUI.navigateToUrl('https://cppm9242-dev.ondemand.ca.com/niku/nu')

WebUI.setText(findTestObject('Page_CA PPM State Street Developmen/input_userName'), 'Ramyasree.SriNidadavolu@contractor.ca.com')

WebUI.setText(findTestObject('Page_CA PPM State Street Developmen/input_passWord'), 'Clarity1')

WebUI.sendKeys(findTestObject('Page_CA PPM State Street Developmen/input_passWord'), Keys.chord(Keys.ENTER))

WebUI.delay(3)

WebUI.click(findTestObject('Page_CA PPM State Street Developmen/span_Home'))


WebUI.delay(2)

WebUI.click(findTestObject('Page_CA PPM State Street Developmen/a_Projects'))

WebUI.click(findTestObject('Page_CA PPM State Street Developmen/img_page_projmgr.projectList_o'))

WebUI.setText(findTestObject('Page_CA PPM State Street Developmen/input_unique_code'), 'PRJ00079')

WebUI.click(findTestObject('Page_CA PPM State Street Developmen/button_Filter'))


WebUI.delay(2)

WebUI.click(findTestObject('Page_CA PPM State Street Developmen/a_Beacon - Demo'))


WebUI.delay(1)
WebUI.click(findTestObject('Page_CA PPM State Street Developmen/a_Financial Plans'))

WebUI.delay(2)
WebUI.click(findTestObject('Object Repository/Page_CA_PPM_Budget_Plan_Demo/Budget_Plan'))

WebUI.delay(3)
WebDriver driver = DriverFactory.getWebDriver()
List<String> table_heading_values = new ArrayList<String>();
List<WebElement> table_heading = driver.findElements(By.xpath('//table[@id="portlet-table-revmgr.costplanList"]//table/thead//th'))


for (var in table_heading) 
{
	table_heading_values.add(var.getText());
}
println table_heading_values

int c_current = table_heading_values.indexOf("Current");
int c_budget_name = table_heading_values.indexOf("Benefit Plan");
List<WebElement> body_row = driver.findElements(By.xpath('//table[@id="portlet-table-revmgr.costplanList"]//table/tbody/tr'))

for(int i=0;i<body_row.size();i++)
{
	List<WebElement> body_column = body_row.get(i).findElements(By.tagName("td"));
	String current_value = body_column.get(c_current);
	println current_value
	if(current_value.equalsIgnoreCase("Yes"))
	{
		body_column.get(c_budget_name).findElement(By.tagName("a")).click();
		break;
	}
}
WebUI.delay(2)
WebUI.click(findTestObject('Object Repository/Page_CA_PPM_Budget_Plan_Demo/Detail_click'))

WebUI.delay(2)
WebUI.click(findTestObject('Page_CA PPM State Street Developmen (1)/img_caui-workspaceWorkspaceHea'))

WebUI.click(findTestObject('Page_CA PPM State Street Developmen (1)/a_Time-scaled Value'))

WebUI.delay(2)

WebUI.click(findTestObject('Page_CA PPM State Street Developmen (1)/input_startDateType'))

WebUI.setText(findTestObject('Page_CA PPM State Street Developmen (1)/input_startDate'), '1/1/2016')

WebUI.click(findTestObject('Page_CA PPM State Street Developmen (1)/input_timeScaleType'))

WebUI.selectOptionByValue(findTestObject('Page_CA PPM State Street Developmen (1)/select_timeScale'), 'year', true)

WebUI.setText(findTestObject('Page_CA PPM State Street Developmen (1)/input_numberOfTimePeriods'), '6')

WebUI.delay(3)

WebUI.click(findTestObject('Page_CA PPM State Street Developmen (1)/button_Save And Return'))


WebUI.delay(5)

try
{
String filename = "D:\\letzNav\\output_file1.xls" ;
HSSFWorkbook workbook = new HSSFWorkbook();
HSSFSheet sheet = workbook.createSheet("FirstSheet")
int rowCount = 0

List<String> table_heading_budget_values = new ArrayList<String>();
List<WebElement> table_heading_budget = driver.findElements(By.xpath('//*[@id="grid-content-revmgr.benefitplandetailList"]/table/thead/tr[1]/th'))

for (var1 in table_heading_budget)
{
	table_heading_budget_values.add(var1.getText());
}
println table_heading_budget_values

int c_business_unit = table_heading_budget_values.indexOf("Business Unit");
println c_business_unit
List<String> table_heading_year_values = new ArrayList<String>();
List<WebElement> table_heading_year_budget = driver.findElements(By.xpath('//*[@id="grid-content-revmgr.benefitplandetailList"]/table/thead/tr[2]/th'))

for (var2 in table_heading_year_budget)
{
	table_heading_year_values.add(var2.getText());
}
println table_heading_year_values

List<String> bu = new ArrayList<String>();
List<List<Integer>> years = new ArrayList<List<Integer>>();
String xpath1 ='//*[@id="grid-content-revmgr.benefitplandetailList"]/table/tbody/tr[not(@class="ppm_aggregation")]';

List<WebElement> table_row_budget = driver.findElements(By.xpath(xpath1));

for(int j=0;j<table_row_budget.size();j++)
{	
	int v= j+1
	String xpath2 = xpath1+"["+v+"]/td[not(@class='tableContent ppm_tsv_cell')]";
	println xpath2
	List<WebElement> body_column_budget = driver.findElements(By.xpath(xpath2));
	///List<WebElement> body_column_budget = table_row_budget.get(j).findElements(By.xpath());
	println body_column_budget
	String business_value = body_column_budget.get(c_business_unit).getText();
	if(business_value == "")
	{
		business_value = "empty";
	}
	println business_value
	int test1 = body_column_budget.size()-1;
	List<WebElement> business_value_benefit = body_column_budget.get(test1).findElements(By.xpath(xpath2+"//td"));
	println business_value_benefit
	int count = -1;
	for (q=0;q<business_value_benefit.size();q++) 
	{
		
		String business_value_benefit_test = business_value_benefit.get(q).getText();
		if(business_value_benefit_test.equalsIgnoreCase("FTE Planned"))
		{
			count = q+1;
		}
	}
	List<Integer> period_data = new ArrayList<Integer>();
	String path1 = '//tr[';
	int path2 = j+1;
	String path3 = ']/td[@class="tableContent ppm_tsv_cell"]//tr[';
	String path4 = "]"
	String final_path = path1 + path2 + path3 + count + path4;
	println final_path;
	List<WebElement> body_column_budget1 = driver.findElements(By.xpath(final_path));
	
	for(int z=0;z<body_column_budget1.size();z++)
	{
		String test = body_column_budget1.get(z).getText().replace(",", "");
		if(test == "")
		{
			test = "0";
		}
		period_data.add(Integer.parseInt(test));
	}
	println period_data
	if(bu.contains(business_value))
	{
		int a =bu.indexOf(business_value);
		List<Integer> period_data1 = new ArrayList<Integer>();
		for(p=0;p<period_data.size();p++)
		{
			Integer b = years.get(a).get(p)+period_data.get(p);
			period_data1.add(b);
		}
		years.set(a, period_data1);
	}
	else
	{
		bu.add(business_value);
		years.add(period_data);
	}	
}

HSSFRow row = sheet.createRow(rowCount++)
int columnCount = 0
row.createCell(columnCount++).setCellValue("Business_unit")
for (def var2 : table_heading_year_values) {
	HSSFCell cell = row.createCell(columnCount++)
	cell.setCellValue(var2)
}
println bu;
println years;

for (int u=0;u<bu.size();u++)
{
int columnCount1 = 0
HSSFRow row1 = sheet.createRow(rowCount++)
row1.createCell(columnCount1++).setCellValue(bu.get(u))

CellStyle style = workbook.createCellStyle();
    Font font = workbook.createFont();
    font.setBoldweight(Font.BOLDWEIGHT_BOLD);
    style.setFont(font);
	font.setColor(IndexedColors.RED.getIndex())
   
   // Setting cell style
   
   
for (def var3 : years.get(u)) {
	HSSFCell cell = row1.createCell(columnCount1++)
	cell.setCellValue(var3)
	cell.setCellStyle(style);
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

