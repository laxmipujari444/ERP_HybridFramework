package driverFactory;

import org.openqa.selenium.WebDriver;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import commonFunction.FuctionLibraries;
import utilities.ExcelFileUtil;

public class DriverScript {
	String inputpath="./FileInput/Hybrid_Framework.xlsx";
	String outputpath="./FileOutput/Hybridresults.xlsx";
	ExtentReports report;
	ExtentTest logger;
	public static WebDriver driver;
	public void startTest() throws Throwable {
		
		String Modulestatus="";
//create object for excelfileutil class
	ExcelFileUtil xl= new ExcelFileUtil(inputpath);
	String TestCases="MasterTestCases";
	
	//iterate all rows in testcase sheet
	for(int i=1;i<=xl.rowCount(TestCases);i++) {
		if(xl.getCellData(TestCases, i, 2).equalsIgnoreCase("Y")) {
			//read all testcases or corresponding sheets
			String TCModule=xl.getCellData(TestCases, i, 1);
			System.out.println(TCModule);
			//define path of html
			report = new ExtentReports("./target/Reports/"+TCModule+FuctionLibraries.generateDate()+".html");
			logger=report.startTest(TCModule);
			//iterate all rows in TCModule sheet
			for(int j=1;j<=xl.rowCount(TCModule);j++) {
				//read all cells from TCModule
				String Description=xl.getCellData(TCModule, j, 0);
				String Object_Type=xl.getCellData(TCModule, j, 1);
				String Locator_Type=xl.getCellData(TCModule, j, 2);
				String Locator_Value=xl.getCellData(TCModule, j, 3);
				String Test_Data=xl.getCellData(TCModule, j, 4);
				try {
					if(Object_Type.equalsIgnoreCase("startBrowser")) {
						//driver=FunctionLibraries.startBrowser();
						driver=FuctionLibraries.startBrowser();
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("openUrl")) {
						FuctionLibraries.openUrl();
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("waitForElement")) {
						FuctionLibraries.waitForElement(Locator_Type, Locator_Value, Test_Data);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("typeAction")) {
						FuctionLibraries.typeAction(Locator_Type, Locator_Value, Test_Data);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("clickAction")) {
						FuctionLibraries.clickAction(Locator_Type, Locator_Value);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("validateTitle")) {
						FuctionLibraries.validateTitle(Test_Data);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("closeBrowser")) {
						FuctionLibraries.closeBrowser();
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("dropDownAction")) {
						FuctionLibraries.dropDownAction(Locator_Type, Locator_Value, Test_Data);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("captureStockNum")) {
						FuctionLibraries.captureStockNum(Locator_Type, Locator_Value);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("stockTable")) {
						FuctionLibraries.stockTable();
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("capturesup"))
					{
						FuctionLibraries.capturesup(Locator_Type, Locator_Value);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("supplierTable"))
					{
						FuctionLibraries.supplierTable();
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("capturecus"))
					{
						FuctionLibraries.capturecus(Locator_Type, Locator_Value);
						logger.log(LogStatus.INFO, Description);
					}
					if(Object_Type.equalsIgnoreCase("customerTable"))
					{
						FuctionLibraries.customerTable();
						logger.log(LogStatus.INFO, Description);
					}
					//write as pass to status column in TCModule
					xl.setCellData(TCModule, j,5, "Pass", outputpath);
					logger.log(LogStatus.PASS, Description);
					Modulestatus="True";
				}catch(Exception e) {
					System.out.println(e.getMessage()+"jjj");
					xl.setCellData(TCModule, j,5, "Fail", outputpath);
					logger.log(LogStatus.FAIL, Description);
					Modulestatus="False";
				}
				if(Modulestatus.equalsIgnoreCase("True")) {
					//write as pass into testcases sheet
					xl.setCellData(TestCases, i, 3, "Pass", outputpath);
				}
				else {
					//write as fail into testcases sheet
					xl.setCellData(TestCases, i, 3, "Fail", outputpath);	
					}	
				report.endTest(logger);
				report.flush();
			}
		}
		else{
			//write as blocked into status cell
			xl.setCellData(TestCases, i, 3, "Blocked",outputpath );
		}
	}
	}
}
