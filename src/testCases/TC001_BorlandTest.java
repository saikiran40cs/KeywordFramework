package testCases;

import java.io.File;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.openqa.selenium.Proxy.ProxyType;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.Reporter;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import excelExportAndFileIO.ReadKiranExcelFile;
import operation.ReadObject;
import operation.UIOperation;

/**
 * @author saikiran40cs
 * THIS IS THE EXAMPLE OF KEYWORD DRIVEN TEST CASE
 *
 */
public class TC001_BorlandTest {

	WebDriver webdriver;

	@BeforeTest
	public void browserSetup() {
		ChromeOptions chromeOptions = new ChromeOptions();
		DesiredCapabilities ChromeCapabilities = DesiredCapabilities.chrome();
		ChromeCapabilities.setCapability(ChromeOptions.CAPABILITY, chromeOptions);
		ChromeCapabilities.setCapability("network.proxy.type", ProxyType.AUTODETECT.ordinal());
		// Set ACCEPT_SSL_CERTS variable to true
		ChromeCapabilities.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
		ChromeCapabilities.setCapability(CapabilityType.ForSeleniumServer.ENSURING_CLEAN_SESSION, true); 
		Map<String, Object> prefs = new HashMap<String, Object>();
		prefs.put("profile.default_content_settings.popups", 0);
		prefs.put("download.extensions_to_open", "pdf");
		prefs.put("download.prompt_for_download", "true");
		chromeOptions.setExperimentalOption("prefs", prefs);
		System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir") + File.separator+"chromedriver_2.30.exe");
		System.setProperty("webdriver.chrome.args", "--disable-logging");
		System.setProperty("webdriver.chrome.silentOutput", "true");
		webdriver = new ChromeDriver(ChromeCapabilities);
		webdriver.manage().window().maximize();
	}

	@Test()
	public void testLogin() throws Exception {
		ReadKiranExcelFile file = new ReadKiranExcelFile();
		ReadObject object = new ReadObject();
		Properties allObjects = object.getObjectRepository();
		UIOperation operation = new UIOperation(webdriver);
		// Read keyword sheet
		Sheet Sheet = file.readExcel(System.getProperty("user.dir") + File.separator, "TestCase.xlsx","KeywordFramework");
		// Find number of rows in excel file
		int rowCount = Sheet.getLastRowNum() - Sheet.getFirstRowNum();
		// Create a loop over all the rows of excel file to read it
		for (int i = 1; i < rowCount + 1; i++) {
			// Loop over all the rows
			Row row = Sheet.getRow(i);
			// Check if the first cell contain a value, if yes, That means it is
			// the new testcase name
			if (row.getCell(0).toString().length() == 0) {
				// Print testcase detail on console
				Reporter.log(row.getCell(1).toString() + "----" + row.getCell(2).toString() + "----"+ row.getCell(3).toString() + "----" + row.getCell(4).toString(),true);
				// Call perform function to perform operation on UI
				operation.perform(allObjects, row.getCell(1).toString(), row.getCell(2).toString(),row.getCell(3).toString(), row.getCell(4).toString());
			} else {
				// Print the new testcase name when it started
				Reporter.log("New Testcase->" + row.getCell(0).toString() + " Started",true);
			}
		}
	}
	
	@AfterTest
	public void closeSession(){
		webdriver.close();
	}

}
