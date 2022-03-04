package week5.day2.assessment;

import java.io.IOException;
import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Parameters;

import io.github.bonigarcia.wdm.WebDriverManager;
import week5.day2.ReadExcel;

public class LeadsBase {
	public ChromeDriver driver;
	public String sheet1, sheet2, sheet3, sheet4, sheet5;

	
	@BeforeClass
	public void setData() {
		sheet1 = "Create Lead";
		sheet2 = "Edit Lead";
		sheet3 = "Merge Lead";
		sheet4 = "Delete Lead";
		sheet5 = "Duplicate Lead";
	}

	@Parameters({"url", "username" , "password"})
	@BeforeMethod
	public void precondition(String URL, String Name, String Pass) {
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver(); 
		driver.manage().window().maximize();
		driver.get(URL);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
		driver.findElement(By.id("username")).sendKeys(Name);
		driver.findElement(By.id("password")).sendKeys(Pass);
		driver.findElement(By.className("decorativeSubmit")).click();
		driver.findElement(By.linkText("CRM/SFA")).click();
		driver.findElement(By.linkText("Leads")).click();		
	}
	
	@AfterMethod
	public void postcondition() {	
		driver.close();
}
	@DataProvider(name="createLeadData")
	public String[][] sendCreateData() throws IOException {
		return ExcelData.fetchCreate(sheet1);
	}
	@DataProvider(name="editLeadData")
	public String[][] sendEditData() throws IOException {
		return ExcelData.fetchEdit(sheet2);	
	}
	@DataProvider(name="mergeLeadData")
	public String[][] sendMergeData() throws IOException {
		return ExcelData.fetchMerge(sheet3);	
	}
	@DataProvider(name="deleteLeadData")
	public String[][] sendDeleteData() throws IOException {
		return ExcelData.fetchDelete(sheet4);	
	}
	@DataProvider(name="duplicateLeadData")
	public String[][] sendDuplicateData() throws IOException {
		return ExcelData.fetchDuplicate(sheet5);	
	}
}