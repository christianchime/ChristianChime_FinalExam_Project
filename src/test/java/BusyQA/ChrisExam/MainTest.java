package BusyQA.ChrisExam;

import org.testng.annotations.Test;

//import testNGFramework.LoggerFile;

//import testNGFramework.ResultsPagePOM;

import org.testng.annotations.BeforeTest;

import java.io.IOException;
import java.time.Duration;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;

public class MainTest {
	static Logger logger = Logger.getLogger(MainTest.class);
	WebDriver driver;
	ExtentReport er;
  @Test
  public void f() throws InterruptedException, IOException {
	  ResultsPagePOM rp = new ResultsPagePOM(driver, er);
	  
	  Thread.sleep(2000);
	  rp.bondsClick();
	  logger.info("Clicked on: " + rp.bondsGetText());
	  Thread.sleep(1000);
	  rp.getWebTables(5);
	  rp.createExcelFilesWithSheets();
	  
  }
  @BeforeTest
  public void beforeTest() throws InterruptedException {
	  PropertyConfigurator.configure("src/test/resources/log4j.properties");
	  driver = new ChromeDriver();
	  driver.manage().window();//.maximize();
	  driver.get("https://www.finmun.finances.gouv.qc.ca/finmun/f?p=100:3000::RESLT");
	  driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
	 
	  er = new ExtentReport(driver);
	  er.initializer();
  }

  @AfterTest
  public void afterTest() {
	  ExtentReport.eReport.flush();  // Ensures the report is written out
	    if (driver != null) {
	        driver.quit();
	    }
  }

}
