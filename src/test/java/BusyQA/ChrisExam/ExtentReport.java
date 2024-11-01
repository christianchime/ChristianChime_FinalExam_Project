package BusyQA.ChrisExam;

import java.io.File;
import java.io.IOException;
import java.time.Duration;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.AssertJUnit;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;

public class ExtentReport {
	public static ExtentSparkReporter eSparkReport;
	public static ExtentReports eReport;
	public static ExtentTest eTest;
	private WebDriver driver;
	private String screenshotPath;
	
    public ExtentReport(WebDriver driver) {
        this.driver = driver;
    }
	//WebDriver driver;
	
	public static void initializer() {
		eSparkReport =  new ExtentSparkReporter(System.getProperty("user.dir")+"/Reports/extentSparkReport.html");
		eSparkReport.config().setDocumentTitle("Automation Report");
		eSparkReport.config().setReportName("Test Execution Report");
		eSparkReport.config().setTheme(Theme.STANDARD);
		eSparkReport.config().setTimeStampFormat("yyyy-MM-dd HH:mm:ss");
		eReport = new ExtentReports();
		eReport.attachReporter(eSparkReport);		
	}
	
	public static String captureScreenshot(WebDriver driver) throws IOException {
		if (driver == null) {
            throw new IllegalArgumentException("WebDriver instance is null.");
        }
		String FileSeparator = System.getProperty("file.separator"); // "/" or "\"
		String Extent_report_path = "."+FileSeparator+"Reports"; // . means parent directory
		File Src = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
		String Screenshotname = "screenshot"+System.currentTimeMillis()+".png";
		File Dst = new File(Extent_report_path+FileSeparator+"Screenshots"+FileSeparator+Screenshotname);
		FileUtils.copyFile(Src, Dst);
		String absPath = Dst.getAbsolutePath();
//		//System.out.println("Absolute path is:"+absPath);
		return absPath;
		//File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
        //String screenshotPath = System.getProperty("user.dir") + "/Reports/Screenshots/screenshot" + System.currentTimeMillis() + ".png";
        //FileUtils.copyFile(src, new File(screenshotPath));
       // return screenshotPath;
	}
	
	
	public void resultPageReport() throws IOException {

		String methodName = new Exception().getStackTrace()[0].getMethodName();
//		String className = new Exception().getStackTrace()[0].getClassName();
		eTest = eReport.createTest(methodName,"Results Table Page");
		eTest.log(Status.INFO, "Page Access Success");
		eTest.assignCategory("Regression Testing");
	    screenshotPath = captureScreenshot(driver);
	  //  eTest.addScreenCaptureFromPath(screenshotPath);
		 // eTest.addScreenCaptureFromPath(captureScreenshot(driver));
//		  
		  String actualTitle = driver.getTitle();
//		  System.out.println("Actual Title is :"+ actualTitle);
		  String expectedTitle = "Résultats publics";
//		  
		  AssertJUnit.assertEquals(expectedTitle,actualTitle);
		//eTest = eReport.createTest("Result Page Test").assignCategory("Regression Testing");
      //  eTest.addScreenCaptureFromPath(captureScreenshot(driver));
		//  return screenshotPath;
	}

	public void iframeWindowReport() throws IOException {

		String methodName = new Exception().getStackTrace()[0].getMethodName();
//		String className = new Exception().getStackTrace()[0].getClassName();
		eTest = eReport.createTest(methodName,"Results Table Page");
		eTest.assignCategory("Regression Testing");
		eTest.addScreenCaptureFromPath(screenshotPath);
		  eTest.addScreenCaptureFromPath(captureScreenshot(driver));
		  eTest.log(Status.PASS, "Successfully captured screenshot of iframe window");
//		  String actualTitle = driver.getTitle();
//		  System.out.println("Actual Title is :"+ actualTitle);
//		  String expectedTitle = "Résultats publics";
//		  
//		  AssertJUnit.assertEquals(expectedTitle,actualTitle);
		//eTest = eReport.createTest("Iframe Window Test").assignCategory("Regression Testing");
       // eTest.addScreenCaptureFromPath(captureScreenshot(driver));
		
	}
//	@AfterTest
//	public void closeMethod() {
//		eReport.flush();
//		driver.close();
//	}
//	
//	@BeforeTest
//	public void driverSetup() {
////		initializer();
////		driver = new ChromeDriver();
////		driver.get("https://demo.guru99.com/eTest/newtours/login.php");
////		driver.manage().window().maximize();
////		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
//	}
	}

