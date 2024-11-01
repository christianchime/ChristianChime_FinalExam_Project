package BusyQA.ChrisExam;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class ResultsPagePOM {
	WebDriver driver;
	@FindBy(xpath = "//a[@href='#OBLIGATIONS']") WebElement bondsLink;
	@FindBy(xpath = "//div[@id='report_OBLIGATIONS']//table[@class='t-Report-report']") WebElement webTable;
	@FindBy(xpath = "//table[2]/preceding-sibling::p[1]") WebElement heading;
	@FindBy(xpath = "//table[2]") WebElement excelTable;

	List<WebElement> webTableBody = new ArrayList<>(); 
	List<WebElement> rows = new ArrayList<>(); 
	List<WebElement> links = new ArrayList<>();
	List<WebElement> columns = new ArrayList<>();
	
	List<WebElement> excelTableRows = new ArrayList<>();
	List<WebElement> excelTableColumns = new ArrayList<>();
	ArrayList<String> columnContents = new ArrayList<>();
	ArrayList<String> rowContents = new ArrayList<>();
	ArrayList<String> linkNames = new ArrayList<>();
	ArrayList<String> allLinkNames = new ArrayList<>();
	ArrayList<Integer> records = new ArrayList<>();
	ArrayList<String> fileNames = new ArrayList<>();
	
	
	WebElement wtb;
	WebElement row;
	WebElement link;
	WebElement column1;
	WebElement iframe;
	WebElement excelTableRow;
	WebElement excelTableColumn;
	
	String headingText;
	String columnContent;
	String outputExcel;
	String updateExcel;
	
	FileOutputStream outputStream;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	XSSFCellStyle boldStyle;
	XSSFFont font;
	XSSFCellStyle centeredBoldStyle;
	//int numberOfTables = 5;
	
	ExtentReport er; //= new ExtentReport(driver);

	public ResultsPagePOM(WebDriver driver, ExtentReport er){
	this.driver = driver;
	this.er = er;
	PageFactory.initElements(driver, this);
	}

	public String bondsGetText() {
		return bondsLink.getText();
	}
	public void bondsClick() throws IOException, InterruptedException {
		bondsLink.click();
		Thread.sleep(1000);
		er.resultPageReport();
	}
	
	public void getWebTables(int numberOfTables) {
		webTableBody = webTable.findElements(By.tagName("tbody"));
		System.out.println("Number of table bodies in the web table:" + webTableBody.size());
		
		for (int i = 0; i <= numberOfTables; i++) { //(int i = 0; i < tbody.size(); i++) {
	          wtb = webTableBody.get(i);
	          // Find all rows in the table body
	          rows = wtb.findElements(By.tagName("tr"));
	          records.add(rows.size()-2); //number of records in table
	          System.out.println("\nNumber of rows in tbody " + i + " is " + rows.size() + " ==> " + records.get(i) + "records");

	          for (int j = 0; j <= rows.size() - 2; j++) { // -3 because we don't need the last two rows
	        	  row = rows.get(j);
	        	  // Find all columns in the each row
		          columns = row.findElements(By.tagName("td"));
		          column1 = columns.get(0);
		          if (j <= (rows.size()-3)) {
		        	  linkNames.add(j,column1.getText());
		        	  link = column1.findElement(By.tagName("a")); //the row with the date has no a tag but i need it for the file name
		        	  System.out.println("\tNumber of columns in row " + j + " is " + columns.size() + "\t first column is: " + linkNames.get(j));
		        	  links.add(link);
		        	  allLinkNames.add(linkNames.get(j));
		          }
		          else
		        	  fileNames.add(column1.getText());
	          }
	      }
		System.out.println("Total number of linkNames: " + allLinkNames.size());
		for (int p = 0; p <= allLinkNames.size() - 1; p++) {
			System.out.println(allLinkNames.get(p));
		}
		System.out.println("Total number of records " + records.size());
		for (int p = 0; p <= records.size() - 1; p++) {
			System.out.println(records.get(p));
		}
		System.out.println("Total number of file names " + fileNames.size());
		for (int p = 0; p <= fileNames.size() - 1; p++) {
			System.out.println(fileNames.get(p));
		}
	}
	
	int z = 0;
	public void createExcelFilesWithSheets() throws InterruptedException {
		for (int x = 0; x <= fileNames.size() - 2; x++) {
			outputExcel = "/Users/a0000/eclipse-workspace/ChrisExam/ExcelFiles/" + fileNames.get(x).replaceAll("[/:*?\"<>|!]", "") + ".xlsx";
			try (XSSFWorkbook workbook = new XSSFWorkbook()) {
                // Add sheets to the workbook 
				Thread.sleep(2000);
                for ( int y = 1; y <= records.get(x + 1); y++) {
                	sheet = workbook.createSheet(allLinkNames.get(z));
         		 	boldStyle = workbook.createCellStyle();
         		 	font = workbook.createFont();
         		 	centeredBoldStyle = workbook.createCellStyle();
         		 	System.out.println("\tAdded sheet: " + allLinkNames.get(z)+ " to " + fileNames.get(x).replaceAll("[/:*?\"<>|!]", "") ); 
         		 	Thread.sleep(2000);
         		 	frameTable(z, sheet);
         		 	z++;
                }
                // Save the workbook to the specified file
                try (FileOutputStream fos = new FileOutputStream(outputExcel)) {
                    workbook.write(fos);
                    System.out.println("Created Excel file with sheets: " + outputExcel);
                } catch (IOException e) {
                    e.printStackTrace();
                }
                
			}   catch (IOException e) {
                e.printStackTrace();
            }
		}
		
		
	}
	
//	public void populateSheetWithData(Sheet sheet, int startingRowIndex, ArrayList<String> data) {
//        for (int i = 0; i < data.size(); i++) {
//            Row row = sheet.createRow(startingRowIndex + i); // Create each row
//            Cell cell = row.createCell(0); // Populate cell in the first column
//            cell.setCellValue(data.get(i)); // Set cell value to the string in ArrayList
//        }
//        System.out.println("Populated sheet: " + sheet.getSheetName());
//    }
	
	public void frameTable(int a, XSSFSheet sheet) throws InterruptedException, IOException {
		String originalWindow = driver.getWindowHandle();
		//er.resultPageReport();
		links.get(a).click();
		 Thread.sleep(2000);
		 er.iframeWindowReport();
		 
		 for (String windowHandle : driver.getWindowHandles()) {
			    if (!windowHandle.equals(originalWindow)) {
			        driver.switchTo().window(windowHandle); // Switch to the new window
			        break;
			    }
			}
		 iframe = driver.findElement(By.xpath("//div[@class='ui-dialog-content ui-widget-content js-dialogReady']//iframe"));

		 driver.switchTo().frame(iframe);
		 Thread.sleep(2000);
		 int lastIndex = heading.getText().lastIndexOf("\n");
		 headingText = (lastIndex != -1) ? heading.getText().substring(lastIndex + 1) : heading.getText();
		 System.out.println("\n" + headingText);
		 
		 //put the header in the sheet
		 int rowNum = 0;
		 XSSFRow headerRow = sheet.createRow(rowNum);

		 headerRow.createCell(0).setCellValue(headingText);
         excelTableRows = excelTable.findElements(By.tagName("tr"));
		 System.out.println("Number of rows in the excel table is :" + excelTableRows.size());
		 
		 for (int k = 0; k <= excelTableRows.size() - 1 ; k++) { 
	          excelTableRow = excelTableRows.get(k);
	          // Find all columns in each row
	          excelTableColumns = excelTableRow.findElements(By.tagName("td"));
	          System.out.println("");//Number of columns in row " + k + " is " + excelTableColumns.size());
	          
	          XSSFRow row = sheet.createRow(rowNum + 1 + k);
	          
	          for (int l = 0; l <= excelTableColumns.size() - 1; l++) {
	        	  excelTableColumn = excelTableColumns.get(l);
	        	  columnContent = excelTableColumn.getText();
		          System.out.print("\t\t" + columnContent);
		          columnContents.add(columnContent);
		          
		          XSSFCell cell = row.createCell(l);
		          cell.setCellValue(columnContent);
	          }
	          
	          //put the column contents into rows *not sure this works
	          //rowContents.addAll(columnContents);
		 }
			// Merge cells from row 0, column 0 to row 0(the second zero), column 3
         sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
         // Set specific width for column 0 (width is 15 characters wide)
         sheet.setColumnWidth(0, 80 * 256); // Multiply by 256 to set width in units
         sheet.setColumnWidth(1, 30 * 256); // Multiply by 256 to set width in units
         sheet.setColumnWidth(2, 30 * 256); // Multiply by 256 to set width in units
		 font.setBold(true);
         boldStyle.setFont(font);
         centeredBoldStyle.setFont(font);
         centeredBoldStyle.setAlignment(HorizontalAlignment.CENTER);

		 headerRow.setRowStyle(centeredBoldStyle);
		 driver.switchTo().defaultContent();
		 driver.findElement(By.xpath("//span[@class='ui-button-icon ui-icon ui-icon-closethick']")).click();	
		 driver.switchTo().window(originalWindow);
//      try {
//       //  driver.navigate().back();
//        // Thread.sleep(2000);
//			
//       }
//       catch(Exception e) {}
		
	}
}
