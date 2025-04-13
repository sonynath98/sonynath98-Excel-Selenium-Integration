package uploadDownload.ProjectExplores;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.Test;

public class UploadDownloadd {
	
	@Test
	public void downloadUpdateExcel() throws IOException {
			WebDriver driver = new ChromeDriver();
			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(5));
			driver.get("https://rahulshettyacademy.com/upload-download-test/index.html");
			
			String fruitName = "Apple";
			String updatedPrice = "311";
			String fileName = "C:\\Users\\DELL\\Downloads\\download.xlsx";
			//Download
			driver.findElement(By.id("downloadButton")).click();
			
			//edit excel
			int col =getColumnNumber(fileName,"price");
			int row =getRowNumber(fileName,"Apple");
			Assert.assertTrue(updateCell(fileName,col,row,updatedPrice));
			
			//upload
			WebElement upload =  driver.findElement(By.cssSelector("input[type='file']"));
			upload.sendKeys(fileName);
			
			//wait for the success message to show up and wait for disappearing
			wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div[role='alert']")));
			String confirmationMessage =  driver.findElement(By.cssSelector("div[role='alert']")).getText();
			Assert.assertEquals("Updated Excel Data Successfully.",confirmationMessage);
			wait.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("div[role='alert']")));
			
			//Verify updated excel data showing in the web table
			String columnPrice = driver.findElement(By.xpath("//div[text()='Price']")).getDomAttribute("data-column-id");
			String actualPrice = driver.findElement(By.xpath("//div[text()='"+fruitName+"']/parent::div/parent::div/div[@id='cell-"+columnPrice+"-undefined']")).getText();
			System.out.println(actualPrice);
			Assert.assertEquals(actualPrice, updatedPrice);
			
			driver.close();
			
	}

	private static boolean updateCell(String fileName, int col, int row, String updatedValue) throws IOException {
		// TODO Auto-generated method stub
		FileInputStream  file = new FileInputStream(fileName);
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		Row rowField = sheet.getRow(row-1);
		Cell cellField = rowField.getCell(col-1);
		cellField.setCellValue(updatedValue);
		FileOutputStream fis = new FileOutputStream(fileName);
		workbook.write(fis);
		workbook.close();
		return true;
		
		
	}

	private static int getColumnNumber(String fileName, String colName) throws IOException {
		// TODO Auto-generated method stub
		FileInputStream  file = new FileInputStream(fileName);
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		Iterator<Row> row = sheet.iterator();
		Row firstRow = row.next();
		Iterator<Cell> cell = firstRow.cellIterator();
		int k=1;
		int column = 0;
		while(cell.hasNext()) {
			Cell value = cell.next();
		if(value.getStringCellValue().equalsIgnoreCase(colName)) {
			column=k;
		}
		k++;
		}
		return column;
		
	}

	private static int getRowNumber(String fileName, String rowName) throws IOException {
				FileInputStream  file = new FileInputStream(fileName);
				XSSFWorkbook workbook = new XSSFWorkbook(file);
				XSSFSheet sheet = workbook.getSheet("Sheet1");
				int k=1;
				int rowIndex=-1;
				Iterator<Row> row = sheet.iterator();
				
				while(row.hasNext()) {
					Row firstRow = row.next();
					Iterator<Cell> cell = firstRow.cellIterator();
					while(cell.hasNext()) {
						Cell value = cell.next();
						if(value.getCellType() == CellType.STRING && value.getStringCellValue().equalsIgnoreCase(rowName)) {
							rowIndex=k;
						}
					}
					k++;	
				}
				return rowIndex;
		
	}
	
		
		
		
	

}
