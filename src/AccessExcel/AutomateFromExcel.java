package AccessExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class AutomateFromExcel {

	public static void main(String[] args) throws IOException, Exception {
		// TODO Auto-generated method stub
		File src = new File("C:\\Users\\Gaurav\\eclipse-workspace\\DataDrivenTrainng\\src\\ExcelFile\\ExcelWorkBook.xlsx");
		FileInputStream finput = new FileInputStream(src);
		XSSFWorkbook workbook;
		XSSFSheet sheet;
		XSSFCell cell;
		workbook = new XSSFWorkbook(finput);
		sheet = workbook.getSheetAt(0);
		System.setProperty("webdriver.chrome.driver", "D:\\Selenium\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.get("file:///C:/Users/Gaurav/Desktop/Selenium/Test.html");
		WebElement fname = driver.findElement(By.cssSelector("#fname"));
		WebElement lname = driver.findElement(By.cssSelector("#lname"));
		WebElement submitButton = driver.findElement(By.cssSelector("#idOfButton"));
		System.out.println("Last row number in sheet is: " + sheet.getLastRowNum());
		for (int i = 1; i <= sheet.getLastRowNum(); i++) {
			cell = sheet.getRow(i).getCell(1);
			cell.setCellType(CellType.STRING);
			System.out.println("For row: " + i + " Cell value for first column is: " + cell);
			fname.sendKeys(cell.getStringCellValue());
			cell = sheet.getRow(i).getCell(2);
			cell.setCellType(CellType.STRING);	
			System.out.println("For row: " + i + " Cell value for second column is: " + cell);
			lname.sendKeys(cell.getStringCellValue());
			submitButton.click();
			Thread.sleep(2000);
			fname.clear();			lname.clear();
			//Validation, Assertion, based on Submitting records, if passes, put pass in excel
			//if fails, put fail in excel
			String result = "pass";
			sheet.getRow(i).createCell(3).setCellValue(result);
			FileOutputStream foutput = new FileOutputStream(src);
			workbook.write(foutput);
			foutput.close();
		}
		workbook.close();
		driver.close();
	}
}