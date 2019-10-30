package ReadExcelData;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import Research.ExcelDataConfig;

@Test
public class ReadExcel {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		WebDriver driver = new ChromeDriver();
		driver.get("http://10.50.50.19:8083/Cronos");
		File src = new File("D:\\WORK\\Selenium Testing\\Project\\AutomationTest\\AutomationTest\\Test Files\\Test Script.xlsx");
		
		FileInputStream fis = new FileInputStream(src);
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		
		XSSFSheet sheet1 = wb.getSheetAt(0);
		
		String action 		= sheet1.getRow(1).getCell(0).getStringCellValue();
		String dataInput 	= sheet1.getRow(1).getCell(1).getStringCellValue();
		String path 		= sheet1.getRow(1).getCell(2).getStringCellValue();
		
		
		//find selenium command
		if(action.equalsIgnoreCase("sendKeys")){

			driver.findElement(By.xpath(path)).sendKeys(dataInput); 
		}else{
			
			System.out.println("TIDAK ADA ACTION");
		}
		
		
		
		
		wb.close();
	}

}
