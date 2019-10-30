package ReadExcelData;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.NoSuchElementException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;


public class ReadExcel1 {
	static double a[];
	
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		WebDriver driver = new ChromeDriver();
		//Get URL
		driver.get("http://opensource.demo.orangehrmlive.com/");
		//maximize windows screen
		driver.manage().window().maximize();
		
		File src = new File("D:\\WORK\\Selenium Testing\\Project\\AutomationTest\\AutomationTest\\Test Files\\Test Script.xlsx");
		
		FileInputStream fis = new FileInputStream(src);
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		
		XSSFSheet sheet1 = wb.getSheetAt(0);
		
		int rowcount = sheet1.getLastRowNum();

		//System.out.println("Total row is "+ rowcount);

		/* set variable for action, dataInput and xpath from EXCEL
		String action 		= sheet1.getRow(1).getCell(0).getStringCellValue();
		String dataInput 	= sheet1.getRow(1).getCell(1).getStringCellValue();
		String path 		= sheet1.getRow(1).getCell(2).getStringCellValue();
		*/
		
		//get all action, dataInput and xpath from excel
		for(int a=0;a<rowcount+1;a++)
		{
							
				//SELENIUM COMMAND
				//find sendKeys command
				if(sheet1.getRow(a).getCell(3).getStringCellValue().equalsIgnoreCase("sendKeys")){
					try{
						driver.findElement(By.xpath(sheet1.getRow(a).getCell(5).getStringCellValue())).sendKeys(sheet1.getRow(a).getCell(4).getStringCellValue());
						sheet1.getRow(a).createCell(6).setCellValue("Pass");
					}
					catch(Exception e){
						sheet1.getRow(a).createCell(6).setCellValue("Fail");
						sheet1.getRow(a).createCell(7).setCellValue(e.getMessage());
					}
					
				} 
				//find click command
				else if(sheet1.getRow(a).getCell(3).getStringCellValue().equalsIgnoreCase("click")){
					
					try{
						driver.findElement(By.xpath(sheet1.getRow(a).getCell(5).getStringCellValue())).click();
						sheet1.getRow(a).createCell(6).setCellValue("Pass");
					}
					catch(Exception e){
						sheet1.getRow(a).createCell(6).setCellValue("Fail");
						sheet1.getRow(a).createCell(7).setCellValue(e.getMessage());
					}
					
				}
				//find click command
				else if(sheet1.getRow(a).getCell(3).getStringCellValue().equalsIgnoreCase("clickAndWait")){
					
					try{
						
						driver.findElement(By.xpath(sheet1.getRow(a).getCell(5).getStringCellValue())).click();
						WebDriverWait wait = new WebDriverWait(driver, 2);
						 
				        // Wait for Alert to be present
				 
				        Alert myAlert = wait.until(ExpectedConditions.alertIsPresent());
						sheet1.getRow(a).createCell(6).setCellValue("Pass");
					}
					catch(Exception e){
						sheet1.getRow(a).createCell(6).setCellValue("Fail");
						sheet1.getRow(a).createCell(7).setCellValue(e.getMessage());
					}
					
				}
				//if there is no command found
				else{
					
					System.out.println("ACTION TIDAK DITEMUKAN");
				}
			
		}
		
		FileOutputStream fout = new FileOutputStream(src);
		
		wb.write(fout);
		wb.close();
		
	   
		
	}

}
