package ReadExcelData;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import com.google.common.io.FileBackedOutputStream;

public class writeExcel {
	static double a[];
	public static void main(String[] args) throws Exception {
		
		File src = new File("D:\\WORK\\Selenium Testing\\Project\\AutomationTest\\AutomationTest\\Test Files\\Test Script.xlsx");
		
		FileInputStream fis = new FileInputStream(src);
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		
		XSSFSheet sheet1 = wb.getSheetAt(0);
		
		sheet1.getRow(1).createCell(3).setCellValue("Pass");
		
		sheet1.getRow(2).createCell(3).setCellValue("Fail");
		
		FileOutputStream fout = new FileOutputStream(src);
		
		wb.write(fout);
		
		wb.close();
	   
		
	}

}
