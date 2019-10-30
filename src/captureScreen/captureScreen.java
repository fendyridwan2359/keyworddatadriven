package captureScreen;

import java.io.File;

import java.io.IOException;

import org.apache.commons.io.FileUtils;

import org.openqa.selenium.OutputType;

import org.openqa.selenium.TakesScreenshot;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.Test;

import AutomationFramework.AutomationScript;


public class captureScreen extends AutomationScript{
	
	@Test
	public void captureScreenShot(int getRowCount)
	{
		// Take screenshot and store as a file format � � � � � � 
		File src=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
		
		// now copy the� screenshot to desired location using copyFile method

		try {
			//D:\WORK\Selenium Testing\Project\AutomationTest\AutomationTest\Test Files\attachment
			//FileUtils.copyFile(src, new File("D:/WORK/Selenium Testing/Project/AutomationTest/AutomationTest/Test Files/attachment/Gambar "+getRowCount+".jpg"));
			FileUtils.copyFile(src, new File("D:/AutomationTest/Test Files/attachment/Gambar "+getRowCount+".jpg"));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
}