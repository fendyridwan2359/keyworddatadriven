package AutomationFramework;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
 
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
 
import org.testng.annotations.DataProvider;
 
import org.testng.annotations.Test;
 
public class DataProviderTest {
 
	private static WebDriver driver;
 
  @DataProvider(name = "Authentication")
 
  public static Object[][] credentials() {
 
        return new Object[][] { { "admin3", "2" }};
 
  }
 
  // Here we are calling the Data Provider object with its Name
 
  @Test(dataProvider = "Authentication")
 
  public void test(String sUsername, String sPassword) {
 
	  driver = new ChromeDriver();
 
      driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
 
      driver.get("http://10.50.50.19:8083/Cronos");
      
      
 
      	driver.findElement(By.xpath("//*[@id='username']")).sendKeys(sUsername); 
      
 		//Type LastName in the LastName text box
 		driver.findElement(By.xpath("//*[@id='password']")).sendKeys(sPassword);
 		
 		driver.findElement(By.xpath("//*[@id='login']")).click();
 
      //driver.quit();
 
  }
 
}
