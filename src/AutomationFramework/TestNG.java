package AutomationFramework;

import org.testng.annotations.Test;
import org.testng.annotations.BeforeMethod;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;

public class TestNG {
	
	public WebDriver driver;
  @Test
  public void main() {
	// Find the element that's ID attribute is 'account'(My Account)
	  
 
      driver.findElement(By.id("username")).sendKeys("fendy");
 
      driver.findElement(By.id("password")).sendKeys("123456789");
 
      driver.findElement(By.id("login")).click();
 
 
  }
  @BeforeMethod
  public void beforeMethod() {
	  // Create a new instance of the Firefox driver
	  
      driver = new ChromeDriver();
 
      //Put a Implicit wait, this means that any search for elements on the page could take the time the implicit wait is set for before throwing exception
 
      driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
 
      //Launch the Online Store Website
 
      driver.get("http://10.50.50.19:8083/Cronos");
  }

  @AfterMethod
  public void afterMethod() {
	// Close the driver
	  
      //driver.quit();
  }

}
