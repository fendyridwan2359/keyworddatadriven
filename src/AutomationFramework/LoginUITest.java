package AutomationFramework;

import java.util.regex.Pattern;
import java.util.concurrent.TimeUnit;
import org.testng.annotations.*;
import static org.testng.Assert.*;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.events.EventFiringWebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class LoginUITest {
  private WebDriver driver;
  private WebDriverWait wait;
  private String baseUrl;
  private boolean acceptNextAlert = true;
  private StringBuffer verificationErrors = new StringBuffer();

  @BeforeClass(alwaysRun = true)
  public void setUp() throws Exception {
    driver = new ChromeDriver();
    //baseUrl = "http://opensource.demo.orangehrmlive.com/";
    baseUrl = "http://10.50.50.19:8083/Cronos";

    driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
    wait = new WebDriverWait(driver, 5);
  }

/*  @Test
  public void testLoginUI() throws Exception {
    driver.get(baseUrl);
    driver.findElement(By.id("txtUsername")).clear();
    driver.findElement(By.id("txtUsername")).sendKeys("admin");
    driver.findElement(By.id("txtPassword")).clear();
    driver.findElement(By.id("txtPassword")).sendKeys("admin");
    driver.findElement(By.id("btnLogin")).click();
    driver.findElement(By.xpath(".//*[@id='menu_pim_viewPimModule']/b")).click();
    
    //test select
    WebElement dd = driver.findElement(By.xpath(".//*[@id='empsearch_employee_status']"));
    Select selectdd = new Select(dd);
    
   // selectdd.selectByIndex(2);
    
    selectdd.selectByVisibleText("Full-Time Permanent");
   
	
  }*/
  @Test
  public void testLoginUI() throws Exception {
	  driver.get(baseUrl);
	  driver.manage().window().maximize();
	    driver.findElement(By.id("username")).clear();
	    driver.findElement(By.id("username")).sendKeys("fendy");
	    driver.findElement(By.id("password")).clear();
	    driver.findElement(By.id("password")).sendKeys("12345678");
	    driver.findElement(By.id("login")).click();
	    driver.findElement(By.linkText("Workbench")).click();
	    

  }
  
  

  @AfterClass(alwaysRun = true)
  public void tearDown() throws Exception {
    //driver.quit();
    String verificationErrorString = verificationErrors.toString();
    if (!"".equals(verificationErrorString)) {
      fail(verificationErrorString);
    }
  }

  private boolean isElementPresent(By by) {
    try {
      driver.findElement(by);
      return true;
    } catch (NoSuchElementException e) {
      return false;
    }
  }

  private boolean isAlertPresent() {
    try {
      driver.switchTo().alert();
      return true;
    } catch (NoAlertPresentException e) {
      return false;
    }
  }

  private String closeAlertAndGetItsText() {
    try {
      Alert alert = driver.switchTo().alert();
      String alertText = alert.getText();
      if (acceptNextAlert) {
        alert.accept();
      } else {
        alert.dismiss();
      }
      return alertText;
    } finally {
      acceptNextAlert = true;
    }
  }
}
