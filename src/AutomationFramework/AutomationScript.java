package AutomationFramework;
/**
 * @author fendy.ridwan
 * 
 *
 */
import static org.testng.Assert.assertEquals;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.events.EventFiringWebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

import captureScreen.captureScreen;

public class AutomationScript {

	private static final TimeUnit SECONDS = null;
	private static final String NULL = null;
	protected static WebDriver driver;
	static FileInputStream fis;
	static XSSFWorkbook wb;
	static XSSFSheet sheet;
	static captureScreen cs = new captureScreen();
	//static ImageAttachmentInDocument iad = new ImageAttachmentInDocument();
	public static String nameFile;
	static String getValue1 = "";
    static String getValue2 = "";
    @Test
	public static void main(String[] args) {

         
        try {
        
    		//File src = new File("D:\\WORK\\Selenium Testing\\Project\\AutomationTest\\AutomationTest\\Test Files\\Test Script OrangeHR.xlsx");
    		
        	//File src = new File("D:\\WORK\\Selenium Testing\\Project\\AutomationTest\\AutomationTest\\Test Files\\Test Script Cronos-LTKM.xlsx");
        	
        	//File src = new File("D:\\WORK\\Selenium Testing\\Project\\AutomationTest\\AutomationTest\\Test Files\\Test Script demo.xlsx");
        	
        	//File src = new File("D:\\WORK\\Selenium Testing\\Project\\AutomationTest\\AutomationTest\\Test Files\\Test Script Cronos User.xlsx");
        	
        	File src = new File("/Users/fendyridwan/Documents/Project/AutomationTest/Test Files/TestScript.xlsx");
    		fis = new FileInputStream(src);
    		
    		wb = new XSSFWorkbook(fis);
    		
    		sheet = wb.getSheetAt(0);
    		
    		int lastrownumber = sheet.getLastRowNum();
    		int getLastRowFile = sheet.getLastRowNum();
    		
    		
    		for(int count = 1;count<=lastrownumber;count++){
    			String action = "";
    	        String dataInput = "";
    	        String locator = "";
    	        String locatorType = "";
    	        
    			System.out.println("urutan ===========================>>>>>>>>>>>>>"+count);
                XSSFRow row = sheet.getRow(count);
                
                // Run the test step for the current test data row
                
                if(!(row.getCell(4) == null || row.getCell(4).equals(Cell.CELL_TYPE_BLANK))) {
                    locatorType = row.getCell(4).toString().toLowerCase();
                } else {
                	locatorType = "";
                }
                if(!(row.getCell(5) == null || row.getCell(5).equals(Cell.CELL_TYPE_BLANK))) {
                    action = row.getCell(5).toString().toLowerCase();
                } else {
                    action = "";
                }
                if(!(row.getCell(6) == null || row.getCell(6).equals(Cell.CELL_TYPE_BLANK))) {
                    dataInput = row.getCell(6).toString();
                } else {
                	dataInput = "";
                }
                if(!(row.getCell(7) == null || row.getCell(7).equals(Cell.CELL_TYPE_BLANK))) {
                	locator = row.getCell(7).toString();
                } else {
                	locator = "";
                }
                
                /* System.out.println("Locator Type: " + locatorType);
                 System.out.println("Test action: " + action);
                 System.out.println("Data Input: " + dataInput);
                 System.out.println("Locator: " + locator);*/
                 
                 //Clear Result And Description Error in output file
                 row.createCell(8).setCellValue("");
                 row.createCell(9).setCellValue("");
                 
                 //Run Command and get the Result from here
                 String result = runTestStep(locatorType,action,dataInput,locator,count,lastrownumber);
                 
                 // Write the result back to the Excel sheet                
                 row.createCell(8).setCellValue(result);
                 
                 
             }
    		
    		
            // Save the Excel sheet and close the file streams
            FileOutputStream fout = new FileOutputStream(src);
            wb.write(fout);
            fis.close();
            fout.close();
             
        } catch (Exception e) {
            System.out.println(e.toString());
        }
	}
	

		public static String runTestStep(String locatorType, String action, String dataInput, String locator, int rowCount, int lastrownumber) throws Exception 
		{
			XSSFRow row = sheet.getRow(rowCount);
			WebElement element;
			Select dropdown;
			String exePath;
		        switch(action.toLowerCase()) 
		        {
			        case "openbrowser":
			            switch(dataInput.toLowerCase()) 
			            {
				            case "firefox":
				            	exePath = "/Users/fendyridwan/Documents/Project/AutomationTest/lib/geckodriver";
				            	System.setProperty("webdriver.chrome.driver", exePath);
				                driver = new FirefoxDriver();
				                driver.manage().timeouts().implicitlyWait(4, TimeUnit.SECONDS);
				                return "PASS";
				            case "chrome":
				            	exePath = "/Users/fendyridwan/Documents/Project/AutomationTest/lib/driver/chromedriver";
				            	System.setProperty("webdriver.chrome.driver", exePath);
				                driver = new ChromeDriver();
				                driver.manage().timeouts().implicitlyWait(4, TimeUnit.SECONDS);
				                return "PASS";
				            default:
				            	//cs.captureScreenShot(rowCount);
			                    nameFile = "Gambar "+rowCount;
			                    //iad.createResult(nameFile,rowCount,lastrownumber,"FAIL");
			                    return "FAIL";
			            }
			        //From here i arrange all the command from a-z (ascending)
			            
			        //Capture Screen function
			        case "capture":
			        	//captureScreen cs = new captureScreen();
			        	 try 
			        	 {
			        		 //Capture the pass result
			        		 //cs.captureScreenShot(rowCount);
			        		 //Create Document 
			        		nameFile = "Gambar "+rowCount;
			        		//iad.createResult(nameFile,rowCount,lastrownumber,"PASS");
			                return "-";
			                 
			             } catch (Exception e) 
			        	 {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
				            return "FAIL";
			             }
			        
			        //Check content value of field
			        case "checkfield":
			        	element = findMyElement(locatorType,locator);
			        	try 
			        	 {
			        		
			        		if (element.getAttribute("value").equals(dataInput)) 
			        		{
			        			return "PASS";
			        		}
			                 
			             } catch (Exception e) 
			        	 {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
		                    return "FAIL";
			             }
			        	
			        	 
			        //Clear function
			        case "clear":
			        	 try 
			        	 {
			        		 element = findMyElement(locatorType,locator);
			        		 element.clear();
			                 return "PASS";
			                 
			             } catch (Exception e) 
			        	 {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
				            return "FAIL";
			             }
			        	 
			        //Click function
			        case "click":
			            try 
			            {
			            	element = findMyElement(locatorType,locator);
			            	element.click();
			                return "PASS";
			            } catch (Exception e) {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			            }
			            
			        //Click with locator modifier for select data table
			        case "clickmodifier1":
			            try 
			            {
			            	String locatormodifier = ".//*[@id='"+getValue1+"']/td[1]";
			            	element = findMyElement(locatorType,locatormodifier);
			            	element.click();
			                return "PASS";
			            } catch (Exception e) {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			            }
			            
			        //Click with locator modifier for select data table
			        case "clickmodifier2":
			            try 
			            {
			            	String locatormodifier = ".//*[@id='"+getValue2+"']/td[1]";
			            	element = findMyElement(locatorType,locatormodifier);
			            	element.click();
			                return "PASS";
			            } catch (Exception e) {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			            }
			            
			        //click and WAIT
			        case "clickandwait":
			            try 
			            {
			            	element = findMyElement(locatorType,locator);
			            	WebDriverWait wait = new WebDriverWait(driver, 4000);
			                wait.until(ExpectedConditions.elementToBeClickable(element)).click();
			                return "PASS";
			            	
			            } catch (Exception e) {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			            }
			            
			        //find variable 1 in the browser
			        case "findvariable1":
			        	 try 
			        	 {
			        		 element = findMyElement(locatorType,locator);
			        		 if(dataInput.equalsIgnoreCase("getvalue1")){
			        			 if (element.getText().equals(getValue1)) 
			        			 {
					                    return "PASS";
					             }
			        			 else
			        			 {
			        				 logErrorValidate(rowCount,lastrownumber,element.getText(),dataInput,"text");
			        				 return "FAIL";
			        			 }
			        		 }
			        		 
			             } catch (Exception e) 
			        	 {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
				            return "FAIL";
			             }
			        	 
			        
			        //find variable 2 in the browser
			        case "findvariable2":
			        	 try 
			        	 {
			        		 element = findMyElement(locatorType,locator);
			        		 if(dataInput.equalsIgnoreCase("getvalue2")){
			        			 if (element.getText().equals(getValue2)) 
			        			 {
					                return "PASS";
					             }
			        			 else
			        			 {
			        				logErrorValidate(rowCount,lastrownumber,element.getText(),dataInput,"text");
						            return "FAIL";
			        			 }
			        		 }
			        		 
			             } catch (Exception e) 
			        	 {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
				            return "FAIL";
			            	
			             }
			        
			        //Get variable 1
			        case "variable1":
			        	 try 
			        	 {
			        		 element = findMyElement(locatorType,locator);
			        		 getValue1 = element.getAttribute("value");
			                 return "PASS";
			                 
			             } catch (Exception e) 
			        	 {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
				            return "FAIL";
			             }
			        
			        //Get Value 2
			        case "variable2":
			        	 try 
			        	 {
			        		 element = findMyElement(locatorType,locator);
			        		 getValue2 = element.getAttribute("value");
			                 return "PASS";
			                 
			             } catch (Exception e) 
			        	 {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
				            return "FAIL";
			             }	 
			        
			        //Navigate (go to link)   
			        case "maximize":
			        	try
			        	{
			        		driver.manage().window().maximize();
				            return ("PASS");
				            
			        	}catch (Exception e){
			        		logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			        	}
			        
			        //Move Cursor   
			        case "movecursor":
			        	try
			        	{
			        		Actions act = new Actions(driver);
			        		element = findMyElement(locatorType,locator);
			        		act.moveToElement(element).perform();
				            return ("PASS");
				            
			        	}catch (Exception e){
			        		logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			        	}	
			        
			        //Navigate (go to link)   
			        case "navigate":
			        	try
			        	{
				            driver.get(dataInput);
				            return ("PASS");
				            
			        	}catch (Exception e){
			        		logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			        	}
			        
			        	
			        //Select Drop Down by index
			        case "selectbyindex":
			        	int idxDropDown = Integer.parseInt(dataInput);
			        	try
			        	{
			        		element = findMyElement(locatorType,locator);
			        		dropdown = new Select(element);
			        		dropdown.selectByIndex(idxDropDown);
				            return ("PASS");
				            
			        	}catch (Exception e)
			        	{
			        		logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			        	}
			        	
			        //Select Drop Down by VALUE
			        case "selectbyvalue":
			        	try
			        	{
			        		element = findMyElement(locatorType,locator);
			        		dropdown = new Select(element);
			        		dropdown.selectByValue(dataInput);
				            return ("PASS");
				            
			        	}catch (Exception e)
			        	{
			        		logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			        	}
			        
			        //Select Drop Down by VISIBLE TEXT
			        case "selectbyvisibletext":
			        	
			        	try
			        	{
			        		element = findMyElement(locatorType,locator);
			        		dropdown = new Select(element);
			        		dropdown.selectByVisibleText(dataInput);
				            return ("PASS");
				                                                                                                                                        
			        	}catch (Exception e)
			        	{
			        		logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			        		
			        	}
			        	
			        //Scroll
			        case "scroll":
			        	try
			        	{
			        		Thread.sleep(1000);
			        		element = findMyElement(locatorType,locator);
			        		JavascriptExecutor je = (JavascriptExecutor) driver;
			        		je.executeScript("arguments[0].scrollIntoView(true);",element);
			        		((JavascriptExecutor)driver).executeScript(dataInput, element);
				            return ("PASS");
				            
			        	} catch (Exception e)
			        	{	
			        		logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			        	}
			        
			        //scroll till find the text
			        case "scrollintoview":
			        	try
			        	{
			        		Thread.sleep(1000);
			        		element = findMyElement(locatorType,locator);
			        		JavascriptExecutor je = (JavascriptExecutor) driver;
			        		je.executeScript("arguments[0].scrollIntoView(true);",element);
			        		
				            return ("PASS");
				            
			        	} catch (Exception e)
			        	{	
			        		logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			        	}
			        	
			        //Scroll Left
			        case "scrollleft":
			        	try
			        	{
			        		Thread.sleep(1000);
			        		EventFiringWebDriver eventFiringWebDriver = new EventFiringWebDriver(driver);
			        		eventFiringWebDriver.executeScript("document.querySelector('"+locator+"').scrollLeft="+dataInput);
				            return ("PASS");
				            
			        	} catch (Exception e)
			        	{	
			        		logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			        	}
			        	
			        //Scroll down
			        case "scrolltop":
			        	try
			        	{
			        		Thread.sleep(1000);
			        		EventFiringWebDriver eventFiringWebDriver = new EventFiringWebDriver(driver);
			        		eventFiringWebDriver.executeScript("document.querySelector('"+locator+"').scrollTop="+dataInput);
			        		
				            return ("PASS");
				            
			        	} catch (Exception e)
			        	{	
			        		logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			        	}
			        
			        //Cek Text present
			        case "textpresent":
			            try 
			            {
			            	element = findMyElement(locatorType,locator);
			            	if(driver.getPageSource().contains(dataInput)){
			            		
			            		return "PASS";
			            	}else{
			            		logErrorValidate(rowCount,lastrownumber,element.getText(),dataInput,"text");
				            	return "FAIL";
			            	}
			                
			            } catch (Exception e) 
			            {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			            }
			        	
			        	
			        //TYPE (using sendKeys)
			        case "type":
			            try 
			            {
			            	element = findMyElement(locatorType,locator);
			            	element.sendKeys(dataInput);
			                return "PASS";
			                
			            } catch (Exception e) 
			            {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			            }
			            
			        //TYPE ATTRIBUTE VALUE 1
			        case "typevariable1":
			            try 
			            {
			            	element = findMyElement(locatorType,locator);
			            	element.sendKeys(getValue1);
			                return "PASS";
			                
			            } catch (Exception e) 
			            {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			            }
			         
			        //TYPE ATTRIBUTE VALUE 2
			        case "typevariable2":
			            try 
			            {
			            	element = findMyElement(locatorType,locator);
			            	element.sendKeys(getValue2);
			                return "PASS";
			                
			            } catch (Exception e) 
			            {
			            	
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			                
			            }    
			         
			        //Check element value
			        case "validate":
			        	try 
			            {
			            	if(rowCount == 326){
			            		System.out.println("stop here");
			            	}
			        		String typeOfElement = "-";
			            	element = findMyElement(locatorType,locator);
			            	typeOfElement = element.getAttribute("type");
			            	//define type of element
			            	if(typeOfElement==null)
			            	{
			            		typeOfElement = typeOfElement = "-"; 
			            	}else
			            	{
			            		typeOfElement = element.getAttribute("type");
			            	}
			            	
			            	
			            	if (typeOfElement.equals("text") || typeOfElement.equals("textarea"))
			            	{
			            		if (element.getAttribute("value").equals(dataInput)) {
				                    return "PASS";
				                }
			            		else{
			            			logErrorValidate(rowCount,lastrownumber,element.getText(),dataInput,typeOfElement);
					            	return "FAIL";
			            		}
			            	}
			            	else if(typeOfElement.equals("select-one"))
			            	{
				        		dropdown = new Select(element);
				        		String selectedValue = dropdown.getFirstSelectedOption().getText();
				        		if(selectedValue.equals(dataInput))
				        		{
				        			return "PASS";
				        		}
				        		else
				        		{
				        			logErrorValidate(rowCount,lastrownumber,selectedValue,dataInput,typeOfElement);
					            	return "FAIL";
				        		}
			            	}else if(typeOfElement.equals("checkbox"))
			            	{
			            		List<WebElement> checkbox = driver.findElements(By.xpath(locator));
			            		System.out.println(checkbox.size());
			            		for(int i=0;i<=checkbox.size();i++)
			            		{
			            			WebElement local_checkbox = checkbox.get(i);
			            			String value = local_checkbox.getAttribute("value");
			            			if(local_checkbox.isSelected())
			            			{
			            				if(value.equalsIgnoreCase(dataInput))
				            			{
				            				return "PASS";
				            			}
			            			}else
			            			{
			            				logErrorValidate(rowCount,lastrownumber,"Value of this element must be "+value,dataInput,"-");
				            			return "FAIL";
			            			}
			            			
			            		}
			            	}
			            	else
			            	{
			            		
			            		if (element.getText().equals(dataInput)) {
				                    return "PASS";
				                }
			            		else
			            		{
			            			logErrorValidate(rowCount,lastrownumber,element.getText(),dataInput,"text");
			            			return "FAIL";
			            		}
			            	}
			            
			            	
			            } catch (Exception e) {
			            	System.out.println("*****************************"+e.toString());
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			            }
			        
			        //WAIT for 2 seconds
			        case "waitfor2":
			            try 
			            {
			            	Thread.sleep(2000);
			            	return "PASS";
			            	
			            } catch (Exception e) {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			            }    
			        
			        //WAIT for 3 seconds
			        case "waitfor3":
			            try 
			            {
			            	Thread.sleep(3000);
			            	return "PASS";
			            	
			            } catch (Exception e) {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			            }    
			        
			        //WAIT for 5 seconds
			        case "waitfor5":
			            try 
			            {
			            	Thread.sleep(5000);
			            	return "PASS";
			            	
			            } catch (Exception e) {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			            }
			            
			        //WAIT for 10 seconds
			        case "waitfor10":
			            try {
			            	Thread.sleep(10000);
			            	return "PASS";
			            } catch (Exception e) {
			            	logErrorValidate(rowCount,lastrownumber,e.toString(),dataInput,"-");
			            	return "FAIL";
			            }
			            
			        //Close Browser    
			        case "closebrowser":
			            driver.quit();
			            return "PASS";
			        default:
			            throw new Exception("Unknown keyword " + action);
			            
		        }
		}


		public static WebElement findMyElement(String locatorType, String locator) throws Exception {
		        switch(locatorType.toLowerCase()) {
		        case "id":
		            return driver.findElement(By.id(locator));
		        case "name":
		            return driver.findElement(By.name(locator));
		        case "xpath":
		            return driver.findElement(By.xpath(locator));
		        case "cssselector":
		            return driver.findElement(By.cssSelector(locator));
		        case "linktext":
		            return driver.findElement(By.linkText(locator));
		        default:
		            throw new Exception("Unknown selector type " + locator);
		        }
		}
		
		public static void logErrorValidate(int rowCount,int lastrownumber,String actualText, String dataInput, String tipeField) throws Exception
		{
			XSSFRow row = sheet.getRow(rowCount);
			//Substring description error if character more than this specified number
			if(actualText.length() > 161)
			{
				actualText = actualText.substring(0,160) + "...";
			}
			//
			if(tipeField.equals("-"))
			{
				row.createCell(9).setCellValue(actualText);
			}else
			{
	            row.createCell(9).setCellValue("Actual text: " + actualText + ", expected text: " + dataInput);
			}
			
			//cs.captureScreenShot(rowCount);
            nameFile = "Gambar "+rowCount;
            //iad.createResult(nameFile,rowCount,lastrownumber,"FAIL");
		}

}
