 package Day_025_ExcelUtil_Apache_POI;



//step1
import java.io.FileInputStream;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;



import org.apache.poi.ss.usermodel.CellType;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import ExcelUtil.ExcelApiTest3;


public  class TC01_Excel_Test_xls2
{

    
    
    
    
    String TestURL,UserName,Password;
   // String TestURL1,UserName1,Password1;
    WebDriver driver;

    @Test
    public void hello()throws Exception
    {
    	
    	ExcelApiTest3 eat=new ExcelApiTest3();
    	 
    TestURL=eat.getCellData("C://HTML Report//OrangeHRM6//TC01_Nationality.xls","Sheet1",1,0);
    UserName=eat.getCellData("C://HTML Report//OrangeHRM6//TC01_Nationality.xls","Sheet1",1,1);
	Password=eat.getCellData("C://HTML Report//OrangeHRM6//TC01_Nationality.xls","Sheet1",1,2);
		
	 System.out.println("TestURL : "+TestURL);
	 System.out.println("UserName : "+UserName);
	 System.out.println("Password : "+Password);
	
	
 
		System.setProperty("webdriver.chrome.driver","C:\\chromedriver_win32\\chromedriver.exe");
		 driver =new ChromeDriver();
		 driver.manage().window().maximize() ;	
		 driver.get(TestURL);
				
		 findElement(By.name("txtUsername")).sendKeys(UserName);
		 findElement(By.name("txtPassword")).sendKeys(Password);
	     findElement(By.id("btnLogin")).click();
	     
		 driver.quit();
		 
	

    }
    
 
 
  
    
      
    
    
    
    
    

	 public  WebElement findElement(By by) throws Exception 
		{
					
			 WebElement elem = driver.findElement(by);    	    
			
			 
			if (driver instanceof JavascriptExecutor) 
			{
			 ((JavascriptExecutor)driver).executeScript("arguments[0].style.border='3px solid red'", elem);
		 
			}
			
			return elem;
		}

	
    
    
    
    
    
    
    
    
    
}

