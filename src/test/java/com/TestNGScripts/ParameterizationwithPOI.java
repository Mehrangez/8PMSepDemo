package com.TestNGScripts;

import org.testng.annotations.Test;

import com.utility.XLs_dataProvider;

import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class ParameterizationwithPOI {



WebDriver driver;


// pass the name of data provide method as parameter to this test method 

@BeforeClass 



        public void startbrowser()
        {
	
	System.setProperty("webdriver.chrome.driver", "C:\\Users\\mmeli\\Downloads\\chromedriver_win32 (2)\\chromedriver.exe");
 	driver = new ChromeDriver();
 	driver.manage().window().maximize();
 	driver.get("https://en.wikipedia.org/w/index.php?title=Special:CreateAccount&returnto=Selenium+%28software%29");
 	
        }

  @Test(dataProvider="testdata")
  
	public  void wikipagedata(String name, String password, String retype, String email)
	{
		driver.findElement(By.id("wpName2")).clear();
		
		driver.findElement(By.id("wpName2")).sendKeys(name);
		
		
		driver.findElement(By.id("wpPassword2")).clear();
		driver.findElement(By.id("wpPassword2")).sendKeys(password);
		driver.findElement(By.id("wpRetype")).clear();
		driver.findElement(By.id("wpRetype")).sendKeys(retype);
		driver.findElement(By.id("wpEmail")).clear();
		driver.findElement(By.id("wpEmail")).sendKeys(email);
	
	}
  
  
  
  @DataProvider(name="testdata")
  
  /// We can automate it by saving in static method in utility and keep calling it 
  
  public Object[][] readExcel12()
  {
	 Object input[][]= XLs_dataProvider.getTestData("Sheet1");
 
  
  return input;
  }
  
 
	// APAche POI 
  @DataProvider(name="testdata")
  
  
         public Object[][] readExcel() throws EncryptedDocumentException, IOException
	
	{
		// 1. provide the location of the file
		
		FileInputStream f=new FileInputStream("C:\\Users\\mmeli\\Downloads\\input1.xlsx");

		// fetch the workboot from the above location	
	
	   Workbook  book =WorkbookFactory.create(f);
	
	
	// 2. fetch the sheet which has data 
	   
	   Sheet s   = book.getSheet("Sheet1");
	
	
	//3. fetch the rows and colums 
	   
	   
	  int rows= s.getLastRowNum();
	  
	  /// 4. fetch the colum 
	  
	 int col  = s.getRow(0).getLastCellNum();
	 
	 String input[][]= new String [rows][col];
	 
	 for(int i=1; i<rows; i++)
		 
	 { 
		 for(int j=0; j<col; j++)
		 {
			input [i][j] = s.getRow(i+1).getCell(j).toString(); // toString() will fetch data from excel 
		 }
	
	 }
	 
	 return input;
	
	
	
	
	
	
	
	
	
	
	
	
	}
	
	
	

	































}
