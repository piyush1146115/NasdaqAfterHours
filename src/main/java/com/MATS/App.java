package com.MATS;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

/**
 * Hello world!
 *
 */
public class App 
{
	public static final String MostActiveFile = "Nasdaq-After-Hours-MostActive.xlsx";
    public static void main( String[] args )
    {
      //  System.out.println( "Hello World!" );
    	System.setProperty("webdriver.chrome.driver", "C:\\Users\\Samsung\\Documents\\chromedriver.exe");
		
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(20,TimeUnit.SECONDS);
		
		try {
		driver.get("https://www.nasdaq.com/extended-trading/afterhours-mostactive.aspx");
		}
		catch(Exception x) {
			
		}
		
		Actions A = new Actions(driver);
		A.moveToElement(driver.findElement(By.xpath("//a[@id='show-all']/span"))).click().build().perform();
	
		
		int count;
		
		count = driver.findElements(By.xpath("//div[@id='_active']/table/tbody/tr")).size();
//		
//		
		 ArrayList<String> SymbolMostActive=new ArrayList<String>();
		 ArrayList<String> CompanyMostActive=new ArrayList<String>();
		 ArrayList<String> LastSaleMostActive=new ArrayList<String>();
		 ArrayList<String> ChangeNetMostActive=new ArrayList<String>();
		 ArrayList<String> ShareVolumeMostActive=new ArrayList<String>();
		 
		 
			SymbolMostActive.add("Symbol");
			CompanyMostActive.add("Company");
			LastSaleMostActive.add(" Last Sale ");
			ChangeNetMostActive.add(" Net Change ");
			ShareVolumeMostActive.add("Share Volume");
		 
		for(int i = 1; i <= count; i++) {
			String Xpath = "//div[@id='_active']/table/tbody/tr[" + String.valueOf(i) + "]/td[1]/h3/a";
			//System.out.println(Xpath);
			SymbolMostActive.add(driver.findElement(By.xpath(Xpath)).getText());
			
			
			Xpath = "//div[@id='_active']/table/tbody/tr[" + String.valueOf(i) + "]/td[2]/b";
			CompanyMostActive.add(driver.findElement(By.xpath(Xpath)).getText());
			
			
			Xpath = "//div[@id='_active']/table/tbody/tr[" + String.valueOf(i) + "]/td[4]";
			LastSaleMostActive.add(driver.findElement(By.xpath(Xpath)).getText());
			
			
			Xpath = "//div[@id='_active']/table/tbody/tr[" + String.valueOf(i) + "]/td[5]";
			
			String test = driver.findElement(By.xpath(Xpath)).getText();
			
			if(test.contains("unch")) {
				ChangeNetMostActive.add(test);
			}
			else {
				Xpath += "/span";
				String className = driver.findElement(By.xpath(Xpath)).getAttribute("class");
				//System.out.println("testing     >>>>>>>>  " + className );
				String value = driver.findElement(By.xpath(Xpath)).getText();

			
				value.replace("[?]", "");
				
				if(className.contains("green")) {
					value += " Up ";
				}
				else {
					value += " Down ";
				}
				
			ChangeNetMostActive.add(value);
			}
//			System.out.println("test " + i + "  ---->>  "+  test);
//			if(test.contains("unch")) {
//				System.out.println("  entered");
//			}
			
			
			Xpath = "//div[@id='_active']/table/tbody/tr[" + String.valueOf(i) + "]/td[6]";
			ShareVolumeMostActive.add(driver.findElement(By.xpath(Xpath)).getText());
			//SymbolMostActive.add(driver.findElement(By.xpath("//div[@id='_active']/table/tbody/tr[i]/td[1]/h3/a")).getText());
		}
		

		 driver.close();
		 
		 
			XSSFWorkbook workbook = new XSSFWorkbook();
	        XSSFSheet sheet = workbook.createSheet("Most-Advanced");
	        
	        for(int i = 0; i <= count; i++) {
	        	Row row = sheet.createRow(i);
	        	int col = 0;
	        	Cell cell = row.createCell(col++);
	        	cell.setCellValue((String) SymbolMostActive.get(i));
	        	
	        	cell = row.createCell(col++);
	        	cell.setCellValue((String) CompanyMostActive.get(i));
	        	
	        	 cell = row.createCell(col++);
	        	cell.setCellValue((String) LastSaleMostActive.get(i));
	        	
	        	cell = row.createCell(col++);
	        	cell.setCellValue((String) ChangeNetMostActive.get(i));
	        	
	        	cell = row.createCell(col++);
	        	cell.setCellValue((String) ShareVolumeMostActive.get(i));
	        	
	        }
	        
	        try {
	            FileOutputStream outputStream = new FileOutputStream(MostActiveFile);
	            workbook.write(outputStream);
	            workbook.close();
	        } catch (FileNotFoundException e) {
	            e.printStackTrace();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }

	        System.out.println("Done");
	        
		
    }
}
