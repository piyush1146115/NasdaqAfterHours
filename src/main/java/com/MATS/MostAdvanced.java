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

public class MostAdvanced {
	public static final String MostAdvancedFile = "Nasdaq-After-Hours-MostAdvanced.xlsx";
	
	public static void main(String[] args) {
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

		int  count = driver.findElements(By.xpath("//div[@id='_advanced']/table/tbody/tr")).size();
//		System.out.println(count);


		ArrayList<String> SymbolMostAdvanced=new ArrayList<String>();
		ArrayList<String> CompanyMostAdvanced=new ArrayList<String>();
		ArrayList<String> LastSaleMostAdvanced=new ArrayList<String>();
		ArrayList<String> ChangeNetMostAdvanced=new ArrayList<String>();
		ArrayList<String> ShareVolumeMostAdvanced=new ArrayList<String>();

		SymbolMostAdvanced.add("Symbol");
		CompanyMostAdvanced.add("Company");
		LastSaleMostAdvanced.add(" Last Sale ");
		ChangeNetMostAdvanced.add(" Net Change ");
		ShareVolumeMostAdvanced.add("Share Volume");
		
		
		for(int i = 1; i <= count; i++) {
			String Xpath = "//div[@id='_advanced']/table/tbody/tr[" + String.valueOf(i) + "]/td[1]/h3/a";
			//System.out.println(Xpath);
			SymbolMostAdvanced.add(driver.findElement(By.xpath(Xpath)).getText());


			Xpath = "//div[@id='_advanced']/table/tbody/tr[" + String.valueOf(i) + "]/td[2]/b";
			CompanyMostAdvanced.add(driver.findElement(By.xpath(Xpath)).getText());


			Xpath = "//div[@id='_advanced']/table/tbody/tr[" + String.valueOf(i) + "]/td[4]";
			LastSaleMostAdvanced.add(driver.findElement(By.xpath(Xpath)).getText());


			Xpath = "//div[@id='_advanced']/table/tbody/tr[" + String.valueOf(i) + "]/td[5]";

			String test = driver.findElement(By.xpath(Xpath)).getText();

			if(test.contains("unch")) {
				ChangeNetMostAdvanced.add(test);
			}
			else {
				Xpath += "/span";
				String className = driver.findElement(By.xpath(Xpath)).getAttribute("class");
				//		System.out.println("testing     >>>>>>>>  " + className );
				String value = driver.findElement(By.xpath(Xpath)).getText();

				value.replace("[?]", "");

				if(className.contains("green")) {
					value += " Up ";
				}
				else {
					value += " Down ";
				}

				ChangeNetMostAdvanced.add(value);
			}			

			Xpath = "//div[@id='_advanced']/table/tbody/tr[" + String.valueOf(i) + "]/td[6]";
			ShareVolumeMostAdvanced.add(driver.findElement(By.xpath(Xpath)).getText());
			//SymbolMostActive.add(driver.findElement(By.xpath("//div[@id='_active']/table/tbody/tr[i]/td[1]/h3/a")).getText());
		}



//		for(String obj:SymbolMostAdvanced)  {
//			System.out.println(obj);  
//		}
//
//		for(String obj:CompanyMostAdvanced)  {
//			System.out.println(obj);  
//		}
//		for(String obj:LastSaleMostAdvanced)  {
//			System.out.println(obj);  
//		}
//		for(String obj:ChangeNetMostAdvanced)  {
//			System.out.println(obj);  
//		}
//		for(String obj:ShareVolumeMostAdvanced)  {
//			System.out.println(obj);  
//		}

		//System.out.println(count);

		driver.close();
		
	
		XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Most-Advanced");
        
        for(int i = 0; i <= count; i++) {
        	Row row = sheet.createRow(i);
        	int col = 0;
        	Cell cell = row.createCell(col++);
        	cell.setCellValue((String) SymbolMostAdvanced.get(i));
        	
        	cell = row.createCell(col++);
        	cell.setCellValue((String) CompanyMostAdvanced.get(i));
        	
        	 cell = row.createCell(col++);
        	cell.setCellValue((String) LastSaleMostAdvanced.get(i));
        	
        	cell = row.createCell(col++);
        	cell.setCellValue((String) ChangeNetMostAdvanced.get(i));
        	
        	cell = row.createCell(col++);
        	cell.setCellValue((String) ShareVolumeMostAdvanced.get(i));
        	
        }
        
        try {
            FileOutputStream outputStream = new FileOutputStream(MostAdvancedFile);
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
