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

public class mostDeclined {
	public static final String MostDeclinedFile = "NasdaqAfterHoursMostDeclined.xlsx";

	public static void main(String[] args) {
		// TODO Auto-generated method stub
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


		int count = driver.findElements(By.xpath("//div[@id='_declined']/table/tbody/tr")).size();
		System.out.println(count);


		ArrayList<String> SymbolMostDeclined=new ArrayList<String>();
		ArrayList<String> CompanyMostDeclined=new ArrayList<String>();
		ArrayList<String> LastSaleMostDeclined=new ArrayList<String>();
		ArrayList<String> ChangeNetMostDeclined=new ArrayList<String>();
		ArrayList<String> ShareVolumeMostDeclined=new ArrayList<String>();
		
		SymbolMostDeclined.add("Symbol");
		CompanyMostDeclined.add("Company");
		LastSaleMostDeclined.add(" Last Sale ");
		ChangeNetMostDeclined.add(" Net Change ");
		ShareVolumeMostDeclined.add("Share Volume");


		for(int i = 1; i <= count; i++) {
			String Xpath = "//div[@id='_declined']/table/tbody/tr[" + String.valueOf(i) + "]/td[1]/h3/a";
			//System.out.println(Xpath);
			SymbolMostDeclined.add(driver.findElement(By.xpath(Xpath)).getText());


			Xpath = "//div[@id='_declined']/table/tbody/tr[" + String.valueOf(i) + "]/td[2]/b";
			CompanyMostDeclined.add(driver.findElement(By.xpath(Xpath)).getText());


			Xpath = "//div[@id='_declined']/table/tbody/tr[" + String.valueOf(i) + "]/td[4]";
			LastSaleMostDeclined.add(driver.findElement(By.xpath(Xpath)).getText());


			Xpath = "//div[@id='_declined']/table/tbody/tr[" + String.valueOf(i) + "]/td[5]";

			String test = driver.findElement(By.xpath(Xpath)).getText();

			if(test.contains("unch")) {
				ChangeNetMostDeclined.add(test);
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

				ChangeNetMostDeclined.add(value);
			}			

			Xpath = "//div[@id='_declined']/table/tbody/tr[" + String.valueOf(i) + "]/td[6]";
			ShareVolumeMostDeclined.add(driver.findElement(By.xpath(Xpath)).getText());
			//SymbolMostActive.add(driver.findElement(By.xpath("//div[@id='_active']/table/tbody/tr[i]/td[1]/h3/a")).getText());
		}



//		for(String obj:SymbolMostDeclined)  {
//			System.out.println(obj);  
//		}
//
//		for(String obj:CompanyMostDeclined)  {
//			System.out.println(obj);  
//		}
//		for(String obj:LastSaleMostDeclined)  {
//			System.out.println(obj);  
//		}
//		for(String obj:ChangeNetMostDeclined)  {
//			System.out.println(obj);  
//		}
//		for(String obj:ShareVolumeMostDeclined)  {
//			System.out.println(obj);  
//		}

		//System.out.println(count);

		driver.close();
		
		XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Most-Declined");
        
        for(int i = 0; i <= count; i++) {
        	Row row = sheet.createRow(i);
        	int col = 0;
        	Cell cell = row.createCell(col++);
        	cell.setCellValue((String) SymbolMostDeclined.get(i));
        	
        	cell = row.createCell(col++);
        	cell.setCellValue((String) CompanyMostDeclined.get(i));
        	
        	 cell = row.createCell(col++);
        	cell.setCellValue((String) LastSaleMostDeclined.get(i));
        	
        	cell = row.createCell(col++);
        	cell.setCellValue((String) ChangeNetMostDeclined.get(i));
        	
        	cell = row.createCell(col++);
        	cell.setCellValue((String) ShareVolumeMostDeclined.get(i));
        	
        }
        
        try {
            FileOutputStream outputStream = new FileOutputStream(MostDeclinedFile);
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
