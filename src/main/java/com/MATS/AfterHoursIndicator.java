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

public class AfterHoursIndicator {
	
	public static final String AfterHoursIndicatorFile = "Nasdaq-AfterHoursIndicator.xlsx";

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		 //div[@id='left-column-div']/div[3]/table/tbody/tr
		
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\Samsung\\Documents\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(25, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(25,TimeUnit.SECONDS);


		try {
			driver.get("https://www.nasdaq.com/extended-trading/afterhours-mostactive.aspx");
		}
		catch(Exception x) {

		}


		Actions A = new Actions(driver);
		A.moveToElement(driver.findElement(By.xpath("//a[@id='tab3norm']/span"))).click().build().perform();


		int count = driver.findElements(By.xpath("//div[@id='left-column-div']/div[3]/table/tbody/tr")).size();
		String Xpath;
		
		ArrayList<String> SymboAHI=new ArrayList<String>();
		ArrayList<String> NameAHI=new ArrayList<String>();
		ArrayList<String> LastSaleAHI=new ArrayList<String>();
		ArrayList<String> AfterHoursLastSaleAHI=new ArrayList<String>();
		ArrayList<String> AfterHoursPercentChangeAHI=new ArrayList<String>();
		ArrayList<String> AfterHoursTimeAHI=new ArrayList<String>();
		
		for(int i = 1; i <= 6; i++) {
			Xpath = "//div[@id='left-column-div']/div[3]/table/thead/tr/th["+ i + "]/a";
			//System.out.println(Xpath);
			//System.out.println(driver.findElement(By.xpath(Xpath)).getText());
			switch(i) {
			case 1:
				NameAHI.add(driver.findElement(By.xpath(Xpath)).getText());
				break;
			case 2:
				SymboAHI.add(driver.findElement(By.xpath(Xpath)).getText());
				break;
			case 3:
				LastSaleAHI.add(driver.findElement(By.xpath(Xpath)).getText());
				break;
			case 4:
				AfterHoursLastSaleAHI.add(driver.findElement(By.xpath(Xpath)).getText());
				break;
			case 5:
				AfterHoursPercentChangeAHI.add(driver.findElement(By.xpath(Xpath)).getText());
				break;
			default:
				AfterHoursTimeAHI.add(driver.findElement(By.xpath(Xpath)).getText());
			}
		}
		
		
		
		for(int i = 1; i <= count; i++) {
			Xpath = "//div[@id='left-column-div']/div[3]/table/tbody/tr[" + i + "]/td[1]";
			//System.out.println(Xpath);
			NameAHI.add(driver.findElement(By.xpath(Xpath)).getText());


			Xpath = "//div[@id='left-column-div']/div[3]/table/tbody/tr[" + i + "]/td[2]/span/h3/a";
			SymboAHI.add(driver.findElement(By.xpath(Xpath)).getText());


			Xpath = "//div[@id='left-column-div']/div[3]/table/tbody/tr[" + i + "]/td[3]";
			LastSaleAHI.add(driver.findElement(By.xpath(Xpath)).getText());

			
			Xpath = "//div[@id='left-column-div']/div[3]/table/tbody/tr[" + i + "]/td[4]";
			AfterHoursLastSaleAHI.add(driver.findElement(By.xpath(Xpath)).getText());
			
			Xpath = "//div[@id='left-column-div']/div[3]/table/tbody/tr[" + i + "]/td[5]/span";

				String test = driver.findElement(By.xpath(Xpath)).getText();

			
				String className = driver.findElement(By.xpath(Xpath)).getAttribute("class");
				//		System.out.println("testing     >>>>>>>>  " + className );
				String value = driver.findElement(By.xpath(Xpath)).getText();

			

				if(className.contains("green")) {
					value += " Up ";
				}
				else if(className.contains("red")){
					value += " Down ";
				}
				else {
					value += " Unch";
				}

				AfterHoursPercentChangeAHI.add(value);
						

				Xpath = "//div[@id='left-column-div']/div[3]/table/tbody/tr[" + i + "]/td[6]";
				AfterHoursTimeAHI.add(driver.findElement(By.xpath(Xpath)).getText());
			//SymbolMostActive.add(driver.findElement(By.xpath("//div[@id='_active']/table/tbody/tr[i]/td[1]/h3/a")).getText());
		}

		
		System.out.println(count);
		
//		for(String obj:	NameAHI)  {
//			System.out.println(obj);  
//		}
//		
		driver.close();
		
		
		XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("AfterHoursIndicator");
        
        
        for(int i = 0; i <= count; i++) {
        	Row row = sheet.createRow(i);
        	int col = 0;
        	Cell cell = row.createCell(col++);
        	cell.setCellValue((String) NameAHI.get(i));
        	
        	 cell = row.createCell(col++);
        	cell.setCellValue((String) SymboAHI.get(i));
        	
        	cell = row.createCell(col++);
        	cell.setCellValue((String)LastSaleAHI.get(i));
        	
        	 cell = row.createCell(col++);
        	cell.setCellValue((String) AfterHoursLastSaleAHI.get(i));
        	
        	cell = row.createCell(col++);
        	cell.setCellValue((String) AfterHoursPercentChangeAHI.get(i));
        	
        	cell = row.createCell(col++);
        	cell.setCellValue((String) AfterHoursTimeAHI.get(i));
        	
        }
        
        try {
            FileOutputStream outputStream = new FileOutputStream(AfterHoursIndicatorFile);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

		
	}

}
