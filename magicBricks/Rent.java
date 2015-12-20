

//16-12-2015
//not working properly for new window
// search
// open new window
//switch to new window and back to main window



import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import java.lang.String;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import java.io.File;
import java.io.IOException;
import java.sql.Date;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.DateTime;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import org.openqa.selenium.JavascriptExecutor;

public class Rent {

	
	static WebDriver driver = new FirefoxDriver();
	
	static File input = new File ("C:\\Users\\SID\\Desktop\\Oyeok\\magicbricks\\rent\\input.xls");
	static int x;
	static Label data;

	public static void main(String[] args) throws IOException, RowsExceededException, WriteException, InterruptedException, BiffException {
		
		Workbook writableBook = Workbook.getWorkbook(input);
		Sheet sheet1 =writableBook.getSheet(0);
	
	for(x=0;x<17;x++){	
	
	
		 Cell exl = sheet1.getCell(1, x);
	     String sexl = exl.getContents();
        Cell city = sheet1.getCell(0, x);
        String sCITY = city.getContents();
        System.out.println(" \n"+sCITY);
		
        if(x==0){
    		search(sCITY);
    	
    		newWindow(sexl);
            }
            else{
            	search1(sCITY);
         
            		newWindow(sexl);
            }
		
		
	}

	}
	
	static void search(String sCITY) throws InterruptedException{
	
		
		//  url address of site
			driver.get("http://www.magicbricks.com/");
			driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
			
			
			// city search 
			
			driver.findElement(By.xpath("//*[@id='rentTab']")).click(); // click on rent tab
			
			//flat
			driver.findElement(By.xpath("//*[@id='rent_propertyType']")).click();
			driver.findElement(By.xpath("//*[@id='propType_rent_chk_10002_10003_10021_10022_10020']")).click();
			Thread.sleep(500);
			
			// budget
			driver.findElement(By.xpath("//*[@id='rent_budget_lbl']")).click();
			Thread.sleep(500);
			driver.findElement(By.xpath("//*[@id='budgetRentRange']/div[2]/ul/li[9]")).click();
			Thread.sleep(500);
			driver.findElement(By.xpath("//*[@id='budgetRentRange']/div[3]/ul/li[7]")).click();
			
			// city
			driver.findElement(By.xpath("//*[@id='keyword']")).click();
		//	driver.findElement(By.xpath("//*[@id='keyword']")).sendKeys("Andheri West");
			driver.findElement(By.xpath("//*[@id='keyword']")).sendKeys(sCITY);
			
			//seach click
			driver.findElement(By.xpath("//*[@id='btnPropertySearch']")).click();
			driver.findElement(By.xpath("//*[@id='btnPropertySearch']")).click();
			Thread.sleep(2000);
			driver.findElement(By.xpath("//*[@id='inputinputListings']")).click();
			driver.findElement(By.xpath("//*[@id='inputListings_I']")).click();
			driver.findElement(By.xpath("//*[@id='srpWrapper']")).click();

			
			Thread.sleep(4000);
		}
	static void search1(String sCITY) throws InterruptedException{
	
		
		//  url address of site
			driver.get("http://www.magicbricks.com/");
			driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
			
			
			// city search 
			
			driver.findElement(By.xpath("//*[@id='rentTab']")).click(); // click on rent tab
	/*		
			//flat
		//	driver.findElement(By.xpath("//*[@id='rent_propertyType']")).click();
			driver.findElement(By.xpath("//*[@id='propType_rent_chk_10002_10003_10021_10022_10020']")).click();
			Thread.sleep(500);
			
			// budget
			driver.findElement(By.xpath("//*[@id='rent_budget_lbl']")).click();
			Thread.sleep(500);
			driver.findElement(By.xpath("//*[@id='budgetRentRange']/div[2]/ul/li[9]")).click();
			Thread.sleep(500);
			driver.findElement(By.xpath("//*[@id='budgetRentRange']/div[3]/ul/li[7]")).click();
			
		*/	// city
			driver.findElement(By.xpath("//*[@id='keyword']")).click();
			driver.findElement(By.xpath("//*[@id='keyword']")).clear();
			driver.findElement(By.xpath("//*[@id='autoSuggestInputDivkeyword']/div/div[2]")).click();
		//	driver.findElement(By.xpath("//*[@id='keyword']")).sendKeys("Andheri West");
			driver.findElement(By.xpath("//*[@id='keyword']")).sendKeys(sCITY);
			
			//seach click
			driver.findElement(By.xpath("//*[@id='btnPropertySearch']")).click();
			driver.findElement(By.xpath("//*[@id='btnPropertySearch']")).click();
			Thread.sleep(2000);
	
			driver.findElement(By.xpath("//*[@id='inputinputListings']")).click();
			driver.findElement(By.xpath("//*[@id='inputListings_I']")).click();
			driver.findElement(By.xpath("//*[@id='srpWrapper']")).click();

			
			Thread.sleep(4000);
			
		}

	static void newWindow(String sexl) throws InterruptedException, IOException, WriteException{
		System.out.println("");
		File fExcel = new File (sexl);
		WritableWorkbook writableBook1 = Workbook.createWorkbook(fExcel);
		writableBook1.createSheet("Data",0);
		WritableSheet writableSheet = writableBook1.getSheet(0);
	      
		Thread.sleep(500);
		List<WebElement> location = new ArrayList<WebElement>(); 	
		location.addAll(driver.findElements(By.xpath("//*[starts-with(@id,'localityName')]")));
		System.out.println(location.size());
		
		List<WebElement> price = new ArrayList<WebElement>(); 	
		price.addAll(driver.findElements(By.xpath("//*[starts-with(@id,'pricePropertyVal')]")));
		System.out.println(price.size());
		
		List<WebElement> post= new ArrayList<WebElement>(); 	
		post.addAll(driver.findElements(By.xpath("//*[starts-with(@id,'postedSince')]")));
		System.out.println(post.size());
		
		
		
		for (int i=1;i<location.size();i++){
			 System.out.println(i);
			// Dealer Price
			try{
				data = new Label(0,i,price.get(i).getText());
				writableSheet.addCell(data);  
			}catch(Exception e){  }
		
			// Room details
			try{
				data = new Label(1,i,post.get(i).getText());
				writableSheet.addCell(data);  
			}catch(Exception e){  }
			
			// Room details
					try{
						data = new Label(2,i,location.get(i).getText());
						writableSheet.addCell(data);  
					}catch(Exception e){  }
		
		String strMainWindow = driver.getWindowHandle();
	//	System.out.println("Window title : \n" + driver.getTitle()); 
	
		try{
		location.get(i).click();
		}catch(Exception e){ }
		Thread.sleep(3000);
		
	//	Thread.sleep(2000);
		Set<String> strHandles = driver.getWindowHandles(); // store all open Window 
		  
		int k=1;
		
		   for(String handle:strHandles){  // Number of Window loop
			   driver.switchTo().window(handle);  // Switch to new Window
			   String strTitle = driver.getTitle();
			   System.out.print(k);
			  
			   if(k>=2){
				// Configuration
					try{
						data = new Label(3,i,driver.findElement(By.xpath("//*[@id='overNav']/div/div[2]/div[5]/div[4]/div[1]/div[2]")).getText());
						writableSheet.addCell(data);	  
					}catch(Exception e){  }   
				   
					// Furnising
					try{
						data = new Label(4,i,driver.findElement(By.xpath("//*[@id='overNav']/div/div[2]/div[5]/div[4]/div[2]/div[2]")).getText());
						writableSheet.addCell(data);	  
					}catch(Exception e){ } 
					
					
					// Floor Details
					try{
						data = new Label(5,i,driver.findElement(By.xpath("//*[@id='overNav']/div/div[2]/div[5]/div[4]/div[3]/div[2]")).getText());
						writableSheet.addCell(data);	  
					}catch(Exception e){  } 
					
					
					// Car Parking
					try{
						data = new Label(6,i,driver.findElement(By.xpath("//*[@id='overNav']/div/div[2]/div[5]/div[4]/div[4]/div[2]")).getText());
						writableSheet.addCell(data);	  
					}catch(Exception e){  } 
					
					
					// Available From
					try{
						data = new Label(7,i,driver.findElement(By.xpath("//*[@id='overNav']/div/div[2]/div[5]/div[4]/li/div/div[2]/span")).getText());
						writableSheet.addCell(data);	  
					}catch(Exception e){  } 
					
					
					// Rent
					try{
						data = new Label(8,i,driver.findElement(By.xpath("//*[@id='overNav']/div/div[6]/div[1]/div[2]")).getText());
						writableSheet.addCell(data);	  
					}catch(Exception e){ } 
					
					
					// Address

					try{
						data = new Label(9,i,driver.findElement(By.xpath("//*[@id='overNav']/div/div[6]/div[2]/div[2]")).getText());
						writableSheet.addCell(data);	  
					}catch(Exception e){ } 
					
					
					// Project & Society
					try{
				//		System.out.println(driver.findElement(By.xpath("//*[@id='results']/div[1]/div["+i+"]/div[2]/div[2]/div[4]")).getText());
						data = new Label(10,i,driver.findElement(By.xpath("//*[@id='overNav']/div/div[6]/div[2]/div[2]")).getText());
						writableSheet.addCell(data);	  
					}catch(Exception e){  } 
					
					
					// Facing
					try{
						data = new Label(11,i,driver.findElement(By.xpath("//*[@id='overNav']/div/div[6]/div[4]/div[2]")).getText());
						writableSheet.addCell(data);	  
					}catch(Exception e){  } 
					
					
					// Overlooking

					try{
						data = new Label(12,i,driver.findElement(By.xpath("//*[@id='overNav']/div/div[6]/div[5]/div[2]")).getText());
						writableSheet.addCell(data);	  
					}catch(Exception e){  } 
					
					
					// Area

					try{
				//		System.out.println(driver.findElement(By.xpath("//*[@id='results']/div[1]/div["+i+"]/div[2]/div[2]/div[4]")).getText());
						data = new Label(13,i,driver.findElement(By.xpath("//*[@id='overNav']/div/div[6]/div[6]/div[2]/div/div")).getText());
						writableSheet.addCell(data);	  
					}catch(Exception e){  } 
					
					
					// Age of Construction

					try{
				//		System.out.println(driver.findElement(By.xpath("//*[@id='results']/div[1]/div["+i+"]/div[2]/div[2]/div[4]")).getText());
						data = new Label(14,i,driver.findElement(By.xpath("//*[@id='overNav']/div/div[6]/div[11]/div[2]")).getText());
						writableSheet.addCell(data);	  
					}catch(Exception e){ } 
					
					// About

					try{
				//		System.out.println(driver.findElement(By.xpath("//*[@id='results']/div[1]/div["+i+"]/div[2]/div[2]/div[4]")).getText());
						data = new Label(14,i,driver.findElement(By.className("dDetail")).getText());
						writableSheet.addCell(data);	  
					}catch(Exception e){ } 
					
					try{
						//		System.out.println(driver.findElement(By.xpath("//*[@id='results']/div[1]/div["+i+"]/div[2]/div[2]/div[4]")).getText());
								data = new Label(14,i,driver.getCurrentUrl());
								writableSheet.addCell(data);	  
							}catch(Exception e){ } 
							
				  
					 driver.close(); // Close new Window
			   }
					   else{ k++;}
		   }
		  
		  // driver.close(); // Close new Window
		   driver.switchTo().window(strMainWindow); // Switch back to main window
		
	} // end for new window
		 writableBook1.write();
			writableBook1.close();
}
	
	
	
}