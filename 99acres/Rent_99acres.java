

//15-12-15
//link scrapping 


import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;


import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Rent_99acres {

	public static void main(String[] args) throws InterruptedException, RowsExceededException, WriteException, IOException, BiffException {
		WebDriver driver = new FirefoxDriver();
		File input = new File ("C:\\Users\\SID\\Desktop\\99\\input.xls");
		Workbook writableBook = Workbook.getWorkbook(input);
		Sheet sheet1 =writableBook.getSheet(0);
	for(int z=0;z<18;z++){	
		
		Cell link = sheet1.getCell(2, z);
        Cell exl = sheet1.getCell(1, z);
        Cell city = sheet1.getCell(0, z);
        String slink = link.getContents();
        String sexl = exl.getContents();
        String sCITY = exl.getContents();
        System.out.println("sCITY \n");
		// Open 99acres page
		driver.get(slink);
		driver.manage().timeouts().implicitlyWait(4, TimeUnit.SECONDS); // wait for 10 sec
		 
		
 		
	
		List<WebElement> phoneOn = new ArrayList<WebElement>(); 	
		phoneOn.addAll(driver.findElements(By.xpath("//*[starts-with(@id,'viewphno')]")));
		System.out.println(phoneOn.size());
		
		int y=0;
		boolean br=true;
		
//		Create Excel sheet 
				File fExcel = new File (sexl);
				WritableWorkbook writableBook1 = Workbook.createWorkbook(fExcel);
				writableBook1.createSheet("Data",0);
				WritableSheet writableSheet = writableBook1.getSheet(0);
				Label data;
				
				
				
				
	//	do{
			Thread.sleep(4000);
			
		List<WebElement> phoneOn1 = new ArrayList<WebElement>(); 	
		phoneOn1.addAll(driver.findElements(By.xpath("//*[starts-with(@id,'viewphno')]")));
		System.out.println(phoneOn1.size());

		
		List<WebElement> rs = new ArrayList<WebElement>(); 	
		rs.addAll(driver.findElements(By.xpath("//*[starts-with(@id,'rs')]")));
		System.out.println(rs.size());
		
		List<WebElement> desc = new ArrayList<WebElement>(); 
		desc.addAll(driver.findElements(By.xpath("//*[starts-with(@id,'desc')]")));
		System.out.println(desc.size());
		
		
		
	//	String start = "//*[@id='results']/div[";
	//	String mid = "]/div[";
	//	String end = "]";
		int j=0;
		
		for(int i =1;i<phoneOn.size();i++){
					
			
			// Dealer Price
			try{
			//	System.out.println(rs.get(j).getText());
				data = new Label(1,y,rs.get(j).getText());
				writableSheet.addCell(data);  
			}catch(Exception e){  }
		
			// Room details
			try{
			//	System.out.println(desc.get(j).getText());
				data = new Label(5,y,desc.get(j).getText());
				writableSheet.addCell(data);  
			}catch(Exception e){  }
		
			
			// Built-up Area 
			try{
		//		System.out.println(driver.findElement(By.xpath("//*[@id='results']/div[1]/div["+i+"]/div[2]/div[2]/span[1]")).getText());
				data = new Label(3,y,driver.findElement(By.xpath("//*[@id='results']/div[1]/div["+i+"]/div[2]/div[2]/span[1]")).getText());
				writableSheet.addCell(data);  
			}catch(Exception e){  }
			
			
			// Society :
			try{
			//	System.out.println(driver.findElement(By.xpath("//*[@id='results']/div[1]/div["+i+"]/div[2]/div[2]/span[2]")).getText());
				data = new Label(4,y,driver.findElement(By.xpath("//*[@id='results']/div[1]/div["+i+"]/div[2]/div[2]/span[2]")).getText());
				writableSheet.addCell(data);	  
			}catch(Exception e){  }
			
			
			// Dealer Deatails
		//	System.out.print("Dealer : ");
		// error when deler = owner
		//	System.out.println(driver.findElement(By.xpath(".//*[@id='results']/div[1]/div["+i+"]/div[2]/div[4]/a")).getText());                  
			
			// Dealer Deatails and Posted on
			try{
		//		System.out.println(driver.findElement(By.xpath("//*[@id='results']/div[1]/div["+i+"]/div[2]/div[4]")).getText());
				data = new Label(0,y,driver.findElement(By.xpath(".//*[@id='results']/div[1]/div["+i+"]/div[2]/div[4]")).getText());
				writableSheet.addCell(data); 
			}catch(Exception e){  }
			
			
			// Description :
			try{
		//		System.out.println(driver.findElement(By.xpath("//*[@id='results']/div[1]/div["+i+"]/div[2]/div[2]/div[4]")).getText());
				data = new Label(12,y,driver.findElement(By.xpath("//*[@id='results']/div[1]/div["+i+"]/div[2]/div[2]/div[4]")).getText());
				writableSheet.addCell(data);	  
			}catch(Exception e){  }
			
						
			System.out.print("***************** "+i+"\n");
			j++;
			y++;
			
			// new window
			
			String start = "//*[@id='results']/div[1]/div[";
			
			String strMainWindow = driver.getWindowHandle();
	//		System.out.println("Window title" + driver.getTitle()); 
			driver.findElement(By.xpath(start+i+"]")).click();
				
				
				Thread.sleep(2000);
				Set<String> strHandles = driver.getWindowHandles(); // store all open Window 
				   
				
				   for(String handle:strHandles){  // Number of Window loop
					   driver.switchTo().window(handle);  // Switch to new Window
					   String strTitle = driver.getTitle();
			//		   System.out.println(driver.getTitle());
					   
					   
					// Deposit :
						try{
		//					System.out.println(driver.findElement(By.xpath("//*[@id='PdInfoStart']/div[3]/div[1]/div[4]/div/div[1]/ul/li/em")).getText());
							data = new Label(2,y,driver.findElement(By.xpath("//*[@id='PdInfoStart']/div[3]/div[1]/div[4]/div/div[1]/ul/li/em")).getText());
							writableSheet.addCell(data);	  
						}catch(Exception e){  }
					   
					// Available from :
						try{
			//				System.out.println(driver.findElement(By.xpath("//*[@id='PdInfoStart']/div[2]/div[1]/div[3]/div/div/i[2]")).getText());
							data = new Label(6,y,driver.findElement(By.xpath("//*[@id='PdInfoStart']/div[2]/div[1]/div[3]/div/div/i[2]")).getText());
							writableSheet.addCell(data);	  
						}catch(Exception e){  }
						
					// Available for :
						try{
			//				System.out.println(driver.findElement(By.xpath("//*[@id='PdInfoStart']/div[2]/div[1]/div[3]/div/div/i[8]")).getText());
							data = new Label(7,y,driver.findElement(By.xpath("//*[@id='PdInfoStart']/div[2]/div[1]/div[3]/div/div/i[8]")).getText());
							writableSheet.addCell(data);	  
						}catch(Exception e){  }
						
					// Property age :
						try{
			//				System.out.println(driver.findElement(By.xpath("//*[@id='PdInfoStart']/div[2]/div[1]/div[3]/div/div/i[4]")).getText());
							data = new Label(8,y,driver.findElement(By.xpath("//*[@id='PdInfoStart']/div[2]/div[1]/div[3]/div/div/i[4]")).getText());
							writableSheet.addCell(data);	  
						}catch(Exception e){  }
						
					//  Floor :
						try{
			//				System.out.println(driver.findElement(By.xpath("//*[@id='total_floorLabel']")).getText());
							data = new Label(9,y,driver.findElement(By.xpath("//*[@id='total_floorLabel']")).getText());
							writableSheet.addCell(data);	  
						}catch(Exception e){  }
						
					//  Transaction Type :
						try{
			//				System.out.println(driver.findElement(By.xpath("//*[@id='PdInfoStart']/div[2]/div[1]/div[3]/div/div/i[6]")).getText());
							data = new Label(10,y,driver.findElement(By.xpath("//*[@id='PdInfoStart']/div[2]/div[1]/div[3]/div/div/i[6]")).getText());
							writableSheet.addCell(data);	  
						}catch(Exception e){  }
					   
						//Additional Details
						try{
				//			System.out.println(driver.findElement(By.xpath("//*[@id='PdInfoStart']/div[3]/div[1]/div[2]/p")).getText());
							data = new Label(13,y,driver.findElement(By.xpath("//*[@id='PdInfoStart']/div[3]/div[1]/div[2]/p")).getText());
							writableSheet.addCell(data);	  
						}catch(Exception e){  }
					   String url=driver.getCurrentUrl();
						data = new Label(14,y,url);
						writableSheet.addCell(data);
					   
				
				   }
				 driver.close(); // Close new Window
				   driver.switchTo().window(strMainWindow); // Switch back to main window
			
			
			 
		}
		
		writableBook1.write();
		writableBook1.close();
		
		
}
	driver.close();
}
}


		

