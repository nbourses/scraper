import java.io.File;
import java.io.IOException;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class esearch {
static int k=1,x=1;
	static WebDriver driver = new FirefoxDriver();
	
//	Create Excel sheet 
		static File fExcel = new File ("C:\\Users\\SID\\Desktop\\selenium test\\esearch\\esearch1.xls");
		static WritableWorkbook writableBook ;
		


	public static void main(String[] args) throws InterruptedException, IOException, RowsExceededException, WriteException {
		
		 // check next button present or not
		   String nxtPageStart = "//*[@id='RegistrationGrid']/tbody/tr[12]/td/table/tbody/tr/td[";
		   String nxtPageEnd= "]/a";
			
		
		search();
		
		boolean br = false;
		do {
		window();
		try{
		 br =driver.findElement(By.xpath(nxtPageStart+x+nxtPageEnd)).isDisplayed();
		 System.out.print(br);
		 driver.findElement(By.xpath(nxtPageStart+x+nxtPageEnd)).click();
		 Thread.sleep(10000);
		}catch (Exception e){ break;}
		}while(br==true);
		
		
		writableBook.write(); // Write on Excel Sheet
		writableBook.close(); // Close Excel Sheet
		driver.close(); // close Main Window
	}

	static void search() throws InterruptedException{
		driver.get("https://esearchigr.maharashtra.gov.in/testingesearch/wfsearch.aspx");
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		System.out.println("url");
		
	// To Select DropDown value = mumbai 	
		WebElement element = driver.findElement(By.xpath("//*[@id='ddlDistrict']"));
		Select DropDown=new Select(element);
		DropDown.selectByValue("31");
		System.out.println("city");
		Thread.sleep(1000);
	 
	//	Select subcity
		driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);//wait for 60 sec.
	    driver.findElement(By.id("txtAreaName")).sendKeys("Malad");
	    System.out.println("malad -1");
	   Thread.sleep(10000);
	  
	   // conform city
	   driver.findElement(By.id("ddlareaname")).click();
	   Thread.sleep(10000);
	   new Select(driver.findElement(By.id("ddlareaname"))).selectByVisibleText("Malad");
	   System.out.println("malad 2");
	   Thread.sleep(5000);
	   driver.findElement(By.id("txtAttributeValue")).sendKeys("*");
	    driver.findElement(By.id("btnSearch")).click();
	    
	    Thread.sleep(30000);
	    System.out.println("all input done");
	}
	static void indexPage(){
		
	}
	static void window() throws InterruptedException, IOException, RowsExceededException, WriteException{
		
	 
	   String strMainWindow = driver.getWindowHandle();
	 //  System.out.println("Window title" + driver.getTitle()); 
	   
	   // new Window Page
	   // click on Index Button 
	   String start="//*[@id='RegistrationGrid']/tbody/tr[";
	   String end = "]/td[10]/input";
	   
	   for(int i=2;i<12;i++){  
		   try{
			driver.findElement(By.xpath(start+i+end)).isDisplayed();
			driver.findElement(By.xpath(start+i+end)).click();
		   }catch(Exception e){continue;}
		   
		// clicked on all Index-II button on page one at a time
			   Thread.sleep(5000);
			   Set<String> strHandles = driver.getWindowHandles(); // store all open Window 
			   
			   
			   for(String handle:strHandles){  // Number of Window loop
				   driver.switchTo().window(handle);  // Switch to new Window
				   String strTitle = driver.getTitle();
				   
				   if (strTitle.equalsIgnoreCase("Index-II")) // new window process
				   {
					   if (k==1){
						   colName();
					   }else{}
					   getData(k);
					  k++;
				   }
				   
	   }driver.close(); // Close new Window
	   driver.switchTo().window(strMainWindow); // Switch back to main window
			   }
	}
	static void getData(int k) throws IOException{
	// Extract data from new window
	   System.out.println("Index :"+k);
	   WritableSheet writableSheet = writableBook.getSheet(0);
	   System.out.println(driver.getTitle());
	   Label data;
	try{
	   data = new Label(1,k,driver.findElement(By.xpath("html/body/table[1]/tbody/tr/td[1]/table/tbody/tr[1]/td/font")).getText());
	   writableSheet.addCell(data);
	   }catch(Exception e){  }
		try{
	   data = new Label(2,k,driver.findElement(By.xpath("html/body/table[1]/tbody/tr/td[1]/table/tbody/tr[2]/td/font")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(3,k,driver.findElement(By.xpath("html/body/table[1]/tbody/tr/td[3]/table/tbody/tr[1]/td")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(4,k,driver.findElement(By.xpath("html/body/table[1]/tbody/tr/td[3]/table/tbody/tr[2]/td")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(5,k,driver.findElement(By.xpath("html/body/table[1]/tbody/tr/td[3]/table/tbody/tr[3]/td")).getText());
	   writableSheet.addCell(data);
	   }catch(Exception e){  }
		try{
	   data = new Label(6,k,driver.findElement(By.xpath("html/body/table[1]/tbody/tr/td[3]/table/tbody/tr[4]/td")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(7,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[1]/td[2]")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(8,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[2]/td[2]")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(9,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[3]/td[2]")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(10,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[4]/td[2]/table/tbody/tr/td")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(11,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[5]/td[2]")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(12,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[6]/td[2]")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(13,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[7]/td[2]/table/tbody/tr[1]/td/font")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(14,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[7]/td[2]/table/tbody/tr[2]/td/font")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(15,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[8]/td[2]/table/tbody/tr[1]/td/font")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{ data = new Label(16,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[8]/td[2]/table/tbody/tr[2]/td/font")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(17,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[9]/td[2]")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(18,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[10]/td[2]")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(19,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[11]/td[2]")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(20,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[12]/td[2]")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{
	   data = new Label(21,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[12]/td[2]")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
		try{data = new Label(22,k,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[14]")).getText());
	   writableSheet.addCell(data);
		}catch(Exception e){  }
	
	   k++;
	   writableBook.write();
	   System.out.println("Index :"+k+"done");
} 
	static void colName() throws IOException, RowsExceededException, WriteException{
//	Create Excel sheet 
	Workbook.createWorkbook(fExcel);
		writableBook.createSheet("Data",0);
		WritableSheet writableSheet = writableBook.getSheet(0);
		Label data;
		
	   data = new Label(1,0,"Date");
	   writableSheet.addCell(data);
	   data = new Label(2,0,"Note");
	   writableSheet.addCell(data);
	   data = new Label(3,0,driver.findElement(By.xpath("html/body/table[1]/tbody/tr/td[3]/table/tbody/tr[1]/td/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(4,0,driver.findElement(By.xpath("html/body/table[1]/tbody/tr/td[3]/table/tbody/tr[2]/td/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(5,0,driver.findElement(By.xpath("html/body/table[1]/tbody/tr/td[3]/table/tbody/tr[3]/td/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(6,0,"Regn");
	   writableSheet.addCell(data);
	   data = new Label(7,0,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[1]/td[1]/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(8,0,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[2]/td[1]/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(9,0,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[3]/td[1]/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(10,0,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[4]/td[1]/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(11,0,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[5]/td[1]/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(12,0,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[6]/td[1]/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(13,0,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[7]/td[1]/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(15,0,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[8]/td[1]/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(17,0,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[9]/td[1]/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(18,0,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[10]/td[1]/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(19,0,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[11]/td[1]/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(20,0,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[12]/td[1]/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(21,0,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[13]/td[1]/font")).getText());
	   writableSheet.addCell(data);
	   data = new Label(22,0,driver.findElement(By.xpath("html/body/table[3]/tbody/tr[14]/td[1]/font")).getText());
	   writableSheet.addCell(data);
	   writableBook.write();
}
}
