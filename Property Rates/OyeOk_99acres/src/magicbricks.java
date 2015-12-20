import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class magicbricks {

	public static void main(String args[]) throws IOException, RowsExceededException, WriteException, InterruptedException{

		Label label1,Label2,label3,Label4,label5;
		int k=1,j,x=1;
		WebDriver driver = new FirefoxDriver();
		driver.get("http://www.magicbricks.com/Property-Rates-Trends/ALL-RESIDENTIAL-rates-in-Mumbai");
		driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
		
		
//		Create Excel sheet 
			File fExcel = new File ("C:\\Users\\SID\\Desktop\\Oyeok\\magicbricks\\magicbricks2.xls");
			WritableWorkbook writableBook = Workbook.createWorkbook(fExcel);
			writableBook.createSheet("Data",0);
			WritableSheet writableSheet = writableBook.getSheet(0);
			Label data;
			boolean br=true;
			
			driver.findElement(By.xpath("//*[@id='pagination']/a[1]/b")).click();
			driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
			System.out.println("next");
			driver.findElement(By.xpath("//*[@id='pagination']/a[2]/b")).click();
			driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
			System.out.println("back");
			Thread.sleep(5000);
			
			do{
				Thread.sleep(5000);
				List<WebElement> locality = new ArrayList<WebElement>(); 
				locality.addAll(driver.findElements(By.xpath("//*[starts-with(@id,'localityName')]")));
				System.out.println(locality.size());
				
				List<WebElement> sale = new ArrayList<WebElement>(); 
				sale.addAll(driver.findElements(By.xpath("//*[starts-with(@id,'salePriceRange')]")));
				System.out.println(sale.size());
				
				List<WebElement> saleAvg = new ArrayList<WebElement>(); 
				saleAvg.addAll(driver.findElements(By.xpath("//*[starts-with(@id,'saleAvgPrice')]")));
				System.out.println(saleAvg.size());
				
				List<WebElement> Rent = new ArrayList<WebElement>(); 
				Rent.addAll(driver.findElements(By.xpath("//*[starts-with(@id,'rentPriceRange')]")));
				System.out.println(Rent.size());
				
				List<WebElement> RentAvg = new ArrayList<WebElement>(); 
				RentAvg.addAll(driver.findElements(By.xpath("//*[starts-with(@id,'rentPriceRange')]"))); //rentPriceRange rentPattern
				System.out.println(RentAvg.size());
				
				for(int i =0;i<=locality.size()-1;i++){
					System.out.print(i);

					
					label1 = new Label(1, x, locality.get(i).getText());
					writableSheet.addCell(label1);
					Label2 = new Label(2, x, sale.get(i).getText());
					writableSheet.addCell(Label2);
					label3 = new Label(3, x, saleAvg.get(i).getText());
					writableSheet.addCell(label3);
					Label4 = new Label(4, x, Rent.get(i).getText());
					writableSheet.addCell(Label4);
					label5 = new Label(5, x, RentAvg.get(i).getText());
					writableSheet.addCell(label5);
					System.out.println(" done"+i);
					x++;
				}
				String start= "//*[@id='pagination']/a[";
				String end= "]";
			if(k==1){
				System.out.println("***********");

					}
			if(k==2){
				k=3;
			}
			else{
				k++;
			}
			try{
				br = driver.findElement(By.xpath(start+k+end)).isDisplayed();
			}catch(Exception e){
				System.out.println(br);
				br=false;
				// System.out.println(br);
				continue;
			  //  throw new Error(e.getMessage());
			    
			}
				if(br==true){
				driver.findElement(By.xpath(start+k+end)).click();	
				}
				else{
					break;
				}
			}while(br==true);
			
			writableBook.write(); // Write on Excel Sheet
			writableBook.close(); // Close Excel Sheet
		
		driver.close();	
	}
}
