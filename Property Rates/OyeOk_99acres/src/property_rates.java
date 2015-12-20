import java.io.File;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.NotFoundException;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class property_rates {

	public static void main(String[] args) throws IOException, RowsExceededException, WriteException ,InterruptedException{
	
		int i=1;
		
WebDriver driver = new FirefoxDriver();
//url address of site//  url address of site
driver.get("http://www.99acres.com/property-rates-and-price-trends-in-mumbai");
driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

File fExcel = new File ("C:\\Users\\SID\\Desktop\\Oyeok\\99acres\\99acres2.xls");
WritableWorkbook writableBook = Workbook.createWorkbook(fExcel);
writableBook.createSheet("Data",0);
WritableSheet writableSheet = writableBook.getSheet(0);
Label data,data1;

String Start = "//*[@id='ptrtable']/div/table/tbody/tr[";
String End1= "]/th";
String End2= "]/td[1]";
boolean br = true;


do{
	data = new Label(1,i,driver.findElement(By.xpath(Start+i+End1)).getText());
	writableSheet.addCell(data);
	i++;
	br =driver.findElement(By.xpath(Start+i+End2)).isDisplayed();
	do{
		String End3="]/td[2]";
		String End5="]/td[5]";
		String End6="]/td[6]";
		String End7="]/td[7]";
		data = new Label(1,i,driver.findElement(By.xpath(Start+i+End2)).getText());
		writableSheet.addCell(data);
		data = new Label(2,i,driver.findElement(By.xpath(Start+i+End3)).getText());
		writableSheet.addCell(data);
		data = new Label(3,i,driver.findElement(By.xpath(Start+i+End5)).getText());
		writableSheet.addCell(data);
		data = new Label(4,i,driver.findElement(By.xpath(Start+i+End6)).getText());
		writableSheet.addCell(data);
		data = new Label(5,i,driver.findElement(By.xpath(Start+i+End7)).getText());
		writableSheet.addCell(data);
		System.out.println(i);
		i++;
		try{
		br =driver.findElement(By.xpath(Start+i+End2)).isDisplayed();
		
		}catch(Exception e){
			br=false;
			System.out.println(br);
			continue;
		     //throw new Error(e.getMessage());
		}
		System.out.println(br);
	}while(br==true);
	try{	br =driver.findElement(By.xpath(Start+i+End1)).isDisplayed();
	
}catch(Exception e){
	System.out.println(br);
	// System.out.println(br);
	continue;
  //  throw new Error(e.getMessage());
    
}
}while(br==true);


writableBook.write(); // Write on Excel Sheet
writableBook.close(); // Close Excel Sheet
driver.close();
	}

}
