import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.lang.String;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import java.io.File;
import java.io.IOException;
import java.sql.Date;

import jxl.Workbook;
import jxl.write.DateTime;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import org.openqa.selenium.JavascriptExecutor;

public class magicBricks4 {

	
	//private static final String String = null;

	public static void main(String[] args) throws IOException, RowsExceededException, WriteException, InterruptedException {
		
		WebDriver driver = new FirefoxDriver();
		//  url address of site
		driver.get("http://www.magicbricks.com/property-for-sale/residential-real-estate?proptype=Multistorey-Apartment,Builder-Floor-Apartment,Penthouse,Studio-Apartment&Locality=Andheri-West&searchLocType=&searchTransMode=&searchLocTime=&cityName=Mumbai&BudgetMin=1-Crores&BudgetMax=2-Crores&searchLocType=&searchTransMode=&searchLocTime=&price=Y&pageOption=");
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		
			JavascriptExecutor javascript = (JavascriptExecutor) driver;  
			for(int j=1;j<1000;j++)
		
			{
				javascript.executeScript("window.scrollTo(0, document.body.scrollHeight);");
		//	Thread.sleep(500);
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		
			
			}
		// Extract all basket and list 
		List<WebElement> baskets = new ArrayList<WebElement>(); 
		
		baskets.addAll(driver.findElements(By.xpath("//*[starts-with(@id,'priceProperty')]")));
		System.out.println(baskets.size());

		List<WebElement> room = new ArrayList<WebElement>(); 
		
		room.addAll(driver.findElements(By.xpath("//*[starts-with(@id,'resultBlockWrapper')]")));
		System.out.println(room.size());

		
		
		File fExcel = new File ("C:\\Users\\SID\\Desktop\\selenium test\\room.xls");
		
		WritableWorkbook writableBook = Workbook.createWorkbook(fExcel);
		
		writableBook.createSheet("Data",0);
		
		
		WritableSheet writableSheet = writableBook.getSheet(0);


		List<WebElement> all = new ArrayList<WebElement>();
	//	List<WebElement> tds = new ArrayList<WebElement>();
 all = driver.findElements(By.xpath("//*[@id='srpColumnsConWrap']"));
 List<WebElement> allDivs = driver.findElements(By.xpath("//*[starts-with(@id,'resultBlockWrapper')]"));
		
		for(int i=0;i<room.size(); i++){
			System.out.println("***************************");
			System.out.println(baskets.get(i).getText()); // prints all the block data 
//			String tds = ((WebElement) room.get(i).findElements(By.className("proDetailsRowElm"))).getText();
	//		System.out.println(tds);
		//	System.out.println(((WebDriver) room).findElements(By.xpath('./div[5]/div[1]/ul/li[3]/div/div[1]/div')));
				System.out.println(room.get(i).getText());
	//	System.out.println(all.get(i).getText());
			System.out.println("***************************");

		
	//		label = baskets.get(i).getText();
	//		date = room.get(i).getText();
			
			Label label1 = new Label(0, i, baskets.get(i).getText());
	   //     Label date = new Label(1, i,baskets.get(i).getText());
		
			 writableSheet.addCell(label1);
				}
		
	//	System.out.println(baskets);
		writableBook.write();
		writableBook.close();
				
	}
}
