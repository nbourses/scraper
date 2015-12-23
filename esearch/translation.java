
import java.io.File;
import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class translation {
	static WebDriver driver = new FirefoxDriver();

		
	public static void main(String[] args) throws BiffException, IOException, RowsExceededException, WriteException, InterruptedException {
		driver.get("https://translate.google.co.in/?hl=en&tab=TT");

//		Create Excel sheet 
			File fExcel = new File ("C:\\Users\\SID\\Desktop\\trans_esearch.xls");
			WritableWorkbook writableBook = Workbook.createWorkbook(fExcel);
			writableBook.createSheet("Data",0);
			WritableSheet writableSheet = writableBook.getSheet(0);
			Label data;

		
		File input = new File ("C:\\Users\\SID\\Desktop\\esearch.xls");
		Workbook writableBook1 = Workbook.getWorkbook(input);
		Sheet sheet1 =writableBook1.getSheet(0);
	for(int i=1;i<100;i++){	
		for(int j=0;j<22;j++){
		Cell col1 = sheet1.getCell(j, i);
        String scol1 = col1.getContents();
        
        String trans1 = trans(scol1);
        
        data = new Label(j,i,trans1);
		writableSheet.addCell(data);
		}
       
	}
	writableBook.write(); // Write on Excel Sheet
	writableBook.close(); // Close Excel Sheet
	driver.close();
		
	}

	
	
	public static String trans(String scol) throws InterruptedException{
		Thread.sleep(2000);
	    driver.findElement(By.xpath("//*[@id='source']")).clear();
	    driver.findElement(By.xpath("//*[@id='source']")).sendKeys(scol);
	    Thread.sleep(3000);
	    String trans = driver.findElement(By.xpath("//*[@id='result_box']")).getText();
	    Thread.sleep(2000);
	    return trans;

		
		
	}
}


