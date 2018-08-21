import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Kimai_Login 
{
	public static WebDriver driver;
	public void tc() throws Exception
	{
		Properties obj = new Properties();   
		FileInputStream objfile = new FileInputStream(System.getProperty("user.dir")+"\\objects.properties");
		obj.load(objfile); 
		obj.getProperty("driverType", "exeFile");
		
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(10,TimeUnit.SECONDS);
		driver.manage().window().maximize();
		driver.get(obj.getProperty("URL"));
		
		ArrayList<String> userName = readExcelData(0);
		ArrayList<String> password = readExcelData(1);
		ArrayList<String> projectName = readExcelData(2);
		ArrayList<String> day1 = readExcelData(3);
		ArrayList<String> day2 = readExcelData(4);
		ArrayList<String> time1 = readExcelData(5);
		ArrayList<String> time2 = readExcelData(6);
		
		for(int i=0; i<userName.size(); i++)
		{
			driver.findElement(By.name("name")).sendKeys(userName.get(i));
			driver.findElement(By.name("password")).sendKeys(password.get(i)); 
			driver.findElement(By.name("password")).sendKeys(Keys.ENTER);
			Thread.sleep(5000);
			for(int j=0; j<userName.size(); j++)
			{
				driver.findElement(By.linkText("add")).click();
				Thread.sleep(10000);    
	    		driver.findElement(By.cssSelector("input[tabindex='2']")).sendKeys(projectName.get(j));
	    		Select dropdown = new Select(driver.findElement(By.id("add_edit_zef_pct_ID")));
	    		List<WebElement> elements = dropdown.getOptions();
	    		Iterator<WebElement> itr = elements.iterator();
	    		while(itr.hasNext())
	    		{	
	    			WebElement row = itr.next();
	    			driver.findElement(By.cssSelector("select#add_edit_zef_pct_ID >option")).click();
	    			driver.findElement(By.cssSelector("option[value='7']")).click();
	    		}
	    		Thread.sleep(3000);
	    		driver.findElement(By.id("edit_in_day")).clear();
				driver.findElement(By.id("edit_in_day")).sendKeys(day1.get(j));
				
				driver.findElement(By.id("edit_out_day")).clear();
				driver.findElement(By.id("edit_out_day")).sendKeys(day2.get(j));
				
				driver.findElement(By.id("edit_in_time")).clear();
				Thread.sleep(2000);
				driver.findElement(By.id("edit_in_time")).sendKeys(time1.get(j).replaceAll("\"", ""));
				
				driver.findElement(By.id("edit_out_time")).clear();
				Thread.sleep(2000);
				driver.findElement(By.id("edit_out_time")).sendKeys(time2.get(j).replaceAll("\"", ""));
				Thread.sleep(2000);    			
				
				driver.findElement(By.cssSelector("input[value='OK']")).click();
				Thread.sleep(10000);			
			}
			break;
		}
		driver.findElement(By.id("main_logout_button")).click();
		Thread.sleep(5000);
		driver.close();
	}

	public ArrayList<String> readExcelData(int colNo) throws Exception
	{
		Properties obj = new Properties();   
		FileInputStream objfile = new FileInputStream(System.getProperty("user.dir")+"\\objects.properties");
		obj.load(objfile); 
		obj.getProperty("driverType", "exeFile");
		
		File file = new File(obj.getProperty("excelSheetLocation"));
	    FileInputStream inputStream = new FileInputStream(file);
	    Workbook wb = new XSSFWorkbook(inputStream);	  
	    Sheet sheet = wb.getSheet("Kimai");
	    
	    Iterator<Row> rowIterator = sheet.iterator();
	    rowIterator.next();
	    
	    ArrayList<String> list = new  ArrayList<String>();
	    while(rowIterator.hasNext())
	    {
	    	list.add(rowIterator.next().getCell(colNo).getStringCellValue());
	    }
	    return list;
	}
	
	public static void main(String[] args) throws Exception 
	{
		Kimai_Login login = new Kimai_Login();
		login.tc();
	}
}	
