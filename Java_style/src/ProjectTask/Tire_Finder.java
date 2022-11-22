package ProjectTask;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
public class Tire_Finder
{	
	static WebDriver driver  ;
	static XSSFWorkbook wb = new XSSFWorkbook();
	static Cell cell ;
	static List<String> sheetData=new ArrayList<String>();
	static Row row;
public static void main(String[] args) throws IOException, InterruptedException
	{
	System.setProperty("webdriver.chrome.driver","C:\\Users\\eclipse-workspace\\Selenium_project\\Driver\\chromedriver.exe"); //driver path
	driver = new ChromeDriver();
	driver.manage().window().maximize();
	read();
	driver.quit();
	System.out.println("Created..");
	}
public static void read() throws IOException, InterruptedException

	{
		FileInputStream fis = new FileInputStream("C:\\Users\\eclipse-workspace\\Report\\Tirefinder.xlsx"); //file path
		wb = new XSSFWorkbook(fis);
		XSSFSheet s = wb.getSheet("Sheet2");
		int rowcount = s.getLastRowNum()-s.getFirstRowNum();
		System.out.println(rowcount);
		int outputsheet =wb.getSheetIndex("sheet1");
		try
		{
			 wb.removeSheetAt(outputsheet);
		}
		catch(Exception e) {}
		for ( int i = 1 ; i <rowcount+1;i++)
			{
				Row r1 = s.getRow(i);
				String url = r1.getCell(0).getStringCellValue();
				sheetData.removeAll(sheetData);
				//driver.manage().timeouts().pageLoadTimeout(40, TimeUnit.SECONDS);
				driver.get(url);
				try
				{
					WebElement popup =driver.findElement(By.xpath("//a[@href='#' and @class='ui-dialog-titlebar-close ui-corner-all']"));
					if(popup.isDisplayed())
					{
						popup.click();
					}
				}
				catch(Exception e){}
				sheetData.add(url);
				TireFinder();
				Report(sheetData,i,"C:\\Users\\eclipse-workspace\\Report\\Tirefinder.xlsx","Sheet1"); // file path
			}
	}
public static void TireFinder() throws InterruptedException
{
	JavascriptExecutor js  = (JavascriptExecutor)driver;
	String Url ="https://www.amitirefinder.com/js/dist/tirefinder2.js";
	List<WebElement> Tire = driver.findElements(By.xpath("//a[contains(text(),'Tire')] |//a[contains(@href,'/service/tire-care-advice.htm')] |//a[contains(text(),'General Service')]"));
	if (Tire.size()!=0)
	{
		WebElement last = Tire.get(Tire.size() - 1); 
			for (  WebElement link : Tire)
			{	
				try
					{
						if ( link.getAttribute("href").contains("Finder")||link.getAttribute("href").contains("finder")||link.getAttribute("href").contains("details")||link.getAttribute("href").contains("tire")||link.getAttribute("href").contains("service") ==true)
							{
								try
									{ 
										js.executeScript("arguments[0].click();",  link);
										Thread.sleep(5000);
											if ( driver.getPageSource().contains(Url))
											{
													sheetData.add(Url);
													sheetData.add("Yes");
													sheetData.add("Pass");	
													break;
											}
											else
											{
												try
												{
													WebElement LearnMore=driver.findElement(By.xpath("/descendant::a[contains(@href,'/tire-care.htm')][2]"));
													LearnMore.click();
											//if (driver.getPageSource().contains("/tire-care.htm"))
													if ( driver.getPageSource().contains(Url))
													{
															sheetData.add(Url);
															sheetData.add("Yes");
															sheetData.add("Pass");	
															break;
													}
												}
												catch(NoSuchElementException e) {}
											}
										}
								catch(StaleElementReferenceException s){}
								}
					}
					catch(StaleElementReferenceException s){}	
		  		Fail( Url, last, link);
	            }
	        }
	else
	{
		sheetData.add(Url);
		sheetData.add("No");
		sheetData.add("Fail");
	}
	
}
public static void Report( List<String>list,int i ,String value,String sheetName) throws IOException
{	
	XSSFSheet sheet  = wb.getSheet(sheetName);
	if (sheet == null)
	{
		sheet = wb.createSheet(sheetName);
		createheader(sheet);
	}
			 row = sheet.createRow(i);
			int colnum=0;
			for(int j=0;j<=list.size();j++)
			{
				cell = row.createCell(colnum++);
				
				try
				{
					cell.setCellValue((String)list.get(j));
				}
				catch(Exception e)
				{
					cell.setCellValue((String)" ");
				}
			
		}
			XSSFCellStyle style = wb.createCellStyle();    
	    	XSSFCellStyle style1 = wb.createCellStyle();
			color(style,style1);
	    	FileOutputStream fos = new FileOutputStream(value);
	    	wb.write(fos);
	    	fos.close();	
}
public static void createheader(XSSFSheet sheet)
		{
			Object Testcase[][]= 
				{
					{"url","Tirefinder","yes/no","TestResult","Comment"},
				};
			int rowcount =0;
			for(Object emp[] : Testcase)  
				{	  
					 row = sheet.createRow(rowcount);
					int colcount=0;
					for (Object value : emp)
						{
							cell = row.createCell(colcount++);
							 
							  if ( value instanceof String)
							  {
								  cell.setCellValue((String)value);
								  
							  }
							  if ( value instanceof Integer)
							  {
								  cell.setCellValue((Integer)value);
								 
							  }
							  if ( value instanceof Boolean)
							  {
								  cell.setCellValue((Boolean)value);
							  }
						}
					}
			}	
public static void color(XSSFCellStyle style,XSSFCellStyle style1)
 	{
	int j=1;
	while( j <row.getLastCellNum())
		{
		cell = row.getCell(j);
		String value  =row.getCell(j).getStringCellValue();	
		if (value.equalsIgnoreCase("Pass"))
	 		{
		 	style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
			style.setFillPattern(FillPatternType.FINE_DOTS);
			cell.setCellStyle(style); 
	 		}
		else if(value.equalsIgnoreCase("Fail"))
 			{
 			style1.setFillBackgroundColor(IndexedColors.ROSE.getIndex());  
 			style1.setFillPattern(FillPatternType.FINE_DOTS);  
 			cell.setCellStyle(style1);      
 			}
		j++;
	}
 	}
public static void Fail(String Url,WebElement last,WebElement link)
{
	
	if(last==link )
	{
		sheetData.add(Url);
		sheetData.add("No");
		sheetData.add("Fail");
	}
}
}

