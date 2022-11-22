package ProjectTask;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import io.github.bonigarcia.wdm.WebDriverManager;
public class AdChoice

	{	
	
	static WebDriver driver  ;
	static JavascriptExecutor js;
	static XSSFWorkbook wb = new XSSFWorkbook();
	static Boolean Adchoice_link ,Adchoice_img ;
	static Cell cell ;
	static WebElement Adchoice;
	static	WebElement Adchoice_image;
	static List<String> sheetData=new ArrayList<String>();
	static Row row;
	
public static void main(String[] args) throws IOException, InterruptedException
	{
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		Actions a = new Actions(driver);
		js  = (JavascriptExecutor)driver;
		read(a);
		driver.quit();
		System.out.println("Created..");
	}
public static void read(Actions a) throws IOException
	{
	//C:\\Users\\pm00817952\\Desktop\\Demo.xlsx
		FileInputStream fis = new FileInputStream("C:\\Users\\eclipse-workspace\\Report\\Adchoice.xlsx"); // file path
		wb = new XSSFWorkbook(fis);
		XSSFSheet s = wb.getSheet("Sheet2");
		int rowcount = s.getLastRowNum()-s.getFirstRowNum();
		System.out.println(rowcount);
		int outputsheet =wb.getSheetIndex("sheet1");
		try
		{
			 wb.removeSheetAt(outputsheet);	
		}
		catch(Exception e){}
	
		for ( int i = 1 ; i <rowcount+1;i++)
			{
				Row r1 = s.getRow(i);
				String url = r1.getCell(0).getStringCellValue();
				sheetData.removeAll(sheetData);
				driver.get(url);
				try
				{
					try
					{
						WebElement popup =driver.findElement(By.xpath("//a[@href='#' and @class='ui-dialog-titlebar-close ui-corner-all']"));
						if(popup.isDisplayed())
						{
							popup.click();
						}
					}
					catch(Exception e){}
					try
					{
						WebElement popup1 =driver.findElement(By.xpath("//button[@class='close']"));
						if(popup1.isDisplayed())
						{
							popup1.click();
						}

					}
					catch(Exception e) {}
				}
				catch(Exception e)
				{
					// an empty block
				}
				a.keyDown(Keys.CONTROL).sendKeys(Keys.END).build().perform();
				Adchoice();
				sheetData.add(url);
				MethodFunctionality(Adchoice_link,Adchoice_img);
				Report(sheetData,i,"C:\\Users\\eclipse-workspace\\Report\\Adchoice.xlsx","Sheet1");// file path
			}
	}
public static void Adchoice()
{
		try 
			{   
				Adchoice =driver.findElement(By.xpath("//a[contains(@href,'https://www.forddirect.com/adchoices')] | //*[contains(@class,'adChoices')]"));
				Adchoice_link = Adchoice.isDisplayed();
				try
				{
					Adchoice_image =driver.findElement(By.xpath("//img[contains(@src,'adchoice/adchoice-new')] |//img[contains(@src,'AdMarker_Icon_Text_')] |//img[contains(@src,'AdMarker_Icon_Text_') and //@class='lazy' and @id='_bapw-icon']"));
					Adchoice_img = Adchoice_image.isDisplayed();
					if( Adchoice_image.isDisplayed()==false)
					{
						Adchoice_img = Adchoice_image.isEnabled();
					}
				}
				catch(NoSuchElementException e)
				{
					Adchoice_img=false;
				}
				
			}
		catch ( Exception e){}
		finally
		{
				js.executeScript("arguments[0].click();", Adchoice);
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
					{"url","Adchoice","CompareUrl","Newtabe","Comment"},
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
		else if(value.equalsIgnoreCase("fail"))
 			{
 			style1.setFillBackgroundColor(IndexedColors.ROSE.getIndex());  
 			style1.setFillPattern(FillPatternType.FINE_DOTS);  
 			cell.setCellStyle(style1);      
 			}
		j++;
	}
 	}	
public static void MethodFunctionality(boolean Ad ,boolean Adimg)
{
		if (Ad && Adimg ==true)
		{
			sheetData.add("Pass");
	    }
		else
		{
			sheetData.add("Fail");
			
		}

	  	String defaultWindow = driver.getWindowHandle();
	  	Set<String> tab = driver.getWindowHandles(); 
	  	int tabsize = tab.size();
	  	
	  	for(String newtab : tab)  	
	  		{
	  			if(!newtab.equalsIgnoreCase(defaultWindow)  )
	  				{
	  					driver.switchTo().window(newtab);
	  					sheetData.set(2,"Pass");   
	  					
	  				}
	  			else
	  				{
	  				sheetData.add("Fail");
	  				}
	  	}

	  	String currentUrl=driver.getCurrentUrl();
	  	String Checkingurl="https://www.forddirect.com/adchoices";
	  	if( currentUrl.equals(Checkingurl))
   		{
	  		sheetData.add("Pass");
	  		}
	  	else
	  		{
	 		sheetData.add("Fail");
	 		sheetData.add("Title not same as given url");
	  		} 
	  	if(Adimg==false)
	  	{
	  		sheetData.add("Image is not Present...");
	  	}
	  	else if (tabsize<2)
	  	{
	  		sheetData.add("link not open in  newTab");
	  	}
	  	
	 	  	  	
	  	if (tabsize>=2)
	  	{
 		 driver.close(); 
	  	 driver.switchTo().window(defaultWindow);
	  	}
}
	}