package com.qa.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.XSSFCell;
//import org.apache.poi.xssf.usermodel.XSSFCellStyle;
//import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONObject;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.Reporter;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Hyperlink;
public class AmazonSearch {
	public static WebDriver driver;
	public static Boolean Status;
	public static String filepath="C:\\eclipse\\git\\MAVENDEMO\\";
	public static String filepath1;
	public static String fileName="data.xlsx";
	public static String writablefileName="data1.xlsx";
	public static String fileName1;
	public static String sheetName="AmazonSheet";
	public static String sheetName1="AmazonSearchSheet";
	public static JSONObject jsondata;
	public static List<WebElement> list ;    
	private static String[] columns = {"productName", "productScreenshot"};
	private static String screenshotPath;
	//private static String hyperlink,hyperlink_1,hyperlink_2;
	
	public static XSSFRow cell;
	public static List<Product> ProductListObj =  new ArrayList<Product>();
	public static boolean waitUntilElementPresent(By elementLocator) {
        try {
        	for (int second=0;second>120;second ++)
    		{
        		WebElement element=driver.findElement(elementLocator);
        		if (element.isDisplayed())
        			break;
    		} 
        	return true;
        }
        catch (NoSuchElementException e) {
            return false;
        }
    }
	
	public static void ElementClick(By elementLocator)
	{
		Status=waitUntilElementPresent(elementLocator);
		if(Status)
		{
			driver.findElement(elementLocator).click();
		}
	}
	
	public static void InputText(By elementLocator,CharSequence alpha )
	{
		Status=waitUntilElementPresent(elementLocator);
		if(Status)
		{
			driver.findElement(elementLocator).clear();
			driver.findElement(elementLocator).sendKeys(alpha);
		}
	}
	
	public static List retrieveElements(By elementLocator)
	{
		list= new ArrayList<WebElement>();
		
		list=driver.findElements(elementLocator);
		
		
		return list;
	}
	public static void TakeScreenshot(WebDriver driver, String filePath) throws Exception
	{
		Boolean manageWindowSize=false;
		TakesScreenshot screenShot=(TakesScreenshot)driver;
		 File SourceFile=screenShot.getScreenshotAs(OutputType.FILE);
		 File DesinationFile=new File(filePath);
		 FileUtils.copyFile(SourceFile, DesinationFile);
		 Reporter.log("<br><img src='"+DesinationFile+"' height='400' width='400' /></br>" );
	
	}
	
	public static void scrollDown(int alpha)
	{
		JavascriptExecutor js=(JavascriptExecutor)driver;
		js.executeScript("window.scrollBy(0,220)");
		
	}
	

public static Map<String,String> readExcel(String filePath,String fileName,String sheetName) throws IOException{
		
	   Map<String,String> ExcelData=new LinkedHashMap<String, String>();
	
		FileInputStream inputStream=null;
		try
		{
			File file =    new File(filePath+"\\"+fileName);
		
		inputStream = new FileInputStream(file);
		}
		catch(FileNotFoundException EX)
		{
			EX.printStackTrace();
		}
		
		
		XSSFWorkbook enlightedWorkbook = new XSSFWorkbook(inputStream);
		
		XSSFSheet enlightedSheet = enlightedWorkbook.getSheet(sheetName);
		
		Iterator<Row> rowIterator=enlightedSheet.iterator();
		
		while (rowIterator.hasNext()) {
			
		      Row row = rowIterator.next();
		      
		      Iterator<Cell> cellIterator =row.cellIterator();
		      
		      while (cellIterator.hasNext()) {
		    	  
		          Cell cell = cellIterator.next();
		          cell.setCellType(CellType.STRING);
		          ExcelData.put(cell.getStringCellValue(),cellIterator.next().getStringCellValue()); 
	                }
		          
		      }
		      System.out.println();
		      inputStream.close();
		      return ExcelData;
		}

public static void writeExcel(String filePath,String fileName,String sheetName,String fileNameParse) throws IOException{

	System.out.println(fileNameParse);
	
	/*Map<String,String> data=AmazonSearch.readExcel(filepath, fileName, sheetName);	
	
	for (int alpha=0;alpha<data.size();alpha++)
	{
	 columns = new String[] {data.get("PRODUCT_"+alpha)};
	}*/
	XSSFWorkbook enlightedWorkbook1 = new XSSFWorkbook();
	XSSFSheet enlightedSheet = enlightedWorkbook1.createSheet(sheetName1);	
	CreationHelper createHelper = enlightedWorkbook1.getCreationHelper();
	
	
	 // Create Other rows and cells with Products data
    int rowNum = 1;
    XSSFRow row ; 
    File filePDF;
    Hyperlink link;

       	
        cell = enlightedSheet.createRow(0);
        cell.createCell(0).setCellValue(columns[0]);
        System.out.println(columns[0]);
        cell.createCell(1).setCellValue(columns[1]);
        System.out.println(columns[1]);
    
    for(Product product: ProductListObj) {
    	row = enlightedSheet.createRow(rowNum);
    	
        row.createCell(0).setCellValue(product.getproductName());
        
        link= createHelper.createHyperlink(HyperlinkType.FILE);
        
        filePDF = new File(product.getproductScreenshot());
        
        link.setAddress(filePDF.toURI().toString());
        
        System.out.println(filePDF.toURI().toString());   
        
        //link.setLabel(product.getproductScreenshot());
        
       // row.createCell(1).setCellValue(convertfilePDF.toURI().toString()); 
        
        row.createCell(1).setHyperlink(link);       
        row.createCell(1).setCellValue(filePDF.toURI().toString());
        rowNum++;
    	}
    
 // Resize all columns to fit the content size
    for(int i = 0; i <=columns.length; i++) {
    	enlightedSheet.autoSizeColumn(i);
    }
   
    
	// cell.setHyperlink(href);
    
 // Write the output to a file
    FileOutputStream fileOut = new FileOutputStream(writablefileName);
    enlightedWorkbook1.write(fileOut);
    fileOut.close();

    // Closing the workbook
    enlightedWorkbook1.close();
}
		
	@BeforeTest
	public void openBrowser() throws FileNotFoundException, IOException{
		
		
		//System.out.println(data.get("PROD_1"));
		driver=new FirefoxDriver();
		System.setProperty("webdriver.firefox.bin",
                    "C:\\Program Files\\Mozilla Firefox\\firefox.exe");
		System.setProperty("webdriver.gecko.driver","C:\\eclipse\\geckodriver-v0.20.0-win64\\geckodriver.exe");
		driver.get("https://www.amazon.in/");
		driver.manage().window().maximize();
		
		if(AmazonSearch.waitUntilElementPresent(By.cssSelector("div.nav-sprite-v1.nav-bluebeacon.nav-packard-glow")))
		{
				System.out.println("Amazon page has been successfully reverted");
		}	
		
	}
	
	@Test
	public void ReadExcelDataAndSearchtheProductThenWriteOutputToExcel() throws Exception
	{
		
		Map<String,String> data=AmazonSearch.readExcel(filepath, fileName, sheetName);
		
		/*for (int a=1;a<data.size();a++)
		{
		String[] columns = {data.get("PRODUCT_"+a)};
		}
		*/
		for (int alpha=1;alpha<data.size();alpha++)
		{
		AmazonSearch.ElementClick(By.xpath("//div[@id='navbar']/descendant::div[@class='nav-search-field ']/input"));
		AmazonSearch.InputText(By.cssSelector("input#twotabsearchtextbox"),data.get("PRODUCT_"+alpha));
		AmazonSearch.ElementClick(By.cssSelector("div#navbar div#nav-search input.nav-input[value='Go']"));
		AmazonSearch.ElementClick(By.cssSelector("div#nav-xshop"));
		AmazonSearch.TakeScreenshot(driver, "C:\\eclipse\\git\\MAVENDEMO\\target\\Product\\Screen"+alpha+".jpeg");
		Thread.sleep(2000);
		list=AmazonSearch.retrieveElements(By.xpath("//div[@id='atfResults']/descendant::h2"));
		String[] ProductList=new String[list.size()];
		System.out.println(list.size());
		int i=0;
		for(WebElement e:list)
		{
			System.out.println(e);
			ProductList[i]=e.getText();
			i++;
		}
		
		System.out.println(ProductList);
		//test each link		
		for (String t : ProductList) {	
			System.out.println(t);
			AmazonSearch.scrollDown(alpha);
			t=t.replace("/", " ");
			t=t.replace("|", " ");
			t=t.replace("\"", " ");
			filepath1="C:\\eclipse\\git\\MAVENDEMO\\target\\EachProduct"+alpha+"\\";
			fileName1=t+".jpeg";
			screenshotPath=filepath1+fileName1;				
			AmazonSearch.TakeScreenshot(driver,filepath1+fileName1);	
			ProductListObj.add(new Product(data.get("PRODUCT_"+alpha),screenshotPath));
			System.out.println(screenshotPath);
			AmazonSearch.writeExcel(filepath,fileName, sheetName,fileName1.toString());
				}							
		
		}
	}

	@AfterTest
	public void closeBrowser()
	{
	driver.quit();
	}
}