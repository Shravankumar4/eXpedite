package expediteFunctions;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.NoSuchElementException;
import java.util.concurrent.TimeUnit;

import java.io.File;
import java.io.FileInputStream;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class eXpedite {
	//private static final CharSequence EmailAddress = null;
	public static WebDriver oBrowser;
    public static Sheet sheet1,sheet2;
	public static int sheet1Cr,sheet1Rc,sheet2Cr,sheet2Rc,i;
    public static Workbook workbook;
    public static Row sheet1row, sheet2row;
    public static Cell sheet1cell,sheet2cell;
    public static FileInputStream file;
    public static int rowcount, cr;
    public static String timeStamp,newFilename,username,password,stack,testset,baselineVersion,EmailAddress,Emailid,EmailBox;
    public static File inputFileName;

    
    public void input() throws Exception
	 {
    	//Copying the Input Excel file to other Location and using that file as the Input File
        File source = new File("C://Automation//eXpedite//eXpedite.xlsx");
        timeStamp = new SimpleDateFormat("HHmmss").format(Calendar.getInstance().getTime());
        newFilename = "eXpedite" + timeStamp + ".xlsx";
        inputFileName = new File("C://Automation//eXpedite//Result//" + newFilename);
        FileUtils.copyFile(source, inputFileName);
        
        
        file = new FileInputStream(inputFileName);
        workbook = WorkbookFactory.create(file);

        // Reading Current Row and Row Count from Sheet1 and Sheet2
        sheet1 = workbook.getSheet("Sheet1");
        sheet1Cr = sheet1.getFirstRowNum();
        sheet1Rc = sheet1.getLastRowNum();
        sheet2 = workbook.getSheet("Sheet2");
        sheet2Cr = sheet2.getFirstRowNum();
        sheet2Rc = sheet2.getLastRowNum();
        
        sheet1row = sheet1.getRow(1);
        
        sheet1cell = sheet1row.getCell(0);
        username = sheet1cell.getStringCellValue();
        System.out.println("Username :"+ username);
        sheet1cell = sheet1row.getCell(1);
        password = sheet1cell.getStringCellValue();
        System.out.println("Password: "+password);
        sheet1cell = sheet1row.getCell(2);
        stack = sheet1cell.getStringCellValue();
        System.out.println("Stack: "+stack);
        sheet1cell = sheet1row.getCell(3);
        EmailAddress = sheet1cell.getStringCellValue();
        System.out.println("EmailBox: "+EmailAddress);
        
        switch (stack.toLowerCase()) 
        {
        case "dis"    : baselineVersion = "Fleet-DISStack";
                        break;
        case "test01" : baselineVersion = "Fleet-Test01";
        				break;
        case "pie"    : baselineVersion = "Fleet-PIE";
        				break;
        case "stage1" : baselineVersion = "FleetStage1";
        				break;
        case "ref1"   : baselineVersion = "Fleet-Ref01";
						break;
        case "ref2"   : baselineVersion = "Fleet-Ref02";
						break;
        }


     
        eXpedite.login();
        
        for (i = 1; i <= sheet2Rc; i++) 
        {
                    	
        	//Getting all the Input Values from Sheet1 and Sheet2
               sheet2row = sheet2.getRow(i);
               sheet2cell = sheet2row.getCell(1);
               String condition = sheet2cell.getStringCellValue();
               
               if (condition.equalsIgnoreCase("yes"))
               {
            	   sheet2cell = sheet2row.getCell(0);
                   testset = sheet2cell.getStringCellValue();
                   System.out.println("Executing the Test Set : "+testset);
            	   eXpedite.execute();
               }
		}
        System.out.println("Completed Execution");
      }     
               
	
    
    
	public  static void login() 
	{   
		oBrowser = new FirefoxDriver();
		oBrowser.get("http://expedite.ind.hp.com/eXpedite/");
		oBrowser.manage().timeouts().implicitlyWait(12000, TimeUnit.SECONDS);
		oBrowser.switchTo().frame("login");
		oBrowser.findElement(By.id("username")).sendKeys(username);
		oBrowser.findElement(By.id("password")).sendKeys(password);
		WebElement selectTeam = oBrowser.findElement(By.id("teamName"));
		selectTeam.sendKeys("Fleet-Team");
		oBrowser.findElement(By.cssSelector("input[type=\"submit\"]")).click();
		oBrowser.switchTo().frame("nav");
		oBrowser.findElement(By.linkText("Test Set Management")).click();
		oBrowser.findElement(By.linkText("Execute Test Set")).click();
		oBrowser.switchTo().defaultContent();
		oBrowser.switchTo().frame("login");
		oBrowser.switchTo().frame("content");
		oBrowser.findElement(By.xpath("/html/body/form/table/thead/tr/th/a")).click();
        
	}
	
	
	
	public static void execute() throws Exception

	{
		
			oBrowser.findElement(By.xpath("//input[@value='"+testset+"']")).click();
			WebElement selectedStack = oBrowser.findElement(By.id("stacks"));
			selectedStack.sendKeys(stack);
			WebElement Device = oBrowser.findElement(By.name("deviceType"));
			Device.sendKeys("Printer Simulator");
			Thread.sleep(3000);
			
			try {
				if(oBrowser.findElement(By.xpath("//input[@id='NotifyBox']")).isDisplayed())
				{
					//System.out.println("displayed");
				WebElement chkbox = oBrowser.findElement(By.xpath("//input[@id='NotifyBox']"));
					System.out.println(chkbox.getTagName());
					//chkbox.submit();
					chkbox.click();
				}
				else
				{
					System.out.println("not displayed");
			}
				oBrowser.findElement(By.xpath("//input[@id='NotifyBox']")).getText();
				oBrowser.findElement(By.xpath("//input[@id='NotifyBox']")).click();			
			} catch (NoSuchElementException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
			Thread.sleep(10000);
			
			WebElement Emailid = oBrowser.findElement(By.id("EmailBox"));
			Emailid.sendKeys(EmailAddress);
			System.out.println(EmailAddress);
			Thread.sleep(6000);
			WebElement baseline = oBrowser.findElement(By.name("baselineVersion"));
			baseline.sendKeys(baselineVersion);
			Thread.sleep(10000);
			oBrowser.findElement(By.xpath("/html/body/form/table[2]/tbody/tr[11]/td/input")).click();
			Thread.sleep(5000);
			oBrowser.findElement(By.xpath("html/body/table/tbody/tr/td/a")).click();
			Thread.sleep(300000);
			oBrowser.close();
			eXpedite.login();
		
				
	}
		 
}
