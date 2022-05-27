package com.sdet34l1.genericUtility;

import java.io.File;
import java.io.IOException;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import javax.lang.model.util.Elements;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.mysql.cj.jdbc.Driver;

/**
 * This class is used to maintain all the Webdriver common actions
 * @author VAISHALI
 *
 */
public class WebDriverOffice {
	static 	Select select;
	static JavascriptExecutor js;
    static JavaOffice j=new JavaOffice();

	/**
	 * This method is used to navigate to the application
	 * @param url
	 * @param driver
	 */
public static void navigateApp(String url,WebDriver driver)
{
	driver.get(url);
}
/**
 * This method is used to maximize the browser and implicitly wait
 * @param longtimeout
 * @param driver
 */
public static void browserSettings(long longtimeout,WebDriver driver)
{
	maximizeBrowser(driver);
	waitTillPageLoad(longtimeout,driver);

}
/**
 * This method is used to maximize the browser
 * @param driver
 */
public static void maximizeBrowser(WebDriver driver)
{
	driver.manage().window().maximize();
}
/**
 * This method is used to implicitly wait and wait Till Page Load
 * @param longtimeout
 * @param driver
 */
public static void waitTillPageLoad(long longtimeout,WebDriver driver)
{
	driver.manage().timeouts().implicitlyWait(longtimeout, TimeUnit.SECONDS);

}


/**
 * This method is used to mouse hover on the element
 * @param morelink
 * @param driver
 */
public static void mouseHoverOntheElement(WebElement morelink, WebDriver driver)
{
	Actions act=new Actions(driver);
	act.moveToElement(morelink).perform();
}
/**
 * This method is used to double click on the element
 * @param morelink
 * @param driver
 */
public static void doubleClick(WebElement morelink, WebDriver driver)
{
	Actions act = new Actions(driver);
	act.doubleClick(morelink).perform();
}

/**
 * This method is used to close the browser
 * @param driver
 */
public static void quitBrowser(WebDriver driver)
{
	driver.quit();
}
/**
 * This method is used to wait
 * @param timeout
 * @param driver
 */
public static void explicitlyWait(long timeout,WebDriver driver)
{
	//WebDriverWait wait = 
	new WebDriverWait(driver, timeout);
}
/**
 * This method is used to switch the window
 * @param driver
 * @param partialTitle
 */
public static void switchToWindow(WebDriver driver,String partialTitle)
{
	Set<String> windows = driver.getWindowHandles();
	for(String win:windows)
	{
		driver.switchTo().window(win);
		if(driver.getTitle().contains(partialTitle))
		{
			break;
		}
	}
}
/**
 * This method is used to select the value in dropdown  by value
 * @param ele
 * @param text
 */
public static void selectByValue(WebElement element, String text) 
{
	select=new Select(element);
	select.selectByValue(text);
}
/**
 * This method is used to select the value in dropdown  by index
 * @param element
 * @param index
 */
public static void selectByIndex(WebElement element, int index) 
{
	select=new Select(element);
	select.selectByIndex(index);
}
/**
 * This method is used to select the value in dropdown  by visible text
 * @param visibletext
 * @param element
 */
public static void selectByVisibleText( String visibletext,WebElement element) 
{
	select=new Select(element);
	select.selectByVisibleText(visibletext);
}
/**
 * This method is used to switch the frame by index
 * @param driver
 * @param index
 */
public static void switchToFrame(WebDriver driver,int index )
{
	  driver.switchTo().frame(index);

}
/**
 * This method is used to switch the frame by name or id
 * @param nameorId
 * @param driver
 */
public static void switchToFrame(int nameorId,WebDriver driver )
{
	  driver.switchTo().frame(nameorId);
}
/**
 * This method is used to switch the frame by element
 * @param element
 * @param driver
 */
public static void switchToFrame(WebElement element,WebDriver driver )
{
	  driver.switchTo().frame(element);
}
/**
 * This method is used to switch to mainpage
 * @param driver
 */
public static void switchBackToHomepage(WebDriver driver )
{
	  driver.switchTo().defaultContent();
}
/**
 * This method is used to initialize Action Class
 */
public static void initializeActionClass(WebDriver driver)
{
	Actions act=new Actions(driver);
    act =new Actions(driver);
}
/**
 * This method is used to initialize Java script
 * @param driver
 */
public static void initializeJscript(WebDriver driver)
{
	js=(JavascriptExecutor)driver;
}
/**
 * This method is used to enter Data Through Java script
 * @param element
 * @param data
 */
public static void enterDataThroughJscript(WebElement element, String data)
{
	js.executeScript("arguments[0].value=arguments[1]", element,data);
}
/**
 * This method is used to  navigate App through Java script
 * @param url
 */
public static void navigateAppthroughJscript(String url)
{
	js.executeScript("window.location=arguments[0]",url);
	
}
/**
 * This method is used to click Through Java script
 * @param element
 */
public static void clickThroghJscript(WebElement element)
{
	js.executeScript("arguments[0].click()",element);
}
/**
 * This method is used to scroll By Specific Height Java script
 * @param height
 */
public static void scrollBySpecificHeightJscript(String height)
{
	js.executeScript("window.scrollBy(0,"+height+")");
}
/**
 * This method is used to scroll Title Element Java script
 */
public static void scrollTitleElementJscript(WebElement element)
{
	js.executeScript("arguments[0].scrollIntoView()",element);
}
/**
 * This method is used to scroll To Bottom
 */
public static void scrollToBottom()
{
	js.executeScript("window.scrollBy(0,document.body.scrollHeight)");
}
/**
 * This method is used to  alert Accept
 * @param driver
 */
public static void alertAccept(WebDriver driver)
{
	driver.switchTo().alert().accept();
}
/**
 *  This method is used to alert Dismiss
 * @param driver
 */
public static void alertDismiss(WebDriver driver)
{
	driver.switchTo().alert().dismiss();
}
/**
 * This method is used to alert Send Data
 * @param driver
 * @param data
 */
public static void alertSendData(WebDriver driver,String data)
{
	driver.switchTo().alert().sendKeys(data);
}
/**
 *  This method is used to get Alert Text
 * @param driver
 * @return
 */

public static void WaitvisibleText(WebDriver driver,long timeout,WebElement ele)
{
	WebDriverWait wait = new WebDriverWait(driver, timeout);
	wait.until(ExpectedConditions.visibilityOf(ele));
}

public static String getAlertText(WebDriver driver)
{
	return driver.switchTo().alert().getText();
}

/**
 * This method is used to take Screenshot and store it in a file
 * @param fileName
 * @param driver
 * @return 
 * @throws IOException
 */
public static String takesScreenshot(String fileName,WebDriver driver) 
{
TakesScreenshot ts=(TakesScreenshot) driver;
File src = ts.getScreenshotAs(OutputType.FILE);
File dst = new File("./screenshot/"+fileName+"_"+j.dateTimeInFormat()+".png");
//System.out.println(dst.getAbsolutePath());
try {
	FileUtils.copyFile(src, dst);
} catch (IOException e) {
	// TODO Auto-generated catch block
	e.printStackTrace();
}
return dst.getAbsolutePath();
}
}
