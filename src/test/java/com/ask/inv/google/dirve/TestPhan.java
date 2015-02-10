package com.ask.inv.google.dirve;

import java.io.File;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.phantomjs.PhantomJSDriver;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.WebDriverWait;

public class TestPhan {
	 public static void main(String[] args) throws IOException {
	        // Create a new instance of the Firefox driver
	        // Notice that the remainder of the code relies on the interface, 
	        // not the implementation.
	        System.setProperty("phantomjs.binary.path", "C:\\phantomjs\\phantomjs.exe");
	        WebDriver driver = new PhantomJSDriver();
	        // And now use this to visit Google
	        driver.get("http://www.google.com");
	        // Alternatively the same thing can be done like this
	        // driver.navigate().to("http://www.google.com");

	        // Find the text input element by its name
	        WebElement element = driver.findElement(By.name("q"));

	        // Enter something to search for
	        element.sendKeys("Cheese!");

	        // Now submit the form. WebDriver will find the form for us from the element
	        element.submit();

	        // Check the title of the page
	        System.out.println("Page title is: " + driver.getTitle());
	        File screenShotFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);  
	        FileUtils.copyFile(screenShotFile, new File("./a.png"));

	        // Google's search is rendered dynamically with JavaScript.
	        // Wait for the page to load, timeout after 10 seconds
	        (new WebDriverWait(driver, 10)).until(new ExpectedCondition<Boolean>() {
	            public Boolean apply(WebDriver d) {
	                return d.getTitle().toLowerCase().startsWith("cheese!");
	            }
	        });

	        // Should see: "cheese! - Google Search"
	        System.out.println("Page title is: " + driver.getTitle());

	        //Close the browser
	        driver.quit();
	    }
}
