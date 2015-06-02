package com.person;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.ConcurrentLinkedQueue;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.phantomjs.PhantomJSDriver;


public class GetWebInfoForAllPage {
	private static Logger logger = Logger.getLogger(GetWebInfoForAllPage.class);
private String urlPrex = "http://study.163.com";
  private String url = "http://study.163.com/find.htm#/find/courselist?ct=31001";
  private String phantomJsPath = "D:\\phantomjs\\phantomjs.exe";
  private WebDriver driver = null;
  private BufferedWriter bWriter = null;
  
  public static void main(String[] args) throws IOException {
	  GetWebInfoForAllPage getWebInfo = new GetWebInfoForAllPage();
	  getWebInfo.getWebInfo();
  }
  
  private void getWebInfo() throws IOException {
//	    driver = new FirefoxDriver();
//	    driver = new HtmlUnitDriver();
//	  driver = new ChromeDriver();
	  
	  System.setProperty("phantomjs.binary.path", phantomJsPath);
	  driver = new PhantomJSDriver();
	  try {
		  driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
		  driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
		  driver.get(url);
//		  java.io.File screenShotFile = ((TakesScreenshot) driver)
//				  .getScreenshotAs(OutputType.FILE);
//		  FileUtils.copyFile(screenShotFile, new java.io.File("D:\\a1.png"));
	  } catch (Exception e) {
	  }

//	    String pageSource = driver.getPageSource();
//	    System.out.println(pageSource);
	  bWriter = new BufferedWriter(new FileWriter("D:\\url.txt"));
	  int successPageCnt = 0;
	  for (int i = 1; i <= 50; i++) {
		  if (i > 1) {
			  List<WebElement> elementList = driver.findElements(By.xpath("//a[@href='#']"));
			  if (elementList != null) {
				  for (int j = 0; j < elementList.size(); j++) {
					  WebElement element = elementList.get(j);
					  String pageNoStr = element.getText();
					  if (String.valueOf(i).equals(pageNoStr)) {
						  element.click();
						  try {
							Thread.sleep(5000);
						} catch (InterruptedException e) {
							e.printStackTrace();
						}
						  java.io.File screenShotFile = ((TakesScreenshot) driver)
						  .getScreenshotAs(OutputType.FILE);
				  FileUtils.copyFile(screenShotFile, new java.io.File("D:\\image\\a" + i + ".png"));
						  if(!write()) {
							  System.out.println("get this page info err, pageNo=" + i);
						  } else {
							  successPageCnt ++;
						  }
						  break;
					  }
				  }
			  }
		  } else {
			  java.io.File screenShotFile = ((TakesScreenshot) driver)
					  .getScreenshotAs(OutputType.FILE);
			  FileUtils.copyFile(screenShotFile, new java.io.File("D:\\image\\a" + i + ".png"));
			  if(!write()) {
				  System.out.println("get this page info err, pageNo=" + i);
			  } else {
				  successPageCnt ++;
			  }
		  }
	  }
	  bWriter.flush();
	  bWriter.close();
	  System.out.println("successPageCnt=" + successPageCnt);
	  driver.close();
	  driver.quit();
  }
  
  private boolean write() {
	  boolean flag = false;
	  List<WebElement> list = driver.findElements(By.className("j-href"));
	  try {
		  if (list != null) {
			  for (WebElement e:list) {
				  String data_href = e.getAttribute("data-href");
				  String url = urlPrex + data_href;
				  bWriter.write(url);
				  bWriter.newLine();
				  flag = true;
			  }
		  }
	  } catch (IOException e) {
		  e.printStackTrace();
		  flag = false;
		  System.out.println("write url err, msg=" + e.getMessage());
	  }
	  
	  return flag;
  }
}
