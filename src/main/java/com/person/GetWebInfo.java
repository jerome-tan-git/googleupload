package com.person;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Callable;
import java.util.concurrent.ConcurrentLinkedQueue;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;
import org.apache.log4j.Logger;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.phantomjs.PhantomJSDriver;
import org.openqa.selenium.remote.DesiredCapabilities;

public class GetWebInfo {
	private static Logger logger = Logger.getLogger(GetWebInfo.class);
	private String phantomJsPath = "D:\\phantomjs\\phantomjs.exe";
	private BufferedWriter bWriter = null;
	private ConcurrentLinkedQueue<String> urlQueue = new ConcurrentLinkedQueue<String>();
	private ConcurrentLinkedQueue<String> resultQueue = new ConcurrentLinkedQueue<String>();

	public static void main(String[] args) throws IOException {
		GetWebInfo getWebInfo = new GetWebInfo();
		getWebInfo.getWebInfo();
	}

	private void getWebInfo() throws IOException {
		//	    driver = new FirefoxDriver();
		//	    driver = new HtmlUnitDriver();
		//	  driver = new ChromeDriver();

		System.setProperty("phantomjs.binary.path", phantomJsPath);

		BufferedReader bReader = new BufferedReader(new FileReader("D:\\url.txt"));
		String url  = null;
		int i = 0;
		while ((url = bReader.readLine()) != null) {
			urlQueue.add(url);
			i++;
//			if (i == 20) {
//				break;
//			}
		}
		bReader.close();

		int threads = 10;
		List<Future> futures = new ArrayList<Future>(threads);
		ExecutorService pool = Executors.newFixedThreadPool(threads);
		for (int j = 0; j < threads; j++) {
			futures.add(pool.submit(new GetWebInfoThread(urlQueue, resultQueue)));
        }
		for(i = 0; i < futures.size(); i++){
			Object result = null;
			try {
				result = futures.get(i).get();
			} catch (InterruptedException e) {
				e.printStackTrace();
			} catch (ExecutionException e) {
				e.printStackTrace();
			}
        	logger.info("future info  "+i+":"+result);
		}
		pool.shutdown();
		
		bWriter = new BufferedWriter(new FileWriter("D:\\result.tsv"));
		for (String line:resultQueue) {
			bWriter.write(line);
			bWriter.newLine();
		}
		
		bWriter.flush();
		bWriter.close();
	}

}

class GetWebInfoThread implements Callable {
	private static Logger logger = Logger.getLogger(GetWebInfoThread.class);
	private WebDriver driver = null;
	private ConcurrentLinkedQueue<String> urlQueue = null;
	private ConcurrentLinkedQueue<String> resultQueue = null;
	public GetWebInfoThread(ConcurrentLinkedQueue<String> urlQueue, ConcurrentLinkedQueue<String> resultQueue) {
		DesiredCapabilities cap = DesiredCapabilities.phantomjs();
		cap.setCapability("phantomjs.page.settings.userAgent","Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.90 Safari/537.36");
		driver = new PhantomJSDriver(cap);
		driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
		this.urlQueue = urlQueue;
		this.resultQueue = resultQueue;
	}
	
	public String call() {
		while (true) {
			String url = urlQueue.poll();
			if (url == null) {
				break;
			}
			matchHtml(url);
			logger.info("url=" + url);
		}
//		String threadName = Thread.currentThread().getName();
//		logger.info("threadName=" + threadName + ", resultQueue size=" + resultQueue.size());
		driver.close();
		driver.quit();
		return Thread.currentThread().getName();
	}
	
	private void matchHtml(String url) {
		driver.get(url);
		//	  java.io.File screenShotFile = ((TakesScreenshot) driver)
		//			  .getScreenshotAs(OutputType.FILE);
		//	  FileUtils.copyFile(screenShotFile, new java.io.File("D:\\a2.png"));
		String pageSource = driver.getPageSource();
		Document doc = Jsoup.parse(pageSource);
		List<String> infoArray = new ArrayList<String>();

		Elements el = doc.select("span[class=j-info]");;// title
		if (el.size() > 0) {
			infoArray.add(el.get(0).text());
		} else {
			infoArray.add("no title");
		}
		//
		el = doc.select("div[class=discountArea f-pa j-discountArea f-cb] p");// price
		if (el.size() > 0) {
			infoArray.add(el.get(0).text());
		} else {
			infoArray.add("no price");
		}
		//
		el = doc.select("div[class=j-info f-cb] div:eq(1)");// teacher name
		if (el.size() > 0) {
			infoArray.add(el.get(0).text());
		} else {
			infoArray.add("no teacher name");
		}
		//
		//
		//      el = doc.select("span[class=renqi]");// teacher_rating
		//      if (el.size() > 0) {
		//             infoArray.add(el.get(0).text());
		//      } else 
		//        {
		infoArray.add("no teacher rating");
		//        }
		//
		//      el = doc.select("div[class=ltxt j-ltxt f-richEditorText edueditor_styleclass_0 edueditor_styleclass_1]");// teacher desc
		//      if (el.size() > 0) {
		//             infoArray.add(el.get(0).text());
		//      } else 
		//        {
		infoArray.add("no teacher desc");
		//        }
		//
		el = doc.select("p[class=j-info]:contains(分类：)");// category path
		if (el.size() > 0) {
			infoArray.add(el.get(0).text());
		} else
		{
			infoArray.add("no category path");
		}
		//
		el = doc.select("div[class=g-sd1 left j-chimg] img");// picture
		if (el.size() > 0) {
			infoArray.add(el.get(0).attr("src"));
		} else {
			infoArray.add("no picture");
		}
		//
		//      el = doc.select("div[class=timeCon]");// start date
		//      if (el.size() > 0) {
		//             infoArray.add(el.get(0).text().split("|")[0]);
		//      } else 
		{
			infoArray.add("No start date");// start date
		}
		//
		//
		el = doc.select("span[class=f-db f-pa valid]");//end date
		if(el.size()>0)
		{
			infoArray.add(el.get(0).text());
		}
		else
		{
			infoArray.add("no end date");
		}
		//
		//      el = doc.select("div[class=picCon] p");//video length same
		//      if(el.size()>0)
		//      {
		//      infoArray.add(el.get(0).text());
		//      }
		//      else
		{
			infoArray.add("no video length");
		}
		//
		el = doc.select("span[class=f-fl f-thide ks]");//course hour
		if(el.size()>0)
		{
			infoArray.add(String.valueOf(el.size()));
		}
		else
		{
			infoArray.add("no course hour");
		}
		//
		//
		// el = doc.select("span:contains(担保期) em");//expire date
		// if(el.size()>0)
		// {
		// infoArray.add(el.get(0).text());
		// }
		// else
		{
			infoArray.add("no expire date");
		}
		//
		infoArray.add("no 3rd platform");// no 3rd platform
		//
		el = doc.select("p[class=j-info f-thide] a");//school url
		if(el.size()>0)
		{
			infoArray.add(el.get(0).attr("href"));
		}
		else
		{
			infoArray.add("no school URL");
		}
		//
		//
		//      el = doc.select("p[class=lname f-thide j-lname]");//school name
		if(el.size()>0)
		{
			infoArray.add(el.get(0).text());
		}
		else
		{
			infoArray.add("no school name");
		}
		//
		//
		// el = doc.select("div[class=tb-property-cont course-time]");//type
		// if(el.size()>0)
		// {
		// infoArray.add(el.get(0).text());
		// }
		// else
		{
			infoArray.add("no type");
		}
		//
		el = doc.select("span[class=j-info starall] div[class=star on]");//rate
		if(el.size()>0)
		{
			infoArray.add(el.size() + "");
		}
		else
		{
			infoArray.add("no rate");
		}
		//
		el = doc.select("span[class=cmt j-cmt]");//comments count 
		if(el.size()>0)
		{
			infoArray.add(el.get(0).text());
		}
		else
		{
			infoArray.add("no comments_count");
		}
		//
		el = doc.select("em[class=num j-num f-fl]");//enrolled_count
		String enrolled_count = "";
		if(el.size()>0)
		{
			enrolled_count = el.get(0).text();
			infoArray.add(enrolled_count);
		}
		else
		{
			infoArray.add("no enrolled_count");
		}
		// detailContentLeft
		el = doc.select("div[class=cintrocon j-courseintro]");//desc with outline
		if(el.size()>0)
		{
			infoArray.add(el.get(0).text());
		}
		else
		{
			infoArray.add("no desc");// desc ajax by
		}                                        

		el = doc.select("div[class=m-chapterList f-pr]");// outline
		if (el.size() > 0) {
			infoArray.add(el.get(0).text());
		} 
		else 
		{
			infoArray.add("no outline");
		}

		infoArray.add(driver.getCurrentUrl());// video URL
		//
		//      el = doc.select("dl[class=CourseTabIntro]");// intro
		//      if (el.size() > 0) {
		//             infoArray.add(el.get(0).text());
		//      } 
		//      else
		{
			infoArray.add("no intro");
		}

		//      el = doc.select("ul[class=CRL_Students] li");// purchased_count
		if (enrolled_count != null && enrolled_count.length() > 0) {
			infoArray.add(enrolled_count);
		} else 
		{
			infoArray.add("no purchased_count");
		}
		String line = "";
		for(String str : infoArray)
		{

			line +=str+ "\t";
		}
		resultQueue.add(line);
	}
}
