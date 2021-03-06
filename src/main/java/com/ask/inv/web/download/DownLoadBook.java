package com.ask.inv.web.download;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.concurrent.TimeUnit;
import org.apache.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.htmlunit.HtmlUnitDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import com.gargoylesoftware.htmlunit.ConfirmHandler;
import com.gargoylesoftware.htmlunit.Page;
import com.gargoylesoftware.htmlunit.WebClient;
import com.gargoylesoftware.htmlunit.WebResponse;
import com.gargoylesoftware.htmlunit.WebWindowEvent;
import com.gargoylesoftware.htmlunit.WebWindowListener;

public class DownLoadBook {
	private static Logger logger = Logger.getLogger(DownLoadBook.class);
	
	private WebDriver driver = null;
	
	// please fill in your name and pwd
	private String loginUrl = "";
	private String name = "";
	private String pwd = "";
	private String downLoadUrl = "";
	private String filePath = "";
	
	private String name_Id = "name";
	private String pwd_Id = "pw";
	private String signIn = "//button[@value = 'submit']";

	public DownLoadBook(String loginUrl, String name, String pwd, String downLoadUrl, String filePath) {
		this.loginUrl = loginUrl;
		this.name = name;
		this.pwd = pwd;
		this.downLoadUrl = downLoadUrl;
		this.filePath = filePath;
		driver = new CustomHtmlUnitDriver();
		driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
	
	}

	private void Login() {
		driver.get(loginUrl);
		driver.findElement(By.id(name_Id)).sendKeys(name);
		driver.findElement(By.id(pwd_Id)).sendKeys(pwd);
		driver.findElement(By.xpath(signIn)).submit();
		new WebDriverWait(driver, 10).until(ExpectedConditions
				.titleContains("Workbook"));
	}

	public void DownLoadfile() {
		Login();
		driver.get(downLoadUrl);

	}

	public static void main(String[] args) {
		DownLoadBook dlb = new DownLoadBook("http://tableau001.vcbrands.net/auth/","jhe","P@ssword7",
				"http://tableau001.vcbrands.net/workbooks/MonthsandPageviewstoROI?format=twb&errfmt=html",
				"./outPut/UploadDataToGoogleDrive/Months and Pageviews to ROI.twbx");
		dlb.DownLoadfile();
	}

	private class CustomHtmlUnitDriver extends HtmlUnitDriver {

		// This is the magic. Keep a reference to the client instance
		protected WebClient modifyWebClient(WebClient client) {

			ConfirmHandler okHandler = new ConfirmHandler() {
				public boolean handleConfirm(Page page, String message) {
					return true;
				}
			};
			client.setConfirmHandler(okHandler);

			client.addWebWindowListener(new WebWindowListener() {

				public void webWindowOpened(WebWindowEvent event) {

				}

				public void webWindowContentChanged(WebWindowEvent event) {

					WebResponse response = event.getWebWindow()
							.getEnclosedPage().getWebResponse();

					// Change or add conditions for content-types that you would
					// to like
					// receive like a file.
					if (response.getContentType().equals("application/x-twb")) {
						getFileResponse(response, filePath);
					}

				}

				public void webWindowClosed(WebWindowEvent event) {

				}
			});

			return client;
		}

	}

	private void getFileResponse(WebResponse response, String fileName) {

		InputStream inputStream = null;

		// write the inputStream to a FileOutputStream
		OutputStream outputStream = null;

		try {

			inputStream = response.getContentAsStream();

			// write the inputStream to a FileOutputStream
			outputStream = new FileOutputStream(new File(fileName));

			int read = 0;
			byte[] bytes = new byte[1024];

			while ((read = inputStream.read(bytes)) != -1) {
				outputStream.write(bytes, 0, read);
			}

			logger.info("get file Done!");

		} catch (IOException e) {
			e.printStackTrace();
			logger.error(e.getMessage());
		} finally {
			if (inputStream != null) {
				try {
					inputStream.close();
				} catch (IOException e) {
					e.printStackTrace();
					logger.error(e.getMessage());
				}
			}
			if (outputStream != null) {
				try {
					// outputStream.flush();
					outputStream.close();
				} catch (IOException e) {
					e.printStackTrace();
					logger.error(e.getMessage());
				}

			}
		}

	}

}
