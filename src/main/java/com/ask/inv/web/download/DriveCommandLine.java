package com.ask.inv.web.download;

import com.ask.inv.Util.ComUtil;
import com.ask.inv.Util.GoogleDriveUtil;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.auth.oauth2.GoogleTokenResponse;
import com.google.api.client.http.FileContent;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.File;
import com.google.api.services.drive.model.FileList;
import com.google.api.services.drive.model.ParentReference;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
//import org.openqa.selenium.
import org.openqa.selenium.htmlunit.HtmlUnitDriver;
import org.openqa.selenium.phantomjs.PhantomJSDriver;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class DriveCommandLine {
	private static Logger logger = Logger.getLogger(DriveCommandLine.class);

	private static String CLIENT_ID = "35851040803-igijign0n6mvmumfca3nfser24fkunvp.apps.googleusercontent.com";
	private static String CLIENT_SECRET = "1tIZf-RCZNGMbtWISrWMlP6Q";

	private static String REDIRECT_URI = "http://www.investopedia.com/oauth2callback";
	private WebDriver driver = null;

	private String name = "";
	private String pwd = "";

	private String name_Id = "Email";
	private String pwd_Id = "Passwd";
	private String signIn = "signIn";
	private String final_id = "submit_approve_access";
	private String phantomJsPath = "";
	private String fileTitle = "Master.twbx";
	private GoogleDriveUtil gDrive;
	private String filePath = "";

	public DriveCommandLine(String _filePath, String _fileTitle) {
		gDrive = new GoogleDriveUtil(
				"35851040803-bt59pos1ge5bhb88f3051iqt018pnp7q@developer.gserviceaccount.com",
				"./PK12.p12");
		this.filePath = _filePath;
		this.fileTitle = _fileTitle;
	}

	// public void initFolderAndPermission(String sharedEmailAddress) {
	// File file = gDrive.createFolder("CORE","CORE shared");
	// if (file != null) {
	// String[] emailAddress = sharedEmailAddress.split(",");
	// for (String email:emailAddress) {
	// gDrive.insertPermission(file.getId(), email, "user", "writer");
	// logger.info("shared File to " + email + " success.");
	// }
	// }
	//
	// gDrive.getFileList(null);
	// }
	public void uploadFile() {
		if (this.ifUpdate()) {
			File file = null;
			String parentFolderId = "0B-egulo89R22OVRESE1oZTVhX3c";
			java.io.File srcfile = new java.io.File(filePath);

			if (!srcfile.exists()) {
				logger.error("The upload file not exist!");
			}
			try {
				List<String> deleteFiles = this.getMasterFile();
				for (String fileName : deleteFiles) {
					System.out.println("Delete file: " + fileName);
					this.gDrive.trashFile(fileName);
					// service.files().trash(fileName).execute();
				}
				// Insert a file
				File body = new File();
				

				SimpleDateFormat sdfName = new SimpleDateFormat("yyyy-MM-dd");
				SimpleDateFormat sdf = new SimpleDateFormat(
						"yyyy-MM-dd HH:mm:ss");
				String dataStrDesc = sdf.format(new Date());
				String dataStrName = sdfName.format(new Date());
				body.setTitle(this.fileTitle);
				body.setDescription("update at: " + dataStrDesc + " | "
						+ System.currentTimeMillis());
				body.setMimeType("application/x-twb");
				// Set the parent folder.
				if (parentFolderId != null && parentFolderId.length() > 0) {
					body.setParents(Arrays.asList(new ParentReference()
							.setId(parentFolderId)));
				}

				FileContent mediaContent = new FileContent("application/x-twb",
						srcfile);

				file = gDrive.getDrive().files().insert(body, mediaContent)
						.execute();
				System.out.println("uploadFile success. filename="
						+ file.getTitle() + ", fileId=" + file.getId());
			} catch (IOException e) {
				e.printStackTrace();
				System.out.println("uploadFile have an error occurred: "
						+ e.getMessage());
			}

		}
	}

	private boolean ifUpdate() {
		boolean result = true;

		List<File> fileList = this.gDrive.getFileList(null);
		if (fileList != null && fileList.size() > 0) {
			for (File file : fileList) {
				// System.out.println(file.getTitle() + " | " +
				// file.getExplicitlyTrashed() + " |" + file.getDownloadUrl());
				if (file.getTitle().trim().toLowerCase().equals("master.twbx")) {
					Boolean ifDelete = file.getExplicitlyTrashed();
					if (ifDelete == null) {
						long timeStamp = file.getCreatedDate().getValue();
						if (System.currentTimeMillis() - timeStamp > 12 * 3600 * 1000) {
							result = true;
						} else {
							System.out.println("Get time stamp: "
									+ new Date(timeStamp).toLocaleString()
									+ " Current: "
									+ new Date(System.currentTimeMillis())
											.toLocaleString() + "gap: "
									+ (System.currentTimeMillis() - timeStamp)
									+ " < " + 12 * 3600 * 1000
									+ " skip upload!");
							result = false;
						}
						// System.out.println("create time: "
						// + file.getCreatedDate().getValue() + " | " +
						// System.currentTimeMillis());
					}

				}

			}
		}

		return result;
	}

	private ArrayList<String> getMasterFile() {
		ArrayList<String> files = new ArrayList<String>();
		List<File> fileList = this.gDrive.getFileList(null);
		if (fileList != null && fileList.size() > 0) {
			for (File file : fileList) {

				if (file.getTitle().trim().toLowerCase().equals("master.twbx")) {
					Boolean ifDelete = file.getExplicitlyTrashed();

					if (ifDelete == null) {

						files.add(file.getId());
					}
				}

			}
		}

		return files;
	}

	// public void a()
	// {
	// gDrive.getFileList("CORE");
	// }
	public static void main(String[] args) {

		new DriveCommandLine("./Master_tmp.twbx", "Master.twbx").uploadFile();
	}

	// 0B-egulo89R22OVRESE1oZTVhX3c

	/*
	 * public static void main(String[] args) throws IOException,
	 * InterruptedException { DriveCommandLine driveCommandLine = new
	 * DriveCommandLine( "C:\\phantomjs\\phantomjs.exe", "Master.twbx","","");
	 * driveCommandLine.uploadFile("./Master_tmp.twbx"); }
	 * 
	 * public void uploadFile(String _fileName) throws IOException,
	 * InterruptedException { HttpTransport httpTransport = new
	 * NetHttpTransport(); JsonFactory jsonFactory = new JacksonFactory();
	 * 
	 * GoogleAuthorizationCodeFlow flow = new
	 * GoogleAuthorizationCodeFlow.Builder( httpTransport, jsonFactory,
	 * CLIENT_ID, CLIENT_SECRET,
	 * Arrays.asList(DriveScopes.DRIVE)).setAccessType("online")
	 * .setApprovalPrompt("auto").build();
	 * 
	 * String url = flow.newAuthorizationUrl().setRedirectUri(REDIRECT_URI)
	 * .build(); System.out .println(
	 * "Please open the following URL in your browser then type the authorization code:"
	 * ); System.out.println("  " + url);
	 * 
	 * System.setProperty("phantomjs.binary.path", phantomJsPath); WebDriver
	 * driver = new PhantomJSDriver(); String currentUrl = ""; try {
	 * driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
	 * driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
	 * driver.get(url); java.io.File screenShotFile = ((TakesScreenshot) driver)
	 * .getScreenshotAs(OutputType.FILE); FileUtils.copyFile(screenShotFile, new
	 * java.io.File("./a1.png"));
	 * driver.findElement(By.id(name_Id)).sendKeys(name);
	 * driver.findElement(By.id(pwd_Id)).sendKeys(pwd);
	 * driver.findElement(By.id(signIn)).submit(); // Thread.sleep(2000); //
	 * java.io.File screenShotFile1 = //
	 * ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE); //
	 * FileUtils.copyFile(screenShotFile1, new // java.io.File("./a2.png")); //
	 * System.out.println(2); // driver.findElement(By.id(final_id)).click();
	 * java.io.File screenShotFile2 = ((TakesScreenshot) driver)
	 * .getScreenshotAs(OutputType.FILE); FileUtils.copyFile(screenShotFile2,
	 * new java.io.File("./a3.png")); // System.out.println(3); //
	 * screenShotFile = //
	 * ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE); //
	 * FileUtils.copyFile(screenShotFile, new java.io.File("./a.png")); // new
	 * WebDriverWait(driver, 10).until(new // ExpectedCondition<Boolean>() { //
	 * 
	 * @Override // public Boolean apply(WebDriver d) { //
	 * System.out.println("invoker this method..."); // return //
	 * d.findElement(By
	 * .xpath("//body")).getText().contains("put your email name here"); // } //
	 * });
	 * 
	 * currentUrl = driver.getCurrentUrl(); //
	 * System.out.println(driver.getPageSource()); System.out.println("request="
	 * + currentUrl); } catch (Exception e) { e.printStackTrace(); } finally {
	 * driver.quit(); }
	 * 
	 * String code = null; if (currentUrl.indexOf("code=") != -1) { code =
	 * currentUrl.substring(currentUrl.indexOf("code=") + 5); } if
	 * (ComUtil.isEmpty(code)) { logger.error("Can't get authorization。");
	 * return; } GoogleTokenResponse response = flow.newTokenRequest(code)
	 * .setRedirectUri(REDIRECT_URI).execute(); GoogleCredential credential =
	 * new GoogleCredential() .setFromTokenResponse(response);
	 * 
	 * // Create a new authorized API client Drive service = new
	 * Drive.Builder(httpTransport, jsonFactory, credential).build(); //
	 * this.getFiles(service); if (true) { // if (this.ifUpdate(service)) { //
	 * Insert a file File body = new File(); SimpleDateFormat sdfName = new
	 * SimpleDateFormat("yyyy-MM-dd"); SimpleDateFormat sdf = new
	 * SimpleDateFormat("yyyy-MM-dd HH:mm:ss"); String dataStrDesc =
	 * sdf.format(new Date()); String dataStrName = sdfName.format(new Date());
	 * body.setTitle(this.fileTitle); body.setDescription("update at: " +
	 * dataStrDesc + " | " + System.currentTimeMillis());
	 * body.setMimeType("application/x-twb");
	 * 
	 * // Set the parent folder. String parentId =
	 * "0B-egulo89R22OVRESE1oZTVhX3c";
	 * 
	 * if (parentId != null && parentId.length() > 0) {
	 * body.setParents(Arrays.asList(new ParentReference() .setId(parentId))); }
	 * 
	 * java.io.File fileContent = new java.io.File(_fileName); FileContent
	 * mediaContent = new FileContent("application/x-twb", fileContent);
	 * 
	 * List<String> deleteFiles = this.getMasterFile(service); for (String
	 * fileName : deleteFiles) { System.out.println("Delete file: " + fileName);
	 * service.files().trash(fileName).execute(); }
	 * 
	 * File file = service.files().insert(body, mediaContent).execute();
	 * 
	 * System.out.println("File ID: " + file.getId()); } }
	 * 
	 * private boolean ifUpdate(Drive service) { boolean result = true;
	 * 
	 * try { List<File> fileList = retrieveAllFiles(service); if (fileList !=
	 * null && fileList.size() > 0) { for (File file : fileList) { //
	 * System.out.println(file.getTitle() + " | " + file.getExplicitlyTrashed()
	 * + " |" + file.getDownloadUrl()); if (file.getTitle().trim().toLowerCase()
	 * .equals("master.twbx")) { Boolean ifDelete = file.getExplicitlyTrashed();
	 * if (ifDelete == null) { long timeStamp =
	 * file.getCreatedDate().getValue(); if (System.currentTimeMillis() -
	 * timeStamp > 12 * 3600 * 1000) { result = true; } else {
	 * System.out.println("Get time stamp: " + new
	 * Date(timeStamp).toLocaleString() + " Current: " + new
	 * Date(System.currentTimeMillis()).toLocaleString() + "gap: "
	 * +(System.currentTimeMillis() - timeStamp)+ " < " + 12 * 3600 * 1000 +
	 * " skip upload!"); result = false; } // System.out.println("create time: "
	 * // + file.getCreatedDate().getValue() + " | " + //
	 * System.currentTimeMillis()); }
	 * 
	 * }
	 * 
	 * } }
	 * 
	 * } catch (IOException e) { // TODO Auto-generated catch block
	 * e.printStackTrace(); }
	 * 
	 * return result; }
	 * 
	 * private ArrayList<String> getMasterFile(Drive service) {
	 * ArrayList<String> files = new ArrayList<String>(); // Master.twbx try {
	 * List<File> fileList = retrieveAllFiles(service); if (fileList != null &&
	 * fileList.size() > 0) { for (File file : fileList) {
	 * 
	 * if (file.getTitle().trim().toLowerCase() .equals("master.twbx")) {
	 * Boolean ifDelete = file.getExplicitlyTrashed();
	 * 
	 * if (ifDelete == null) {
	 * 
	 * files.add(file.getId()); } }
	 * 
	 * } } } catch (IOException e) { e.printStackTrace();
	 * logger.error("getFiles have an error occurred: " + e.getMessage()); }
	 * return files; }
	 * 
	 * private void getFiles(Drive service) { try { List<File> fileList =
	 * retrieveAllFiles(service); if (fileList != null && fileList.size() > 0) {
	 * for (File file : fileList) { System.out.println("file title=" +
	 * file.getTitle() + ", file id=" + file.getId() + ", file time=" +
	 * file.getModifiedDate()); } } } catch (IOException e) {
	 * e.printStackTrace(); logger.error("getFiles have an error occurred: " +
	 * e.getMessage()); } }
	 * 
	 * /** Retrieve a list of File resources.
	 * 
	 * @return List of File resources.
	 * 
	 * @throws IOException
	 * 
	 * private List<File> retrieveAllFiles(Drive service) throws IOException {
	 * List<File> result = new ArrayList<File>();
	 * com.google.api.services.drive.Drive.Files.List request = service
	 * .files().list();
	 * 
	 * do { try { FileList files = request.execute();
	 * 
	 * result.addAll(files.getItems());
	 * request.setPageToken(files.getNextPageToken()); } catch (IOException e) {
	 * logger.error("retrieveAllFiles have an error occurred: " + e);
	 * request.setPageToken(null); } } while (request.getPageToken() != null &&
	 * request.getPageToken().length() > 0);
	 * 
	 * return result; }
	 */
}
