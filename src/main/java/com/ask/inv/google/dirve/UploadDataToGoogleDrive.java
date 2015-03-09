package com.ask.inv.google.dirve;

import java.io.IOException;
import java.security.GeneralSecurityException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import org.apache.log4j.Logger;

import com.ask.inv.Util.ComUtil;
import com.ask.inv.Util.GoogleDriveUtil;
import com.ask.inv.web.download.DownLoadBook;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.FileContent;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.File;
import com.google.api.services.drive.model.FileList;
import com.google.api.services.drive.model.Permission;

public class UploadDataToGoogleDrive {
	private static Logger logger = Logger.getLogger(UploadDataToGoogleDrive.class);
	private Drive service;
	private String loginUrl = "";
	private String loginName = "";
	private String loginPwd = "";
	private String downLoadUrl = "";
	private String filePath = "";
	
	private String emailAddress = "";
	private String credentialP12KeyFilePath="";
	private String sharedEmailAddress = "";
	private String sharedType="user";
	private String sharedRole="reader";
	private GoogleDriveUtil gDrive;
	
	
	
	public UploadDataToGoogleDrive() {
		gDrive = new GoogleDriveUtil("35851040803-bt59pos1ge5bhb88f3051iqt018pnp7q@developer.gserviceaccount.com ", "./PK12.p12");
		Properties prop = new Properties();
		try {
			prop.load(UploadDataToGoogleDrive.class.getClassLoader().getResourceAsStream("UploadDataToGoogleDrive.properties"));
		} catch (IOException e) {
			e.printStackTrace();
			logger.error("constructor have An error occurred: " + e.getMessage());
			prop = null;
		}
		if (prop == null) return;
		
		loginUrl = prop.getProperty("downLoad.login.url");
		logger.debug("loginUrl=" + loginUrl);
		loginName = prop.getProperty("downLoad.login.name");
		logger.debug("loginName=" + loginName);
		loginPwd = prop.getProperty("downLoad.login.password");
		logger.debug("loginPwd=" + loginPwd);
		downLoadUrl = prop.getProperty("downLoad.url");
		logger.debug("downLoadUrl=" + downLoadUrl);
		filePath = prop.getProperty("downLoad.filePath");
		logger.debug("filePath=" + filePath);
		emailAddress = prop.getProperty("email.address");
		logger.debug("emailAddress=" + emailAddress);
		credentialP12KeyFilePath = prop.getProperty("credential.P12Key.filePath");
		logger.debug("credentialP12KeyFilePath=" + credentialP12KeyFilePath);
		sharedEmailAddress = prop.getProperty("shared.email.address");
		logger.debug("sharedEmailAddress=" + sharedEmailAddress);
		sharedType = prop.getProperty("shared.type");
		logger.debug("sharedType=" + sharedType);
		sharedRole = prop.getProperty("shared.role");
		logger.debug("sharedRole=" + sharedRole);
		
		init();
	}
	
	private void init() {
		java.io.File p12File = new java.io.File(credentialP12KeyFilePath);
		JsonFactory jsonFactory = JacksonFactory.getDefaultInstance();
		HttpTransport httpTransport = null;
		GoogleCredential credential = null;
		try {
			try {
				httpTransport = GoogleNetHttpTransport.newTrustedTransport();
				credential = new GoogleCredential.Builder()
			    .setTransport(httpTransport)
			    .setJsonFactory(jsonFactory)
			    .setServiceAccountId(emailAddress)
			    .setServiceAccountPrivateKeyFromP12File(p12File)
			    .setServiceAccountScopes(Collections.singleton(DriveScopes.DRIVE))
			    .build();
			} catch (IOException e) {
				e.printStackTrace();
			}
		} catch (GeneralSecurityException e) {
			e.printStackTrace();
			logger.error("init service have an error occurred: " + e.getMessage());
			return;
		}
		
		//Create a new authorized API client
	    service = new Drive.Builder(httpTransport, jsonFactory, credential).build();
	}
	
	public void downloadFileToDirve() throws Exception {
		DownLoadBook downLoadBook = new DownLoadBook(loginUrl, loginName, loginPwd, downLoadUrl, filePath);
		downLoadBook.DownLoadfile();
		uploadFile();
	}
	
	private void uploadFile() {
		java.io.File srcfile = new java.io.File(filePath);
		
	    String newFilePath = "";
	    if (!srcfile.exists()) {
	    	logger.error("The upload file not exist!");
	    	return;
	    } else {
	    	SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
	    	String path = srcfile.getAbsolutePath();
	    	int pointIndex = path.lastIndexOf(".");
	    	newFilePath = path.substring(0, pointIndex) + "_" + format.format(new Date()) + path.substring(pointIndex);
	    }
	    java.io.File fileContent = new java.io.File(newFilePath);
	    if (fileContent.exists()) {
	    	fileContent.delete();
	    }
	    if (!srcfile.renameTo(fileContent)) {
	    	logger.error("file rename error!");
	    	return;
	    }
		try {		    
		    //Insert a file  
		    File body = new File();
		    body.setTitle(fileContent.getName());
		    body.setDescription("workbooks file");
		    body.setMimeType("text/plain");
		    
		 // Set the parent folder.
//		    String parentId = "";
//		    if (parentId != null && parentId.length() > 0) {
//		      body.setParents(
//		          Arrays.asList(new ParentReference().setId(parentId)));
//		    }
		    
		    
		    FileContent mediaContent = new FileContent("text/plain", fileContent);

		    File file = service.files().insert(body, mediaContent).execute();
		    logger.info("uploadFile success. filename=" + file.getTitle() + ", fileId=" + file.getId());
		    String[] emailAddress = sharedEmailAddress.split(",");
		    for (String email:emailAddress) {
		    	insertPermission(service, file.getId(), email, sharedType, sharedRole);
				logger.info("shared File to " + email + " success.");
		    }
		} catch (IOException e) {
			e.printStackTrace();
			logger.error("uploadFile have an error occurred: " + e.getMessage());
		}
	}
	
	private void getFiles() {
		try {
			List<File> fileList = retrieveAllFiles();
			if (fileList != null && fileList.size() > 0) {
				for (File file:fileList) {
					logger.info("file title=" + file.getTitle() + ", file id=" + file.getId() + ", file time=" + file.getModifiedDate());
				}
			}
		} catch (IOException e) {
			e.printStackTrace();
			logger.error("getFiles have an error occurred: " + e.getMessage());
		}
	}
	
	/**
	 * Retrieve a list of File resources.
	 * @return List of File resources.
	 * @throws IOException
	 */
	private List<File> retrieveAllFiles() throws IOException {
		List<File> result = new ArrayList<File>();
		com.google.api.services.drive.Drive.Files.List request = service.files().list();

		do {
			try {
				FileList files = request.execute();

				result.addAll(files.getItems());
				request.setPageToken(files.getNextPageToken());
			} catch (IOException e) {
				logger.error("retrieveAllFiles have an error occurred: " + e);
				request.setPageToken(null);
			}
		} while (request.getPageToken() != null &&
				request.getPageToken().length() > 0);

		return result;
	}
	
	private void deleteExpiredFile(int days) {
		Calendar calendar = Calendar.getInstance();
		calendar.add(Calendar.DAY_OF_YEAR, -days);
		java.io.File srcfile = new java.io.File(filePath);
		String name = srcfile.getName();
		String fileName = name.substring(0, name.lastIndexOf("."));
		try {
			List<File> fileList = retrieveAllFiles();
			if (fileList != null && fileList.size() > 0) {
				for (File file:fileList) {
					String title = file.getTitle();
					if (file.getModifiedDate().getValue() < calendar.getTimeInMillis()
							&& title != null && title.indexOf(fileName) != -1) {
						deleteFile(file.getId());
						logger.info("deleteExpiredFile success. file title=" + file.getTitle() + ", file id=" + file.getId() + ", file time=" + file.getModifiedDate());
					}
				}
			}
		} catch (IOException e) {
			e.printStackTrace();
			logger.error("deleteExpiredFile have an error occurred: " + e.getMessage());
		}
	}
	
	/**
	   * Permanently delete a file, skipping the trash.
	   *
	   * @param service Drive API service instance.
	   * @param fileId ID of the file to delete.
	   */
	private void deleteFile(String fileId) {
		try {
			service.files().delete(fileId).execute();
			logger.info("deleteFile success. fileId =" + fileId);
		} catch (IOException e) {
			logger.error("deleteFile have an error occurred: " + e.getMessage());
		}
	}
	
	/**
	 * Move a file to the trash.
	 * Files moved to the trash still appear by default in results from the files.list method. To permanently remove a file, use files.delete.
	 * @param service
	 * @param fileId
	 * @return
	 */
	private File trashFile(String fileId) {
		try {
			File file = service.files().trash(fileId).execute();
			logger.info("trashFile success. fileId =" + fileId);
			return file;
		} catch (IOException e) {
			logger.error("trashFile have an error occurred: " + e.getMessage());
		}
		return null;
	}
	
	/**
	   * Insert a new permission.
	   *
	   * @param service Drive API service instance.
	   * @param fileId ID of the file to insert permission for.
	   * @param value User or group e-mail address, domain name or {@code null}
	                  "default" type.
	   * @param type The value "user", "group", "domain" or "default".
	   * @param role The value "owner", "writer" or "reader".
	   * @return The inserted permission if successful, {@code null} otherwise.
	   */
	private static Permission insertPermission(Drive service, String fileId,
			String value, String type, String role) {
		Permission newPermission = new Permission();

		newPermission.setValue(value);
		newPermission.setType(type);
		newPermission.setRole(role);
		try {
			return service.permissions().insert(fileId, newPermission).execute();
		} catch (IOException e) {
			logger.error("insertPermission have an error occurred: " + e.getMessage());
		}
		return null;
	}

	/**
	 * @param args[0]:action=list,trash,delete,deleteExpiredFile
	 * @throws Exception 
	 */
	public static void main(String[] args) throws Exception {
		UploadDataToGoogleDrive upload = new UploadDataToGoogleDrive();
		if (args != null && args.length > 0) {
			String action = args[0];
			if ("list".equalsIgnoreCase(action)) {
				upload.getFiles();
			} else if ("trash".equalsIgnoreCase(action)) {
				if (args.length > 1 && ComUtil.isNotEmpty(args[1])) {
					String fileId = args[1];
					upload.trashFile(fileId);
				} else {
					logger.error("no fileId.");
				}
			} else if ("delete".equalsIgnoreCase(action)) {
				if (args.length > 1 && ComUtil.isNotEmpty(args[1])) {
					String fileId = args[1];
					upload.deleteFile(fileId);
				} else {
					logger.error("no fileId.");
				}
			} else if ("deleteExpiredFile".equalsIgnoreCase(action)) {
				if (args.length > 1 && ComUtil.isNotEmpty(args[1])) {
					int days = Integer.parseInt(args[1]);
					upload.deleteExpiredFile(days);
				} else {
					logger.error("no days.");
				}
			}
		} else{
			upload.downloadFileToDirve();
		}
		
	}

}
