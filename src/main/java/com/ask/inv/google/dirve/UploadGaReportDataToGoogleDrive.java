package com.ask.inv.google.dirve;

import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import org.apache.log4j.Logger;
import com.ask.inv.Util.GoogleDriveUtil;
import com.ask.inv.google.analytics.model.FileInfo;
import com.google.api.client.http.FileContent;
import com.google.api.services.drive.model.File;
import com.google.api.services.drive.model.ParentReference;

public class UploadGaReportDataToGoogleDrive {
	private static Logger logger = Logger.getLogger(UploadGaReportDataToGoogleDrive.class);
	private String parentFolderId = "0B24FtjAXhWsOT2QzSnNmekstTFU";
	private String oldFolderId = "0B24FtjAXhWsOLVBFU1FMa3hucHc";
	private GoogleDriveUtil gDrive;
	
	public UploadGaReportDataToGoogleDrive(String emailAddress, String credentialP12KeyFilePath) {			
		gDrive = new GoogleDriveUtil(emailAddress, credentialP12KeyFilePath);
		List<File> fileList = gDrive.retrieveAllFiles();
		if (fileList != null && fileList.size() > 0) {
			int i = 0;
			for (File file:fileList) {
				if (file != null && "Weekly Reporting".equals(file.getTitle())) {
					parentFolderId = file.getId();
					i++;
				}
				if (file != null && "Old".equals(file.getTitle())) {
					oldFolderId = file.getId();
					i++;
				}
				if (i == 2) {
					break;
				}
			}
		}
	}
	
	public void printFileList(String fileName) {
		gDrive.getFileList(fileName);
	}
	
	public File moveFile(String fileId) {
		return gDrive.moveFile(fileId, oldFolderId);
	}
	
	public List<FileInfo> getInputFiles() {
		List<FileInfo> fileList = new ArrayList<FileInfo>();
		List<File> allFileList = gDrive.retrieveAllFiles();
		if (allFileList != null && allFileList.size() > 0) {
			SimpleDateFormat df = new SimpleDateFormat("MM.dd.yyyy");
			for (File f:allFileList) {
				if (f.getTitle() != null && f.getTitle().endsWith(" - Marketing.xlsx")) {
					List<ParentReference> plist = f.getParents();
					if (plist != null && plist.size() > 0 && plist.get(0).getId().equals(parentFolderId)) {
						String dateStr = f.getTitle().replace(" - Marketing.xlsx", "");
						Date date = null;
						try {
							date = df.parse(dateStr);
						} catch (ParseException e) {
							e.printStackTrace();
						}
						if (date != null) {
							FileInfo fileInfo = new FileInfo();
							fileInfo.setId(f.getId());
							fileInfo.setTitle(f.getTitle());
							fileInfo.setFileNameDate(date);
							fileList.add(fileInfo);
							logger.info("file title=" + f.getTitle() + ", file id=" + f.getId() + ", file time=" + f.getModifiedDate());
						}
					}
				}
			}
		}
		if (fileList.size() == 0) {
			logger.warn("no file exits like this fileName.");
		} else {
			Collections.sort(fileList, new Comparator<FileInfo>() {
	            public int compare(FileInfo file1, FileInfo file2) {
	            	return file1.getFileNameDate().compareTo(file2.getFileNameDate()) * -1;
	            }
	        });
			
		}
		
		return fileList;
	}
	
	public FileInfo getInputFile() {
		List<FileInfo> fileList = getInputFiles();
		
		if (fileList != null && fileList.size() > 0) {
			return fileList.get(0);
		} else {
			return null;
		}
	}
	
	public InputStream downloadExcelFile(String fileId, String exportType) {
		return gDrive.downloadExcelFile(fileId, exportType);
	}
	
	public void initFolderAndPermission(String sharedEmailAddress) {
		File file = gDrive.createFolder("reporting","reporting");
		if (file != null) {
			String[] emailAddress = sharedEmailAddress.split(",");
		    for (String email:emailAddress) {
		    	gDrive.insertPermission(file.getId(), email, "user", "writer");
				logger.info("shared File to " + email + " success.");
		    }
		}

		gDrive.getFileList(null);
	}
	
	public File uploadExcelFile(String filePath) {
		File file = null;
		java.io.File srcfile = new java.io.File(filePath);

		if (!srcfile.exists()) {
			logger.error("The upload file not exist!");
			return null;
		}
		try {		    
			//Insert a file  
			File body = new File();
			body.setTitle(srcfile.getName());
			body.setDescription("");
			body.setMimeType("application/vnd.google-apps.spreadsheet");

			// Set the parent folder.
			if (parentFolderId != null && parentFolderId.length() > 0) {
				body.setParents(Arrays.asList(new ParentReference().setId(parentFolderId)));
			}

			FileContent mediaContent = new FileContent("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", srcfile);

			file = gDrive.getDrive().files().insert(body, mediaContent).execute();
			logger.info("uploadFile success. filename=" + file.getTitle() + ", fileId=" + file.getId());
		} catch (IOException e) {
			e.printStackTrace();
			logger.error("uploadFile have an error occurred: " + e.getMessage());
			return null;
		}
		return file;
	}

	/**
	 * @param args[0]:action=list,trash,delete,deleteExpiredFile
	 */
	public static void main(String[] args) {
		UploadGaReportDataToGoogleDrive upload = new UploadGaReportDataToGoogleDrive("143882641310-pdqqjspreku64borke74s0mbvu490lur@developer.gserviceaccount.com",
				"C:\\Page_WorkSpace\\InvJavaBatch\\src\\main\\resource\\My Project-1599f4404e3d(jack.he@askmedia.com).p12");
//		String filePath = "C:\\Page_WorkSpace\\InvJavaBatch\\outPut\\01.12.15 - Marketing(test uesed).xlsx";
		List<FileInfo> srcFileList = upload.getInputFiles();
		if (srcFileList != null && srcFileList.size() > 0) {
			for (FileInfo file:srcFileList) {
				System.out.println("title=" + file.getTitle() + ", fileId=" + file.getId());
			}
		}
		
//		File file = upload.uploadExcelFile(filePath);
//		System.out.println(file.getAlternateLink());
//		if (file != null) {
//			upload.moveFile(file.getId());
//		}		
	}

}
