package com.ask.inv.Util;

import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.List;
import org.apache.log4j.Logger;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.GenericUrl;
import com.google.api.client.http.HttpResponse;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.File;
import com.google.api.services.drive.model.FileList;
import com.google.api.services.drive.model.ParentReference;
import com.google.api.services.drive.model.Permission;

public class GoogleDriveUtil {
	private static Logger logger = Logger.getLogger(GoogleDriveUtil.class);
	private Drive drive;
	
	public Drive getDrive() {
		return drive;
	}

	public void setDrive(Drive drive) {
		this.drive = drive;
	}
	
	public GoogleDriveUtil(Drive drive) {
		this.drive = drive;
	}

	public GoogleDriveUtil(String emailAddress, String credentialP12KeyFilePath) {
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
		}
		
		//Create a new authorized API client
		drive = new Drive.Builder(httpTransport, jsonFactory, credential).build();
	}
	
	/**
	 * Retrieve a list of File resources.
	 * @return List of File resources.
	 * @throws IOException
	 */
	public List<File> retrieveAllFiles(){
		List<File> result = new ArrayList<File>();
		com.google.api.services.drive.Drive.Files.List request;
		try {
			request = drive.files().list();
		} catch (IOException e1) {
			e1.printStackTrace();
			logger.error(e1.getMessage());
			return result;
		}

		do {
			try {
				FileList files = request.execute();

				result.addAll(files.getItems());
				request.setPageToken(files.getNextPageToken());
			} catch (IOException e) {
				e.printStackTrace();
				logger.error("retrieveAllFiles have an error occurred: " + e);
				request.setPageToken(null);
			}
		} while (request.getPageToken() != null &&
				request.getPageToken().length() > 0);

		return result;
	}
	
	public List<File> getFileList(String fileName) {
		List<File> fileList = new ArrayList<File>();
		List<File> allFileList = retrieveAllFiles();
		int i = 0;
		if (allFileList != null && allFileList.size() > 0) {
			for (File file:allFileList) {
				String title = file.getTitle();
				if (fileName == null || title != null && title.indexOf(fileName) != -1) {
					fileList.add(file);
					i++;
					logger.debug("file title=" + file.getTitle() + ", file id=" + file.getId() + ", file time=" + file.getModifiedDate());
				}
				
			}
		}
		if (i == 0) {
			logger.warn("no file exists like this fileName. fileName = " + fileName);
		}
		
		return fileList;
	}
	
	public File getFileById(String fileId) {
		File file = null;
		try {
			file = drive.files().get(fileId).execute();
			if (file == null) {
				logger.warn("no file exists equls this fileId. fileId = " + fileId);
			}
		} catch (IOException e) {
			e.printStackTrace();
			logger.error("getFileById have an error occurred: " + e.getMessage());
		}
		
		return file;
	}
	
	public File moveFile(String fileId, String targetFolderId) {
		File file = getFileById(fileId);
		if (file == null) {
			logger.error("The file is not exists. fileId=" + fileId);
			return null;
		}
		File targetFolder = getFileById(targetFolderId);
		if (targetFolder == null) {
			logger.error("The targetFolder is not exists. targetFolderId=" + targetFolderId);
			return null;
		}
		
		ParentReference newParent = new ParentReference();
		newParent.setSelfLink(targetFolder.getSelfLink());
		List<ParentReference> list = targetFolder.getParents();
		if (list != null && list.size() > 0) {
			newParent.setParentLink(list.get(0).getSelfLink());
		}
		newParent.setId(targetFolderId);
		newParent.setKind(targetFolder.getKind());
		newParent.setIsRoot(false);
		List<ParentReference> parentsList = new ArrayList<ParentReference>();
		parentsList.add(newParent);
		file.setParents(parentsList);
		try {
			File updatedFile = drive.files().update(fileId, file).execute();
			return updatedFile;
		} catch (IOException e) {
			e.printStackTrace();
			logger.error(e.getMessage());
			return null;
		}
	}
	
	public List<File> getFile(String fileName) {
		List<File> fileList = new ArrayList<File>();
		List<File> allFileList = retrieveAllFiles();
		if (allFileList != null && allFileList.size() > 0) {
			for (File f:allFileList) {
				if (fileName != null && fileName.equals(f.getTitle())) {
					fileList.add(f);
					logger.debug("file title=" + f.getTitle() + ", file id=" + f.getId() + ", file time=" + f.getModifiedDate());
				}
				
			}
		}
		if (fileList.size() == 0) {
			logger.warn("no file exits like this fileName. fileName = " + fileName);
		}
		
		return fileList;
	}
	
	public File createFolder(String folderName, String folderDescription) {
		File body = new File();
	    body.setTitle(folderName);
	    body.setDescription(folderDescription);
	    body.setMimeType("application/vnd.google-apps.folder");
	    try {
	    	File file = drive.files().insert(body).execute();
	    	logger.debug("createFolder success. filename=" + file.getTitle() + ", fileId=" + file.getId());
	    	return file;
		} catch (IOException e) {
			e.printStackTrace();
			logger.error(e.getMessage());
			return null;
		}
	}
	
	public void deleteExpiredFile(String fileName, int days) {
		Calendar calendar = Calendar.getInstance();
		calendar.add(Calendar.DAY_OF_YEAR, -days);
		List<File> fileList = retrieveAllFiles();
		if (fileList != null && fileList.size() > 0) {
			for (File file:fileList) {
				String title = file.getTitle();
				if (file.getModifiedDate().getValue() < calendar.getTimeInMillis()
						&& title != null && title.indexOf(fileName) != -1) {
					deleteFile(file.getId());
					logger.debug("deleteExpiredFile success. file title=" + file.getTitle() + ", file id=" + file.getId() + ", file time=" + file.getModifiedDate());
				}
			}
		}
	}
	
	/**
	   * Permanently delete a file, skipping the trash.
	   *
	   * @param service Drive API service instance.
	   * @param fileId ID of the file to delete.
	   */
	public void deleteFile(String fileId) {
		try {
			drive.files().delete(fileId).execute();
		} catch (IOException e) {
			e.printStackTrace();
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
	public File trashFile(String fileId) {
		try {
			File file = drive.files().trash(fileId).execute();
			return file;
		} catch (IOException e) {
			e.printStackTrace();
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
	public Permission insertPermission(String fileId,
			String value, String type, String role) {
		Permission newPermission = new Permission();

		newPermission.setValue(value);
		newPermission.setType(type);
		newPermission.setRole(role);
		try {
			return drive.permissions().insert(fileId, newPermission).execute();
		} catch (IOException e) {
			e.printStackTrace();
			logger.error("insertPermission have an error occurred: " + e.getMessage());
		}
		return null;
	}

	/**
	   * Download a file's content.
	   *
	   * @param fileId ID of the file to print metadata for.
	   * @return InputStream containing the file's content if successful,
	   *         {@code null} otherwise.
	   */
	public InputStream downloadFile(String fileId) {
		 File file = null;
		try {
			file = drive.files().get(fileId).execute();
			logger.debug("file title=" + file.getTitle() + ", file id=" + file.getId() + ", file time=" + file.getModifiedDate()
					+ ", Description=" + file.getDescription() + ", MIME type=" + file.getMimeType());
		} catch (IOException e1) {
			e1.printStackTrace();
			logger.error("downloadFile have an error occurred: " + e1.getMessage());
			return null;
		}
	    if (file != null && file.getDownloadUrl() != null && file.getDownloadUrl().length() > 0) {
	      try {
	        HttpResponse resp =
	        		drive.getRequestFactory().buildGetRequest(new GenericUrl(file.getDownloadUrl()))
	                .execute();
	        return resp.getContent();
	      } catch (IOException e) {
	        e.printStackTrace();
	        logger.error("downloadFile have an error occurred: " + e.getMessage());
	        return null;
	      }
	    } else {
	      // The file doesn't have any content stored on Drive.
	      return null;
	    }
	  }
	
	/**
	 * 
	 * @param fileId
	 * @param exportType:text/csv=csv,application/pdf=pdf,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet=xlsx
	 * @return
	 */
	public InputStream downloadExcelFile(String fileId, String exportType) {
		 File file = null;
		try {
			file = drive.files().get(fileId).execute();
			logger.debug("file title=" + file.getTitle() + ", file id=" + file.getId() + ", file time=" + file.getModifiedDate()
					+ ", Description=" + file.getDescription() + ", MIME type=" + file.getMimeType());
		} catch (IOException e1) {
			e1.printStackTrace();
			logger.error("downloadFile have an error occurred: " + e1.getMessage());
			return null;
		}
	    if (file != null && file.getExportLinks() != null && file.getExportLinks().size() > 0) {
	      try {
	        HttpResponse resp = drive.getRequestFactory().buildGetRequest(new GenericUrl(file.getExportLinks().get(exportType))).execute();
	        return resp.getContent();
	      } catch (IOException e) {
	        e.printStackTrace();
	        logger.error("downloadFile have an error occurred: " + e.getMessage());
	        return null;
	      }
	    } else {
	      // The file doesn't have any content stored on Drive.
	      return null;
	    }
	  }
}
