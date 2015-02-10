package com.ask.inv.Util;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import org.apache.log4j.Logger;

import com.google.api.client.http.GenericUrl;
import com.google.api.client.http.HttpResponse;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.model.File;
import com.google.api.services.drive.model.FileList;
import com.google.api.services.drive.model.Permission;

public class GoogleDriveUtil {
	private static Logger logger = Logger.getLogger(GoogleDriveUtil.class);
	private Drive drive;

	public GoogleDriveUtil(Drive drive) {
		this.drive = drive;
	}
	
	/**
	 * Retrieve a list of File resources.
	 * @return List of File resources.
	 * @throws IOException
	 */
	public List<File> retrieveAllFiles() throws IOException {
		List<File> result = new ArrayList<File>();
		com.google.api.services.drive.Drive.Files.List request = drive.files().list();

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
	
	public void getFileList(String fileName) {
		try {
			List<File> fileList = retrieveAllFiles();
			int i = 0;
			if (fileList != null && fileList.size() > 0) {
				for (File file:fileList) {
					String title = file.getTitle();
					if (fileName == null || title != null && title.indexOf(fileName) != -1) {
						i++;
						logger.info("file title=" + file.getTitle() + ", file id=" + file.getId() + ", file time=" + file.getModifiedDate());
					}
					
				}
			}
			if (i == 0) {
				logger.info("no file exits like this fileName. fileName = " + fileName);
			}
		} catch (IOException e) {
			e.printStackTrace();
			logger.error("getFiles have an error occurred: " + e.getMessage());
		}
	}
	
	public void deleteExpiredFile(String fileName, int days) {
		Calendar calendar = Calendar.getInstance();
		calendar.add(Calendar.DAY_OF_YEAR, -days);
		try {
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
}
