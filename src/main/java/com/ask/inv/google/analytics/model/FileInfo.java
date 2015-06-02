package com.ask.inv.google.analytics.model;

import java.util.Date;

public class FileInfo {
	private String Id;
	private String title;
	private Date fileNameDate;
	public String getId() {
		return Id;
	}
	public void setId(String id) {
		Id = id;
	}
	public String getTitle() {
		return title;
	}
	public void setTitle(String title) {
		this.title = title;
	}
	public Date getFileNameDate() {
		return fileNameDate;
	}
	public void setFileNameDate(Date fileNameDate) {
		this.fileNameDate = fileNameDate;
	}
}
