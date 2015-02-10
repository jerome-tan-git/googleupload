package com.ask.inv.google.analytics.model;

public class UnsampledInfo {
	private int id;
	private String date;
	private String accountId;
	private String webPropertyId;
	private String unsampledReportId;
	public int getId() {
		return id;
	}
	public void setId(int id) {
		this.id = id;
	}
	public String getDate() {
		return date;
	}
	public void setDate(String date) {
		this.date = date;
	}
	public String getAccountId() {
		return accountId;
	}
	public void setAccountId(String accountId) {
		this.accountId = accountId;
	}
	public String getWebPropertyId() {
		return webPropertyId;
	}
	public void setWebPropertyId(String webPropertyId) {
		this.webPropertyId = webPropertyId;
	}
	public String getUnsampledReportId() {
		return unsampledReportId;
	}
	public void setUnsampledReportId(String unsampledReportId) {
		this.unsampledReportId = unsampledReportId;
	}
}
