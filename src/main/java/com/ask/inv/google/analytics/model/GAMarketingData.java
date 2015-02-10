package com.ask.inv.google.analytics.model;

public class GAMarketingData {
	private String dimensions;
	private String metrics;
	private String filters;
	private int sessionsIndex = 0;
	private int sessions = 0;
	private int pageviewsIndex =0;
	private int pageviews = 0;
	private String[] columnsArray;
	private String columns;
	public String getDimensions() {
		return dimensions;
	}
	public void setDimensions(String dimensions) {
		this.dimensions = dimensions;
	}
	public String getMetrics() {
		return metrics;
	}
	public void setMetrics(String metrics) {
		this.metrics = metrics;
	}
	public String getFilters() {
		return filters;
	}
	public void setFilters(String filters) {
		this.filters = filters;
	}
	public int getSessionsIndex() {
		return sessionsIndex;
	}
	public void setSessionsIndex(int sessionsIndex) {
		this.sessionsIndex = sessionsIndex;
	}
	public int getSessions() {
		return sessions;
	}
	public void setSessions(int sessions) {
		this.sessions = sessions;
	}
	public int getPageviewsIndex() {
		return pageviewsIndex;
	}
	public void setPageviewsIndex(int pageviewsIndex) {
		this.pageviewsIndex = pageviewsIndex;
	}
	public int getPageviews() {
		return pageviews;
	}
	public void setPageviews(int pageviews) {
		this.pageviews = pageviews;
	}
	public String[] getColumnsArray() {
		return columnsArray;
	}
	public void setColumnsArray(String[] columnsArray) {
		this.columnsArray = columnsArray;
	}
	public String getColumns() {
		return columns;
	}
	public void setColumns(String columns) {
		this.columns = columns;
	}
	public void setColumns() {
		this.columns = "";
		if (this.dimensions != null && this.dimensions.length() > 0) {
			this.columns += this.dimensions;
		}
		if (this.metrics != null && this.metrics.length() > 0) {
			this.columns += "," + this.metrics;
		}
		this.columns = this.columns.replace("ga:", "");
		this.columnsArray = this.columns.split(",");
		for (int i = 0; i < columnsArray.length; i++) {
			if ("sessions".equalsIgnoreCase(columnsArray[i])) {
				this.sessionsIndex = i;
			} else if ("pageviews".equalsIgnoreCase(columnsArray[i])) {
				this.pageviewsIndex = i;
			}
		}
	}
}
