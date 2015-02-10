package com.ask.inv.google.analytics.report;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.security.GeneralSecurityException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import org.apache.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ask.inv.Util.ComUtil;
import com.ask.inv.Util.GoogleDriveUtil;
import com.ask.inv.Util.SendMailUtil;
import com.ask.inv.db.DBConnectionManager;
import com.ask.inv.db.NullConnectionException;
import com.ask.inv.google.analytics.model.GAMarketingData;
import com.ask.inv.google.analytics.model.GAReportInfo;
import com.ask.inv.google.analytics.model.UnsampledInfo;
import com.csvreader.CsvReader;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.json.GoogleJsonResponseException;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.Joiner;
import com.google.api.services.analytics.Analytics;
import com.google.api.services.analytics.AnalyticsScopes;
import com.google.api.services.analytics.model.GaData;
import com.google.api.services.analytics.model.GaData.ProfileInfo;
import com.google.api.services.analytics.model.GaData.Query;
import com.google.api.services.analytics.model.UnsampledReport;
import com.google.api.services.analytics.model.UnsampledReport.DriveDownloadDetails;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;

public class GAMarketingDataRequestReport {
	private static Logger logger = Logger.getLogger(GAMarketingDataRequestReport.class);
	
	private static DBConnectionManager dbm = DBConnectionManager.getInstance(GAMarketingDataRequestReport.class.getName());
	
	private Analytics service;
	private Drive drive;
	private String credentialP12KeyFilePath;
	private String serviceAccountEmail;
	private String outPutReportPath;
	private String dbconfigTagName = "";
	private String applicationName = "GAMarketingDataReqeustReport";
	private String profileId = "33962777";
	private DateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
	private GoogleDriveUtil driveUtil;
	private StringBuffer sqlStr = new StringBuffer();
	private Connection conn = null;
	private Date maxDbDate = null;
	private boolean mailSendFlag = false;
	private Properties prop = null;
	private Map<String, Integer> monthWeeksMap = new HashMap<String, Integer>();
	private String fileNameDate;
	
	public GAMarketingDataRequestReport() {
		prop = new Properties();
		try {
			prop.load(GAMarketingDataRequestReport.class.getClassLoader().getResourceAsStream("GAMarketingDataRequestReport.properties"));
		} catch (IOException e) {
			e.printStackTrace();
			logger.error("constructor have An error occurred: " + e.getMessage());
			prop = null;
		}
		if (prop == null) return;
		
		credentialP12KeyFilePath = prop.getProperty("credential.P12Key.filePath");
		logger.debug("credentialP12KeyFilePath=" + credentialP12KeyFilePath);
		serviceAccountEmail = prop.getProperty("service.account.email");
		logger.debug("serviceAccountEmail=" + serviceAccountEmail);
		outPutReportPath = prop.getProperty("output.report.path");
		logger.debug("outPutReportPath=" + outPutReportPath);
		dbconfigTagName = prop.getProperty("dbconfig.tag.name");
    	logger.debug("dbconfigTagName:" + dbconfigTagName);
    	mailSendFlag = Boolean.parseBoolean(prop.getProperty("mail.send.flag"));
    	logger.debug("mailSendFlag:" + mailSendFlag);
    	
		try {
			initAnalyticsService();
		} catch (GeneralSecurityException e) {
			e.printStackTrace();
			logger.error(e.getMessage());
		} catch (IOException e) {
			e.printStackTrace();
			logger.error(e.getMessage());
		}
	}
	
	private void initAnalyticsService() throws GeneralSecurityException, IOException {
		Set<String> scopes = new HashSet<String>();
		scopes.add(AnalyticsScopes.ANALYTICS); // You can set other scopes if needed
		scopes.add(DriveScopes.DRIVE);

		HttpTransport httpTransport = new NetHttpTransport();
		JacksonFactory jsonFactory = new JacksonFactory();
		GoogleCredential credential = new GoogleCredential.Builder()
				.setTransport(httpTransport)
				.setJsonFactory(jsonFactory)
				.setServiceAccountId(serviceAccountEmail)
				.setServiceAccountScopes(scopes)
				.setServiceAccountPrivateKeyFromP12File(
						new java.io.File(credentialP12KeyFilePath))
				.build();
		
		service = new Analytics.Builder(httpTransport, jsonFactory,
				null).setHttpRequestInitializer(credential)
				.setApplicationName(applicationName).build();
		drive = new Drive.Builder(httpTransport, jsonFactory, credential).build();
		driveUtil = new GoogleDriveUtil(drive);
	}
	
	private GaData executeDataQuery(GAMarketingData gaMarketingData, String dateStartStr, String dateEndStr,
			int pagesize,int retrytimes,int pagenum) {

		GaData gaData= null;
		boolean retry = true;
		int count = 0;
		while(retry){
			++count;
			try {
				gaData = service.data()
							.ga()
							.get("ga:" + profileId, // Table Id. ga: + profile id.
									dateStartStr, // Start date.
									dateEndStr, // End date.
									gaMarketingData.getMetrics())
							// Metrics.
							.setDimensions(gaMarketingData.getDimensions())
							.setSamplingLevel("HIGHER_PRECISION")
							.setFilters(gaMarketingData.getFilters())
							.setStartIndex(pagesize * pagenum + 1).setMaxResults(pagenum).execute();
			} catch (IOException e) {
				if(count <= retrytimes){
					logger.warn(e.getMessage()+",Reconnect:"+count);
					retry = true;
					continue;
				} else {
					logger.error(e.getMessage()+",Reconnect:"+count);
				}
			}
			retry = false;
		}
	
		return gaData;
	}
	
	private List<String[]> editRunDateList(String runDateStr) {
		List<String[]> runDateList = new ArrayList<String[]>();
    	SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
    	Calendar calendar = Calendar.getInstance();
		Calendar yearFirstDayCal = Calendar.getInstance();
		yearFirstDayCal.set(2014, 0, 1);
		Date yearFirstDay = yearFirstDayCal.getTime();
		
		// week to date
		List<String[]> dateList = new ArrayList<String[]>();		
		Date runDate = null;
		if (runDateStr != null && runDateStr.trim().length() > 0) {
			runDate = stringToDate(runDateStr);
		}
		if (runDate != null) {
			calendar.setTime(runDate);
		}
		calendar.set(Calendar.DAY_OF_WEEK, 1);
		calendar.add(Calendar.DAY_OF_YEAR, -1);
		
		Date date = calendar.getTime();
		int i = 0;
		int currentMonthWeeks = 0;
		Calendar calendarMonth = null;
		String[] dates = null;
		Calendar newCalendar = Calendar.getInstance();
		int previouYear = 0;
		while (yearFirstDay.compareTo(date) != 0) {
			dates = new String[2];
			calendar.set(Calendar.DAY_OF_WEEK, 7);
			newCalendar.setTime(calendar.getTime());
			date = calendar.getTime();
			if (yearFirstDay.after(date)) {
				date = yearFirstDay;
			}
			int year1 = calendar.get(Calendar.YEAR);
			dates[1] = format.format(date);
			
			if (i == 0) {
				calendarMonth = (Calendar)calendar.clone();
				calendarMonth.set(Calendar.DAY_OF_MONTH, 1);
				fileNameDate = dates[1];
			}
			if (previouYear != 0 && previouYear != year1 && !monthWeeksMap.containsKey(String.valueOf(previouYear))) {
				monthWeeksMap.put(String.valueOf(previouYear), currentMonthWeeks);
				calendarMonth = (Calendar)newCalendar.clone();
				calendarMonth.set(Calendar.DAY_OF_MONTH, 1);
				currentMonthWeeks = 0;
			}
			if (!calendarMonth.getTime().after(date)) {
				currentMonthWeeks++;
			}
			
			calendar.set(Calendar.DAY_OF_WEEK, 1);
			date = calendar.getTime();
			if (yearFirstDay.after(date)) {
				date = yearFirstDay;
			}
			int year0 = calendar.get(Calendar.YEAR);
			if (year0 != year1 && date.after(yearFirstDay)) {
				newCalendar.set(Calendar.DAY_OF_MONTH, 1);
				dates[0] = format.format(newCalendar.getTime());
				dateList.add(dates);
				monthWeeksMap.put(String.valueOf(year1), currentMonthWeeks);
				
				dates = new String[2];
				newCalendar.add(Calendar.DAY_OF_YEAR, -1);
				dates[1] = format.format(newCalendar.getTime());
				
				calendarMonth = (Calendar)newCalendar.clone();
				calendarMonth.set(Calendar.DAY_OF_MONTH, 1);
				currentMonthWeeks = 1;
			}
			
			dates[0] = format.format(date);
			dateList.add(dates);
			calendar.add(Calendar.DAY_OF_YEAR, -1);

			previouYear = year1;
			i++;
		}
		
		monthWeeksMap.put(String.valueOf(previouYear), currentMonthWeeks);
		
		for(i = dateList.size() - 1; i >= 0; i--) {
			dates = dateList.get(i);
			runDateList.add(dates);
		}
		
		return runDateList;
    }
	
	private List<String[]> editRunDateListForWholeYear(int year) {
		List<String[]> runDateList = new ArrayList<String[]>();
    	SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
    	Calendar calendar = Calendar.getInstance();
    	calendar.set(year+1, 0, 1);
    	calendar.add(Calendar.DAY_OF_YEAR, -1);
		Calendar yearFirstDayCal = (Calendar)calendar.clone();
		yearFirstDayCal.set(year, 0, 1);
		Date yearFirstDay = yearFirstDayCal.getTime();
		
		String[] dates = null;
		// week to date
		List<String[]> dateList = new ArrayList<String[]>();
		dates = new String[2];
		dates[1] = format.format(calendar.getTime());
		calendar.set(Calendar.DAY_OF_WEEK, 1);
		dates[0] = format.format(calendar.getTime());
		dateList.add(dates);
		int currentMonthWeeks = 1;
		calendar.add(Calendar.DAY_OF_YEAR, -1);
		
		Date date = calendar.getTime();
		int i = 0;
		Calendar calendarMonth = null;
		while (yearFirstDay.compareTo(date) != 0) {
			dates = new String[2];
			calendar.set(Calendar.DAY_OF_WEEK, 7);
			date = calendar.getTime();
			if (yearFirstDay.after(date)) {
				date = yearFirstDay;
			}
			dates[1] = format.format(date);
			
			if (i == 0) {
				calendarMonth = (Calendar)calendar.clone();
				calendarMonth.set(Calendar.DAY_OF_MONTH, 1);
				fileNameDate = dates[1];
			}
			if (!calendarMonth.getTime().after(date)) {
				currentMonthWeeks++;
			}
			
			calendar.set(Calendar.DAY_OF_WEEK, 1);
			date = calendar.getTime();
			if (yearFirstDay.after(date)) {
				date = yearFirstDay;
			}
			dates[0] = format.format(date);
			dateList.add(dates);
			
			calendar.add(Calendar.DAY_OF_YEAR, -1);
			i++;
		}
		
		for(i = dateList.size() - 1; i >= 0; i--) {
			dates = dateList.get(i);
			runDateList.add(dates);
		}
		monthWeeksMap.put(fileNameDate.substring(0, 4), currentMonthWeeks);
		return runDateList;
    }
	
	private Date stringToDate(String dateStr) {
		Date date = null;
		if (dateStr != null && dateStr.length() > 0) {
			try {
				date = fmt.parse(dateStr);
			} catch (ParseException e) {
				logger.error(e.getMessage());
				e.printStackTrace();
			}
		}
		
		return date;
	}
	
	private String dateStringformat(String dateStr) {
		String dateString = "";
		if (dateStr != null && dateStr.length() > 0) {
			try {
				Date date = fmt.parse(dateStr);
				SimpleDateFormat df = new SimpleDateFormat("MM.dd.yy");
				dateString = df.format(date);
			} catch (ParseException e) {
				logger.error(e.getMessage());
				e.printStackTrace();
			}
		}
		
		return dateString;
	}
	
	private Map<Integer, GAMarketingData> setGAMarketingDataParameter() {
		Map<Integer, GAMarketingData> gaMarketingDataMap = new HashMap<Integer, GAMarketingData>();
		
		if (sqlStr.length()>0) sqlStr.delete(0, sqlStr.length());
		sqlStr.append("SELECT id,dimensions,metrics,filters FROM ga_marketing_data_index_comment;");
    	
    	Statement stmt = null;
    	ResultSet rs = null;
        
    	try{
    		logger.debug("SQL:" + sqlStr.toString());
			stmt = getConnect().createStatement(java.sql.ResultSet.TYPE_FORWARD_ONLY,java.sql.ResultSet.CONCUR_READ_ONLY);
			stmt.setFetchSize(Integer.MIN_VALUE);
			stmt.setEscapeProcessing(false);
			rs = stmt.executeQuery(sqlStr.toString());
			GAMarketingData gaMarketingData = null;
			int i = 0;
			while (rs.next()){
				gaMarketingData = new GAMarketingData();
				gaMarketingData.setDimensions(rs.getString("dimensions"));
				gaMarketingData.setMetrics(rs.getString("metrics"));
				gaMarketingData.setFilters(rs.getString("filters"));
				//from Dimensions to Metrics, index start with 0
				gaMarketingData.setColumns();
				gaMarketingDataMap.put(rs.getInt("id"), gaMarketingData);
				i++;
			}
			logger.debug("setGAMarketingDataParameter count=" + i);
    	} catch (SQLException sqle){
			sqle.printStackTrace();
			logger.error(sqle.getMessage());
		} catch (Exception e){
			e.printStackTrace();
			logger.error(e.getMessage());
		} finally {
			try{
        		if (rs != null && !rs.isClosed()){
        			rs.close();
        			rs = null;
        		}
			} catch (SQLException sqle){
				sqle.printStackTrace();
			}
			try{
        		if (stmt != null && !stmt.isClosed()){
        			stmt.close();
        			stmt = null;
        		}
			} catch (SQLException sqle){
				sqle.printStackTrace();
			}
		}
		
		return gaMarketingDataMap;
	}
	
	public int outPutReport(String outPutFile, List<String[]> runDateList, int index) {
		Map<Integer, GAMarketingData> gaMarketingDataMap = setGAMarketingDataParameter();
		
		maxDbDate = stringToDate(getMaxDateFromDb());
		
		List<GAReportInfo> reportInfoList = getReportInfo(runDateList, gaMarketingDataMap, index);
		
		writeReport(outPutFile, "DASHBOARD", reportInfoList);
		
		// send mail
		if (mailSendFlag && reportInfoList != null && reportInfoList.size() > 0) {
			SendMailUtil mail = new SendMailUtil(prop);
			mail.sendBatchReport(outPutFile);
		}
		
		dbm.freeConnection(dbconfigTagName, conn);
		return 0;
	}
	
	private Connection getConnect() {
		try {
			if (conn == null || conn.isClosed()) {
				conn = dbm.getConnection(dbconfigTagName);
			}
		} catch (SQLException e1) {
			e1.printStackTrace();
			logger.error(e1.getMessage());
		} catch (NullConnectionException e){
        	logger.error(e.getMessage());
        	e.printStackTrace();
        }
    	return conn;
	}
	
	private List<GAReportInfo> getReportInfo(List<String[]> runDateList, Map<Integer, GAMarketingData> gaMarketingDataMap, int gaIndex) {
		List<GAReportInfo> reportInfoList = new ArrayList<GAReportInfo>();
		int pagenum = 10000;
		boolean checkDb = true;
		if (gaIndex != -1) {
			checkDb = false;
		}
		try {
			if (runDateList != null && runDateList.size() > 0) {
				for (String[] runDate:runDateList) {
					int days = (int)((stringToDate(runDate[1]).getTime() - stringToDate(runDate[0]).getTime())/(1000*3600*24) + 1);
					for (int index:gaMarketingDataMap.keySet()) {
						if (gaIndex == index || gaIndex == -1) {
							getGaData(gaMarketingDataMap.get(index), runDate, index, days, pagenum, checkDb);
						}
					}
					
					GAReportInfo gaReportInfo = new GAReportInfo();
					gaReportInfo.setDateFrom(runDate[0]);
					gaReportInfo.setDateTo(runDate[1]);
					
					// Yahoo Sessions
					// Yahoo Finance
					gaReportInfo.setYahooFinance(gaMarketingDataMap.get(1).getSessions());
					// Yahoo Other
					gaReportInfo.setYahooOther(gaMarketingDataMap.get(2).getSessions());
					// Yahoo Full Content
					gaReportInfo.setYahooFullContent(gaMarketingDataMap.get(3).getSessions());
					// Yahoo SA Quote Pages
					gaReportInfo.setYahooSAQuotePages(gaMarketingDataMap.get(4).getSessions());
					// Yahoo Repub
					gaReportInfo.setYahooRepub(gaMarketingDataMap.get(5).getSessions());
					
					// Yahoo PVs
					// Yahoo Finance
					gaReportInfo.setYahooFinancePvs(gaMarketingDataMap.get(1).getPageviews());
					// Yahoo Other
					gaReportInfo.setYahooOtherPvs(gaMarketingDataMap.get(2).getPageviews());
					// Yahoo Full Content
					gaReportInfo.setYahooFullContentPvs(gaMarketingDataMap.get(3).getPageviews());
					// Yahoo SA Quote Pages
					gaReportInfo.setYahooSAQuotePagesPvs(gaMarketingDataMap.get(4).getPageviews());
					// Yahoo Repub
					gaReportInfo.setYahooRepubPvs(gaMarketingDataMap.get(5).getPageviews());
					
					// Other Sessions
					// Nasdaq
					gaReportInfo.setNasdaq(gaMarketingDataMap.get(6).getSessions());
					// Forbes
					gaReportInfo.setForbes(gaMarketingDataMap.get(7).getSessions());
					// Fox Business
					gaReportInfo.setFoxBusiness(gaMarketingDataMap.get(8).getSessions());
					// MSN/NBC
					gaReportInfo.setMsnNbc(gaMarketingDataMap.get(9).getSessions());
					// Google News
					gaReportInfo.setGoogleNews(gaMarketingDataMap.get(10).getSessions());
					// Blackrock
					gaReportInfo.setBlackrock(gaMarketingDataMap.get(15).getSessions());
					// Ask
					gaReportInfo.setAsk(gaMarketingDataMap.get(16).getSessions());
					// Other
					gaReportInfo.setOther(gaMarketingDataMap.get(11).getSessions() + gaMarketingDataMap.get(12).getSessions());
					
					// Other PVs
					// Nasdaq
					gaReportInfo.setNasdaqPvs(gaMarketingDataMap.get(6).getPageviews());
					// Forbes
					gaReportInfo.setForbesPvs(gaMarketingDataMap.get(7).getPageviews());
					// Fox Business
					gaReportInfo.setFoxBusinessPvs(gaMarketingDataMap.get(8).getPageviews());
					// MSN/NBC
					gaReportInfo.setMsnNbcPvs(gaMarketingDataMap.get(9).getPageviews());
					// Google News
					gaReportInfo.setGoogleNewsPvs(gaMarketingDataMap.get(10).getPageviews());
					// Blackrock
					gaReportInfo.setBlackrockPvs(gaMarketingDataMap.get(15).getPageviews());
					// Ask
					gaReportInfo.setAskPvs(gaMarketingDataMap.get(16).getPageviews());
					// Other
					gaReportInfo.setOtherPvs(gaMarketingDataMap.get(11).getPageviews() + gaMarketingDataMap.get(12).getPageviews());
					
					// Social (All Networks)
					// All Visits -> sessions
					gaReportInfo.setAllVisits(gaMarketingDataMap.get(13).getSessions());
					// All PVs
					gaReportInfo.setAllPvs(gaMarketingDataMap.get(13).getPageviews());
					// All Posts -> sessions
					gaReportInfo.setAllPosts(gaMarketingDataMap.get(14).getSessions());
					// All PVs from Posting
					gaReportInfo.setAllPVsFromPosting(gaMarketingDataMap.get(14).getPageviews());
					
					reportInfoList.add(gaReportInfo);
				}
			}
		} catch (ParseException e) {
			e.printStackTrace();
			logger.error(e.getMessage());
		}
		
		return reportInfoList;
	}
	
	private void getGaData(GAMarketingData gaMarketingData,
			String[] runDate, int index, int days, int pagenum, boolean checkDb) throws ParseException {		
		// init
		// sessions
		gaMarketingData.setSessions(0);
		// pageviews
		gaMarketingData.setPageviews(0);
		
		String startDate = runDate[0];
		
		Calendar calendar = Calendar.getInstance();
		int times = 3;
		
		for(int past = 0 ; past < days; past++) {
			calendar.setTime(fmt.parse(startDate));
			boolean hasDbData = false;
			int preSessions = gaMarketingData.getSessions();
			int prePageviews = gaMarketingData.getPageviews();
			if (checkDb && maxDbDate != null && !maxDbDate.before(calendar.getTime())) {
				getGaDataFromDb(gaMarketingData, startDate, index);
				hasDbData = true;
			}
			
			int count = 0;
			GaData result = null;
			int insertCnt = 0;
			boolean unsampledFlag = false;
			
			String accountId = "";
			String webPropertyId = "";
			String unsampledReportId = "";
			int isSampledData = 0;
			
			while (!hasDbData) {
				logger.debug("Start fetch data: index=" + index + ", from " + startDate + " to " + startDate + " | " + count);
				result = executeDataQuery(gaMarketingData, startDate, startDate, count, times, pagenum);
				count++;
				if (result !=null) {
					List<List<String>> dataList = result.getRows();
					if (dataList == null) {
						logger.debug("break ==============================================="+ (result == null));
						break;
					}
					if (result.getContainsSampledData() && !unsampledFlag) {
						logger.debug("This result is based on sampled data. Contains Sampled Data :"+result.getContainsSampledData());
						isSampledData = 1;
						
						  // Use the “query” object to construct an unsampled report object.
						String title = GAMarketingDataRequestReport.class.getSimpleName() + "_" + index + "_" + startDate;
						  Query query = result.getQuery();
						  UnsampledReport report = new UnsampledReport()
						      .setDimensions(query.getDimensions())
						      .setMetrics(Joiner.on(',').join(query.getMetrics()))
						      .setStartDate(startDate)
						      .setEndDate(startDate)
						      .setSegment(query.getSegment())
						      .setFilters(query.getFilters())
						      .setTitle(title);
						  
						  ProfileInfo profileInfo = result.getProfileInfo();
						  
						  accountId = profileInfo.getAccountId();
						  webPropertyId = profileInfo.getWebPropertyId();
						try {
							// Generate reports
							UnsampledReport response = service.management().unsampledReports().insert(accountId, webPropertyId, profileId, report).execute();
							unsampledReportId = response.getId();
							logger.debug("Create Umsampled data success. accountId:"+accountId+",webPropertyId:"+webPropertyId
									+",profileId:"+profileInfo.getProfileId() + ", unsampledReportId=" + unsampledReportId + ", title=" + title);
							unsampledFlag = true;
						} catch (GoogleJsonResponseException e) {
							e.printStackTrace();
					          logger.error("There was a service error: "
								      + e.getDetails().getCode() + " : "
								      + e.getDetails().getMessage());
					          unsampledFlag = false;
						} catch (IOException e) {
							e.printStackTrace();
							logger.error(e.getMessage());
							unsampledFlag = false;
						}
					}
					
					insertCnt += writeGaData(gaMarketingData, dataList, "ga_marketing_data", startDate, index);
					if (insertCnt == -1) {
						logger.error("get ga data error(writeGaData error). index=" + index + ", from " + startDate + " to " + startDate + " | " + count);
					}
					logger.debug("get ga data success. rows=" + dataList.size() + ", insertCnt=" + insertCnt);
				} else {
					logger.debug("miss: from" + startDate + " to " + startDate + " | " + (count-1));
				}
			}
			
			if (insertCnt > 0) {
				int cnt = writeGaDataTotal(gaMarketingData.getSessions() - preSessions, gaMarketingData.getPageviews() - prePageviews,
						startDate, index, isSampledData, accountId, webPropertyId, unsampledReportId, 0);
				if (cnt == -1) {
					logger.error("get ga data error(writeGaDataTotal error). index=" + index + ", from " + startDate + " to " + startDate + " | " + count);
				}
			}
			
			calendar.add(Calendar.DATE, 1);	
			startDate = fmt.format(calendar.getTime());
		}
		
		logger.info("Date is from " + runDate[0] + " To " + runDate[1] + ", index=" + index + ", total sessions="
		+ gaMarketingData.getSessions() + ", total pageviews=" + gaMarketingData.getPageviews());
	}
	
	private void repairSampledData(String date, int index) {
		long startTime = System.currentTimeMillis();
		Map<Integer, GAMarketingData> gaMarketingDataMap = setGAMarketingDataParameter();
		List<UnsampledInfo> list = getUnsampledInfo(date, index);
		int generateUnsampledReportCnt = 0;
		int repairCount = 0;
		if (list.size() > 0) {
			for (UnsampledInfo info:list) {
				String unsampledReportId = info.getUnsampledReportId();
				int insertCnt = 0;
				GAMarketingData gaMarketingData = gaMarketingDataMap.get(info.getId());
				
				if (unsampledReportId != null && unsampledReportId.trim().length() > 0) {
					insertCnt = getUnsampledData(gaMarketingData, info.getAccountId(), info.getWebPropertyId(),
							unsampledReportId, info.getDate(), info.getId());
				} else {
					// Use the “query” object to construct an unsampled report object.
					String title = GAMarketingDataRequestReport.class.getSimpleName() + "_" + info.getId() + "_" + info.getDate();
					  UnsampledReport report = new UnsampledReport()
					      .setDimensions(gaMarketingData.getDimensions())
					      .setMetrics(gaMarketingData.getMetrics())
					      .setStartDate(info.getDate())
					      .setEndDate(info.getDate())
//					      .setSegment(query.getSegment())
					      .setFilters(gaMarketingData.getFilters())
					      .setTitle(title);
					try {
						// Generate reports
						UnsampledReport response = service.management().unsampledReports().insert(info.getAccountId(), info.getWebPropertyId(), profileId, report).execute();
						unsampledReportId = response.getId();
						logger.debug("Create Umsampled data success. accountId:"+info.getAccountId()+",webPropertyId:"+info.getWebPropertyId()
								+",profileId:"+profileId + ", unsampledReportId=" + unsampledReportId + ", title=" + title);
						int cnt = updateGaDataTotal(unsampledReportId, info.getDate(), info.getId());
						if (cnt == -1) {
							logger.error("update unsampledReportId error(updateGaDataTotal error). index=" + info.getId() + ", date=" + info.getDate() + ", unsampledReportId=" + unsampledReportId);
						}
						generateUnsampledReportCnt++;
						try {
							Thread.sleep(500);
						} catch (InterruptedException e) {
							e.printStackTrace();
						}
					} catch (GoogleJsonResponseException e) {
						e.printStackTrace();
				          logger.error("There was a service error: "
							      + e.getDetails().getCode() + " : "
							      + e.getDetails().getMessage());
				          break;
					} catch (IOException e) {
						e.printStackTrace();
						logger.error(e.getMessage());
						break;
					}
				}
				
				if (insertCnt == -1) {
					logger.error("write Unsampled Data error(getUnsampledData error). index=" + info.getId() + ", date=" + info.getDate());
				}
				
				if (insertCnt > 0) {
					repairCount++;
					GAMarketingData dbGaData = new GAMarketingData();
					getGaDataFromDb(dbGaData, info.getDate(), info.getId());
					if (gaMarketingData.getSessions() >= dbGaData.getSessions() && gaMarketingData.getPageviews() >= dbGaData.getPageviews()) {
						logger.debug("Unsampled Data update. source db:[sessions=" + dbGaData.getSessions() + ", pageviews=" + dbGaData.getPageviews() +
								"]. Report:[sessions=" + gaMarketingData.getSessions() + ", pageviews=" + gaMarketingData.getPageviews() + "]");
					} else {
						logger.warn("Unsampled Data Is not correct. source db:[sessions=" + dbGaData.getSessions() + ", pageviews=" + dbGaData.getPageviews() +
								"]. Report:[sessions=" + gaMarketingData.getSessions() + ", pageviews=" + gaMarketingData.getPageviews() + "]");
					}
					int cnt = updateGaDataTotal(gaMarketingData.getSessions(), gaMarketingData.getPageviews(), info.getDate(), info.getId(), 1);
					if (cnt == -1) {
						logger.error("update ga data error(updateGaDataTotal error). index=" + info.getId() + ", date=" + info.getDate());
					}
				}
			}
		}
		long endTime = System.currentTimeMillis();
		logger.info("repairSampledData result: repairCount=" + repairCount + ", generateUnsampledReportCnt=" + generateUnsampledReportCnt + ", used time=" + (endTime - startTime) + "ms");
	}
	
	private List<UnsampledInfo> getUnsampledInfo(String date, int index) {
		List<UnsampledInfo> list = new ArrayList<UnsampledInfo>();
		if (sqlStr.length()>0) sqlStr.delete(0, sqlStr.length());
		sqlStr.append("SELECT id,`date`,accountId,webPropertyId,unsampledReportId, 0 AS num FROM ga_marketing_data_total WHERE isSampledData = 1 AND unsampledSuccess = 0 AND unsampledReportId != ''");
		if (index > 0) {
			sqlStr.append(" and id = " + index);
		}
		if (date != null && date.length() > 0) {
			sqlStr.append(" and date = '" + date + "'");
		}
		sqlStr.append(" UNION ");
		sqlStr.append("SELECT id,`date`,accountId,webPropertyId,unsampledReportId, 1 AS num FROM ga_marketing_data_total WHERE isSampledData = 1 AND unsampledSuccess = 0 AND unsampledReportId = ''");
		if (index > 0) {
			sqlStr.append(" and id = " + index);
		}
		if (date != null && date.length() > 0) {
			sqlStr.append(" and date = '" + date + "'");
		}
		sqlStr.append(" ORDER BY num, `date`,id");
    	
    	Statement stmt = null;
    	ResultSet rs = null;
        
    	try{
    		logger.debug("SQL:" + sqlStr.toString());
			stmt = getConnect().createStatement(java.sql.ResultSet.TYPE_FORWARD_ONLY,java.sql.ResultSet.CONCUR_READ_ONLY);
			stmt.setFetchSize(Integer.MIN_VALUE);
			stmt.setEscapeProcessing(false);
			rs = stmt.executeQuery(sqlStr.toString());
			while (rs.next()){
				UnsampledInfo unsampledInfo = new UnsampledInfo();
				unsampledInfo.setId(rs.getInt("id"));
				unsampledInfo.setDate(rs.getString("date"));
				unsampledInfo.setAccountId(rs.getString("accountId"));
				unsampledInfo.setWebPropertyId(rs.getString("webPropertyId"));
				unsampledInfo.setUnsampledReportId(rs.getString("unsampledReportId"));
				list.add(unsampledInfo);
			}
			logger.debug("getUnsampledInfo get rows=" + list.size());
    	} catch (SQLException sqle){
			sqle.printStackTrace();
			logger.error(sqle.getMessage());
		} catch (Exception e){
			e.printStackTrace();
			logger.error(e.getMessage());
		} finally {
			try{
        		if (rs != null && !rs.isClosed()){
        			rs.close();
        			rs = null;
        		}
			} catch (SQLException sqle){
				sqle.printStackTrace();
			}
			try{
        		if (stmt != null && !stmt.isClosed()){
        			stmt.close();
        			stmt = null;
        		}
			} catch (SQLException sqle){
				sqle.printStackTrace();
			}
		}
    	
    	return list;
	}
	
	private int getUnsampledData(GAMarketingData gaMarketingData, String accountId, String webPropertyId,
			String unsampledReportId, String date, int index) {
		int cnt = 0;
		// get report
		UnsampledReport unsampledReport;
		
		// init
		// sessions
		gaMarketingData.setSessions(0);
		// pageviews
		gaMarketingData.setPageviews(0);
		
		try {
			unsampledReport = service.management().unsampledReports().get(accountId, webPropertyId, profileId, unsampledReportId).execute();
			DriveDownloadDetails drivedDetails = unsampledReport.getDriveDownloadDetails();
			if (drivedDetails != null) {
				String fileId = drivedDetails.getDocumentId();
				List<List<String>> dataList = readUnsampledReport(fileId);
				cnt = writeGaData(gaMarketingData, dataList, "ga_marketing_data_unsampled", date, index);
			}
		} catch (IOException e) {
			e.printStackTrace();
			logger.error(e.getMessage());
			return -1;
		}
		return cnt;
	}
	
	private List<List<String>> readUnsampledReport(String fileId) {
		List<List<String>> dataList = new ArrayList<List<String>>();
		InputStream inputStream = driveUtil.downloadFile(fileId);
		if (inputStream != null) {
			CsvReader csvReader = new CsvReader(inputStream, Charset.forName("UTF-8"));
			String [] value;
			try {
				int i = 0;
				while(csvReader.readRecord()) {
					value = csvReader.getValues();
					if (value.length > 1) {
						if (i > 0) {
							dataList.add(Arrays.asList(value));
						}
						i++;
					}
				}
			} catch (IOException e) {
				e.printStackTrace();
				logger.error(e.getMessage());
			}
		}
		
		return dataList;
	}
	
	private String getMaxDateFromDb() {
		String date = "";
		if (sqlStr.length()>0) sqlStr.delete(0, sqlStr.length());
		sqlStr.append("SELECT MAX(`date`) `date` FROM ga_marketing_data_total;");
    	
    	Statement stmt = null;
    	ResultSet rs = null;
        
    	try{
    		logger.debug("SQL:" + sqlStr.toString());
			stmt = getConnect().createStatement(java.sql.ResultSet.TYPE_FORWARD_ONLY,java.sql.ResultSet.CONCUR_READ_ONLY);
			stmt.setFetchSize(Integer.MIN_VALUE);
			stmt.setEscapeProcessing(false);
			rs = stmt.executeQuery(sqlStr.toString());
			while (rs.next()){
				date = rs.getString("date");
			}
			logger.debug("getMaxDateFromDb max Date=" + date);
    	} catch (SQLException sqle){
			sqle.printStackTrace();
			logger.error(sqle.getMessage());
		} catch (Exception e){
			e.printStackTrace();
			logger.error(e.getMessage());
		} finally {
			try{
        		if (rs != null && !rs.isClosed()){
        			rs.close();
        			rs = null;
        		}
			} catch (SQLException sqle){
				sqle.printStackTrace();
			}
			try{
        		if (stmt != null && !stmt.isClosed()){
        			stmt.close();
        			stmt = null;
        		}
			} catch (SQLException sqle){
				sqle.printStackTrace();
			}
		}
		
		return date;
	}
	
	private boolean getGaDataFromDb(GAMarketingData gaMarketingData, String startDate, int index) {
		boolean hasData = false;
		if (sqlStr.length()>0) sqlStr.delete(0, sqlStr.length());
		sqlStr.append("SELECT sessions,pageviews FROM ga_marketing_data_total WHERE id = " + index + " and date='" + startDate + "'");
    	
    	Statement stmt = null;
    	ResultSet rs = null;
        
    	try{
    		logger.debug("SQL:" + sqlStr.toString());
			stmt = getConnect().createStatement(java.sql.ResultSet.TYPE_FORWARD_ONLY,java.sql.ResultSet.CONCUR_READ_ONLY);
			stmt.setFetchSize(Integer.MIN_VALUE);
			stmt.setEscapeProcessing(false);
			rs = stmt.executeQuery(sqlStr.toString());
			while (rs.next()){
				gaMarketingData.setSessions(gaMarketingData.getSessions() + rs.getInt("sessions"));
				gaMarketingData.setPageviews(gaMarketingData.getPageviews() + rs.getInt("pageviews"));
				hasData = true;
			}
			logger.debug("getGaDataFromDb hasData=" + hasData + ", index= " + index + ", date=" + startDate + ", sessions=" + gaMarketingData.getSessions() + ", pageviews=" + gaMarketingData.getPageviews());
    	} catch (SQLException sqle){
			sqle.printStackTrace();
			logger.error(sqle.getMessage());
		} catch (Exception e){
			e.printStackTrace();
			logger.error(e.getMessage());
		} finally {
			try{
        		if (rs != null && !rs.isClosed()){
        			rs.close();
        			rs = null;
        		}
			} catch (SQLException sqle){
				sqle.printStackTrace();
			}
			try{
        		if (stmt != null && !stmt.isClosed()){
        			stmt.close();
        			stmt = null;
        		}
			} catch (SQLException sqle){
				sqle.printStackTrace();
			}
		}
		
		return hasData;
	}
	
	private int writeGaData(GAMarketingData gaMarketingData, List<List<String>> dataList, String tableName, String startDate, int index) {
		int j = 0;
		int insertCnt = 0;
		int dailyInsertStep = 300;
		String[] columnsArray = gaMarketingData.getColumnsArray();
		Statement stmt = null;
		try {
			stmt = getConnect().createStatement();
			if (dataList != null && dataList.size() > 0) {
				for (int i = 0; i < dataList.size(); i++) {
					List<String> row = dataList.get(i);
					
					// sessions
					gaMarketingData.setSessions(gaMarketingData.getSessions() + Integer.parseInt(row.get(gaMarketingData.getSessionsIndex())));
					// pageviews
					gaMarketingData.setPageviews(gaMarketingData.getPageviews() + Integer.parseInt(row.get(gaMarketingData.getPageviewsIndex())));
					
					if (j == 0) {
			    		if (sqlStr.length()>0) sqlStr.delete(0, sqlStr.length());
		    	    	sqlStr.append("INSERT INTO " + tableName + " (id,date," + gaMarketingData.getColumns() + ") values ");
			    	} else {
			    		sqlStr.append(",");
					}
					sqlStr.append("(" + index + ",'" + startDate + "'");
			    	for (int k = 0; k < columnsArray.length; k++) {
			    		String value = row.get(k);
			    		if (value == null) {
			    			value = "";
			    		}
			    		value = value.replace("\\", "\\\\").replace("'", "\\'");
			    		if (value.length() > 1024 &&
			    				("landingPagePath".equalsIgnoreCase(columnsArray[k]) || "PagePath".equalsIgnoreCase(columnsArray[k]))) {
			    			value = value.substring(0, 1024);
			    		}
			    		if (value.length() > 255 && !"landingPagePath".equalsIgnoreCase(columnsArray[k]) && !"PagePath".equalsIgnoreCase(columnsArray[k])) {
			    			value = value.substring(0, 255);
			    		}
			    		sqlStr.append(",'" + value + "'");
			    	}
			    	sqlStr.append(")");
			    	
			    	j++;
			    	
			    	if ((j == dailyInsertStep || i == dataList.size() - 1) && j > 0) {
			    		sqlStr.append(";");
			    		logger.debug("SQL:" + sqlStr.toString());
			    		try {
							insertCnt += stmt.executeUpdate(sqlStr.toString());
						} catch (SQLException e) {
							e.printStackTrace();
							logger.error(e.getMessage());
							return -1;
						}
			    		
			    		j = 0;
			    	}
				}
			}
		} catch (SQLException e) {
			e.printStackTrace();
			logger.error(e.getMessage());
			return -1;
		} finally {
			try{
        		if (stmt != null && !stmt.isClosed()){
        			stmt.close();
        			stmt = null;
        		}
			} catch (SQLException sqle){
				sqle.printStackTrace();
				logger.error(sqle.getMessage());
			}
		}
		
		return insertCnt;
	}
	
	private int writeGaDataTotal(int sessions, int pageviews, String startDate, int index, int isSampledData, String accountId, String webPropertyId, String unsampledReportId, int unsampledSuccess) {
		int cnt = 0;
		if (sqlStr.length()>0) sqlStr.delete(0, sqlStr.length());
    	sqlStr.append("INSERT INTO ga_marketing_data_total (id,date,sessions,pageviews,isSampledData,accountId,webPropertyId,unsampledReportId,unsampledSuccess) values ")
    	.append("(" + index + ",'" + startDate + "'," + sessions + "," + pageviews + "," + isSampledData + ",'" + accountId + "','" + webPropertyId + "','" + unsampledReportId + "'," + unsampledSuccess + ")");
    	
    	Statement stmt = null;
        
    	try{
    		logger.debug("SQL:" + sqlStr.toString());
			stmt = getConnect().createStatement();
			cnt = stmt.executeUpdate(sqlStr.toString());
			logger.debug("writeGaDataTotal count=" + cnt);
    	} catch (SQLException sqle){
			sqle.printStackTrace();
			logger.error(sqle.getMessage());
			return -1;
		} catch (Exception e){
			e.printStackTrace();
			logger.error(e.getMessage());
			return -1;
		} finally {
			try{
        		if (stmt != null && !stmt.isClosed()){
        			stmt.close();
        			stmt = null;
        		}
			} catch (SQLException sqle){
				sqle.printStackTrace();
				logger.error(sqle.getMessage());
			}
		}
		
		return cnt;
	}
	
	private int updateGaDataTotal(String unsampledReportId, String date, int index) {
		int cnt = 0;
		if (sqlStr.length()>0) sqlStr.delete(0, sqlStr.length());
    	sqlStr.append("UPDATE ga_marketing_data_total set unsampledReportId = '" + unsampledReportId + "' where id = " + index + " and date = '" + date + "'");
    	
    	Statement stmt = null;
        
    	try{
    		logger.debug("SQL:" + sqlStr.toString());
			stmt = getConnect().createStatement();
			cnt = stmt.executeUpdate(sqlStr.toString());
			logger.debug("updateGaDataTotal count=" + cnt + ", unsampledReportId=" + unsampledReportId);
    	} catch (SQLException sqle){
			sqle.printStackTrace();
			logger.error(sqle.getMessage());
			return -1;
		} catch (Exception e){
			e.printStackTrace();
			logger.error(e.getMessage());
			return -1;
		} finally {
			try{
        		if (stmt != null && !stmt.isClosed()){
        			stmt.close();
        			stmt = null;
        		}
			} catch (SQLException sqle){
				sqle.printStackTrace();
				logger.error(sqle.getMessage());
			}
		}
		
		return cnt;
	}
	
	private int updateGaDataTotal(int sessions, int pageviews, String date, int index, int unsampledSuccess) {
		int cnt = 0;
		if (sqlStr.length()>0) sqlStr.delete(0, sqlStr.length());
    	sqlStr.append("UPDATE ga_marketing_data_total set sessions = " + sessions + ", pageviews = " + pageviews + ", unsampledSuccess = 1 where id = " + index + " and date = '" + date + "'");
    	
    	Statement stmt = null;
        
    	try{
    		logger.debug("SQL:" + sqlStr.toString());
			stmt = getConnect().createStatement();
			cnt = stmt.executeUpdate(sqlStr.toString());
			logger.debug("updateGaDataTotal count=" + cnt);
    	} catch (SQLException sqle){
			sqle.printStackTrace();
			logger.error(sqle.getMessage());
			return -1;
		} catch (Exception e){
			e.printStackTrace();
			logger.error(e.getMessage());
			return -1;
		} finally {
			try{
        		if (stmt != null && !stmt.isClosed()){
        			stmt.close();
        			stmt = null;
        		}
			} catch (SQLException sqle){
				sqle.printStackTrace();
				logger.error(sqle.getMessage());
			}
		}
		
		return cnt;
	}
	
	private int writeReport(String outPutFile, String sheetName,
			List<GAReportInfo> reportInfoList) {
		if (reportInfoList.size() == 0) {
			logger.error("size Error: reportInfoList size=" + reportInfoList.size());
			return -1;
		}
		XSSFWorkbook xwb = new XSSFWorkbook();
		GAMarketingReportWorkbook gaReport = new GAMarketingReportWorkbook(xwb);
		// j is level cells
		short j = 0;
		XSSFSheet sheet = null;
		String previouYear = "";
		String year = "";
		for (int k = 0; k < reportInfoList.size(); k++) {
			GAReportInfo reportInfo = reportInfoList.get(k);
			String dateTo = reportInfo.getDateTo();
			year = dateTo.substring(0, 4);
			
			if (ComUtil.isNotEmpty(previouYear) && !previouYear.equals(year)) {
				// MTD
				j++;
				gaReport.addMTDColumn(sheet, j, monthWeeksMap.get(previouYear));
				
				// YTD
				j++;
				gaReport.addYTDColumn(sheet, j);
				
				// Notes
				j++;
				gaReport.addNotesColumn(sheet, j);
				
				// hidden rows and resize
				gaReport.hiddenRowsAndResize(sheet, j);
			}
			
			if (!previouYear.equals(year)) {
				j = 0;
				sheet = xwb.createSheet(sheetName + "(" + year + ")");
				
				gaReport.addHeaderColumn(sheet, j);
			}
			
			j++;
			gaReport.addDetailColumn(sheet, j, dateTo, reportInfo);
			
			previouYear = year;
		}
		
		if (ComUtil.isNotEmpty(previouYear)) {
			// MTD
			j++;
			gaReport.addMTDColumn(sheet, j, monthWeeksMap.get(previouYear));
			
			// YTD
			j++;
			gaReport.addYTDColumn(sheet, j);
			
			// Notes
			j++;
			gaReport.addNotesColumn(sheet, j);
			
			// hidden rows and resize
			gaReport.hiddenRowsAndResize(sheet, j);
		}
		
		OutputStream out;
		try {
			out = new FileOutputStream(outPutFile);
			xwb.write(out);
			if (out != null) {
				out.close();
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			logger.error(e.getMessage());
		} catch (IOException e) {
			e.printStackTrace();
			logger.error(e.getMessage());
		}
		return 0;
	}

	/**
	 * @param args[0] args[1] args[2]
	 * 1: (null)
	 * 2: $index
	 * 3: $runDate
	 * 4: $runDate $index
	 * 5: $dateFrom $dateTo
	 * 6: $dateFrom $dateTo $index
	 * 7: deleteExpiredFile $days
	 * 8: repairSampledData
	 * 9: repairSampledData $date or $index
	 * 10: repairSampledData $date $index (or $index $date)
	 * 11: getFileList (filename)
	 * 12: wholeYearReport $year
	 */
	public static void main(String[] args) {
		long startTime = System.currentTimeMillis();
		GAMarketingDataRequestReport gaReport = new GAMarketingDataRequestReport();
		List<String[]> runDateList = null;
		// wholeYearReport
		if (args != null && args.length > 1 && "wholeYearReport".equalsIgnoreCase(args[0])) {
			int year = Integer.parseInt(args[1]);
			if (year > 0) {
				runDateList = gaReport.editRunDateListForWholeYear(year);
				String fileName = year + " - Marketing";
				java.io.File filePath = new java.io.File(gaReport.outPutReportPath);
				if (!filePath.exists()) {
					filePath.mkdir();
				}
				String outPutFile = "";
				if (gaReport.outPutReportPath.endsWith("/")) {
					outPutFile = gaReport.outPutReportPath + fileName + ".xlsx";
				} else {
					outPutFile = gaReport.outPutReportPath + "/" + fileName + ".xlsx";
				}
				
				gaReport.outPutReport(outPutFile, runDateList, -1);
				
				long endTime = System.currentTimeMillis();
				logger.debug("GAMarketingDataRequestReport total used time:" + (endTime-startTime) + "ms");
			}
			
			return;
		}
		// getFileList
		if (args != null && args.length > 0 && "getFileList".equalsIgnoreCase(args[0])) {
			if (args.length > 1) {
				gaReport.driveUtil.getFileList(args[1]);
			} else {
				gaReport.driveUtil.getFileList(null);
			}
			
			return;
		}
		
		int index = -1;
		// repairSampledData
		if (args != null && args.length > 0) {
			String command = args[0];
			if ("repairSampledData".equalsIgnoreCase(command)) {
				String date = null;
				if (args.length == 2) {
					if (ComUtil.isDate(args[1], "yyyy-MM-dd")) {
						date = args[1];
					} else {
						try {
							index = Integer.parseInt(args[1]);
						} catch (Exception e) {
							logger.error("the parameter(args[1]) is not a day string(yyyy-MM-dd) or integer value!");
							return;
						}
					}
				} else if (args.length == 3) {
					if (ComUtil.isDate(args[1], "yyyy-MM-dd")) {
						date = args[1];
						try {
							index = Integer.parseInt(args[2]);
						} catch (Exception e) {
							logger.error("the parameter(args[2]) is not integer value!");
							return;
						}
					} else {
						try {
							index = Integer.parseInt(args[1]);
						} catch (Exception e) {
							logger.error("the parameter(args[1]) is not integer value!");
							return;
						}
						
						if (ComUtil.isDate(args[2], "yyyy-MM-dd")) {
							date = args[2];
						} else {
							logger.error("the parameter(args[2]) is not a day string(yyyy-MM-dd)!");
							return;
						}
					}
				}
				gaReport.repairSampledData(date, index);
				return;
			}
		}
		
		String dateFrom = null;
		String dateTo = null;
		String runDateStr = null;
		boolean singleRunflag = false;
		
		if (args != null && args.length == 1) {
			if (!ComUtil.isDate(args[0], "yyyy-MM-dd")) {
				try {
					index = Integer.parseInt(args[0]);
				} catch (Exception e) {
					logger.error("the parameter is not a day string(yyyy-MM-dd) or integer value!");
					return;
				}
				
			} else {
				runDateStr = args[0];
			}
		} else if(args != null && args.length == 2){
			dateFrom = args[0];
			dateTo = args[1];
			if ("deleteExpiredFile".equalsIgnoreCase(args[0])) {
				gaReport.driveUtil.deleteExpiredFile(GAMarketingDataRequestReport.class.getSimpleName(), Integer.parseInt(args[1]));
				return;
			}
			if (!ComUtil.isDate(dateTo, "yyyy-MM-dd")) {
				if (!ComUtil.isDate(dateFrom, "yyyy-MM-dd")) {
					logger.error("the dateFrom is not a day string(yyyy-MM-dd)!");
					return;
				} else {
					runDateStr = dateFrom;
					try {
						index = Integer.parseInt(args[1]);
					} catch (Exception e) {
						logger.error("the parameter is not a day string(yyyy-MM-dd) or integer value!");
						return;
					}
				}
			} else {
				singleRunflag = true;
			}
		} else if (args != null && args.length == 3) {
			dateFrom = args[0];
			dateTo = args[1];
			if (!ComUtil.isDate(dateFrom, "yyyy-MM-dd")) {
				logger.error("the dateFrom is not a day string(yyyy-MM-dd)!");
				return;
			}
			if (!ComUtil.isDate(dateTo, "yyyy-MM-dd")) {
				logger.error("the dateTo is not a day string(yyyy-MM-dd)!");
				return;
			}
			try {
				index = Integer.parseInt(args[2]);
			} catch (Exception e) {
				logger.error("the parameter is not a day string(yyyy-MM-dd) or integer value!");
				return;
			}
			singleRunflag = true;
		}
		
		if (singleRunflag) {
			String[] dates = new String[2];
			dates[0] = dateFrom;
			dates[1] = dateTo;
			runDateList = new ArrayList<String[]>();
			runDateList.add(dates);
			gaReport.monthWeeksMap.put(dateTo.substring(0, 4), 1);
		} else {
			runDateList = gaReport.editRunDateList(runDateStr);
		}
		
		String suffix = "";
		if (singleRunflag) {
			suffix = gaReport.dateStringformat(dateFrom) + "_" + gaReport.dateStringformat(dateTo);
		} else {
			if (gaReport.fileNameDate != null) {
				suffix = gaReport.dateStringformat(gaReport.fileNameDate);
			} else {
				suffix = ComUtil.dateTostring(new Date(), "MM.dd.yy");
			}
		}
		String fileName = suffix + " - Marketing";
		java.io.File filePath = new java.io.File(gaReport.outPutReportPath);
		if (!filePath.exists()) {
			filePath.mkdir();
		}
		String outPutFile = "";
		if (gaReport.outPutReportPath.endsWith("/")) {
			outPutFile = gaReport.outPutReportPath + fileName + ".xlsx";
		} else {
			outPutFile = gaReport.outPutReportPath + "/" + fileName + ".xlsx";
		}
		
		gaReport.outPutReport(outPutFile, runDateList, index);
		
		long endTime = System.currentTimeMillis();
		logger.debug("GAMarketingDataRequestReport total used time:" + (endTime-startTime) + "ms");
	}

}
