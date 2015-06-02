package com.ask.inv.google.analytics.report;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.security.GeneralSecurityException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.ask.inv.Util.ComExcelUtil;
import com.ask.inv.Util.ComUtil;
import com.ask.inv.Util.SendMailUtil;
import com.ask.inv.db.DBConnectionManager;
import com.ask.inv.db.NullConnectionException;
import com.ask.inv.google.analytics.model.FileInfo;
import com.ask.inv.google.analytics.model.GAMarketingData;
import com.ask.inv.google.analytics.model.GAReportInfo;
import com.ask.inv.google.dirve.UploadGaReportDataToGoogleDrive;
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
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.File;

public class GAReportForGoogleDrive {
	private static Logger logger = Logger.getLogger(GAReportForGoogleDrive.class);
	
	private static DBConnectionManager dbm = DBConnectionManager.getInstance(GAReportForGoogleDrive.class.getName());
	
	private Analytics service;
	private String credentialP12KeyFilePath;
	private String serviceAccountEmail;
	private String outPutReportPath;
	private String dbconfigTagName = "";
	private String applicationName = "GAReportForGoogleDrive";
	private String profileId = "33962777";
	private SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
	private StringBuffer sqlStr = new StringBuffer();
	private Connection conn = null;
	private Date maxDbDate = null;
	private boolean mailSendFlag = false;
	private Properties prop = null;
	private Map<String, Integer> monthWeeksMap = new HashMap<String, Integer>();
	private Map<String, Integer> yearWeeksMap = new HashMap<String, Integer>();
	private String fileNameDate;
	private String googleDriveEmailAddress = "";
	private String googleDriveCredentialP12KeyFilePath = "";
	private String googleDriveSharedEmailAddress = "";
	private boolean uploadFileFlag = false;
	private UploadGaReportDataToGoogleDrive upload;
	private int weeks = -1;
	
	public GAReportForGoogleDrive() {
		prop = new Properties();
		try {
			prop.load(GAReportForGoogleDrive.class.getClassLoader().getResourceAsStream("GAReportForGoogleDrive.properties"));
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
    	googleDriveEmailAddress = prop.getProperty("googleDrive.email.address");
    	logger.debug("googleDriveEmailAddress:" + googleDriveEmailAddress);
    	googleDriveCredentialP12KeyFilePath = prop.getProperty("googleDrive.credential.P12Key.filePath");
    	logger.debug("googleDriveCredentialP12KeyFilePath:" + googleDriveCredentialP12KeyFilePath);
    	googleDriveSharedEmailAddress = prop.getProperty("googleDrive.shared.email.address");
    	logger.debug("googleDriveSharedEmailAddress:" + googleDriveSharedEmailAddress);
    	uploadFileFlag = Boolean.parseBoolean(prop.getProperty("googleDrive.uploadFile.flag"));
    	logger.debug("uploadFileFlag:" + uploadFileFlag);
		try {
			initAnalyticsService();
		} catch (GeneralSecurityException e) {
			e.printStackTrace();
			logger.error(e.getMessage());
		} catch (IOException e) {
			e.printStackTrace();
			logger.error(e.getMessage());
		}
		
		upload = new UploadGaReportDataToGoogleDrive(googleDriveEmailAddress, googleDriveCredentialP12KeyFilePath);
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
		int currentyearWeeks = 0;
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
				Calendar thisCalendar = (Calendar)calendar.clone();
				thisCalendar.add(Calendar.DAY_OF_YEAR, 2);
				fileNameDate = format.format(thisCalendar.getTime());
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
			
			if (previouYear != 0 && previouYear != year1 && !yearWeeksMap.containsKey(String.valueOf(previouYear))) {
				yearWeeksMap.put(String.valueOf(previouYear), currentyearWeeks);
				currentyearWeeks = 0;
			}
			currentyearWeeks++;
			
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
				yearWeeksMap.put(String.valueOf(year1), currentyearWeeks);
				i++;
				
				dates = new String[2];
				newCalendar.add(Calendar.DAY_OF_YEAR, -1);
				dates[1] = format.format(newCalendar.getTime());
				
				calendarMonth = (Calendar)newCalendar.clone();
				calendarMonth.set(Calendar.DAY_OF_MONTH, 1);
				currentMonthWeeks = 1;
				currentyearWeeks = 1;
			}
			
			dates[0] = format.format(date);
			dateList.add(dates);
			calendar.add(Calendar.DAY_OF_YEAR, -1);

			previouYear = year1;
			i++;
		}
		
		monthWeeksMap.put(String.valueOf(previouYear), currentMonthWeeks);
		yearWeeksMap.put(String.valueOf(previouYear), currentyearWeeks);
		
		for(i = dateList.size() - 1; i >= 0; i--) {
			dates = dateList.get(i);
			runDateList.add(dates);
		}
		
		return runDateList;
    }
	
	private Date stringToDate(String dateStr) {
		Date date = null;
		if (dateStr != null && dateStr.length() > 0) {
			try {
				date = format.parse(dateStr);
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
				Date date = format.parse(dateStr);
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
	
	public int outPutReport(String outPutFile, List<String[]> runDateList) {
		// get old report from Google Drive
		FileInfo file = upload.getInputFile();
		if (file == null) {
			logger.error("the readFile is not exists in google drive.");
			return -1;
		}
		
		maxDbDate = stringToDate(getMaxDateFromDb());
		if (maxDbDate == null) {
			logger.error("Can't get maxDbDate from db.");
			return -1;
		}
		
		Map<Integer, GAMarketingData> gaMarketingDataMap = setGAMarketingDataParameter();
		if (gaMarketingDataMap == null || gaMarketingDataMap.size() == 0) {
			logger.error("Can't get gaMarketing config Data from db.");
			return -1;
		}
		List<GAReportInfo> reportInfoList = getReportInfo(runDateList, gaMarketingDataMap);
		
		XSSFWorkbook xwb = new XSSFWorkbook();
		// get file content
		InputStream is = upload.downloadExcelFile(file.getId(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		//try again
		if (is == null) {
			is = upload.downloadExcelFile(file.getId(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		}
		if (is != null) {
			try {
				xwb = new XSSFWorkbook(is);
			} catch (IOException e) {
				logger.error(e.getMessage());
				e.printStackTrace();
				return -1;
			}
		} else {
			logger.error("download File from google drive error!");
			return -1;
		}
		
		int result = writeReport(xwb, "DASHBOARD", reportInfoList);
		
		if (result == 0) {
			OutputStream out;
			try {
				out = new FileOutputStream(outPutFile);
				xwb.write(out);
				if (out != null) {
					out.close();
				}
				logger.debug("write Excel file success. file=" + outPutFile);
			} catch (FileNotFoundException e) {
				e.printStackTrace();
				logger.error(e.getMessage());
			} catch (IOException e) {
				e.printStackTrace();
				logger.error(e.getMessage());
			}
		}
		
		// upload report To Google Drive
		if (reportInfoList != null && reportInfoList.size() > 0) {
			if (uploadFileFlag) {
				File uploadFile = upload.uploadExcelFile(outPutFile);
				// move source file to Old Folder
				if (uploadFile != null) {
					upload.moveFile(file.getId());
				}
				if (mailSendFlag) {
					SendMailUtil mail = new SendMailUtil(prop);
					if (uploadFile == null) {
						mail.setMail_subject("Upload GAMarketingReport File to Google drive error!");
						mail.sendMail();
					} else {
						mail.setMail_subject("Upload GAMarketingReport File to Google drive success!");
						mail.setMail_content("link: " + uploadFile.getAlternateLink());
						mail.sendMail();
					}
//					mail.sendBatchReport(outPutFile);
				}
			}
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
	
	private List<GAReportInfo> getReportInfo(List<String[]> runDateList, Map<Integer, GAMarketingData> gaMarketingDataMap) {
		List<GAReportInfo> reportInfoList = new ArrayList<GAReportInfo>();
		int pagenum = 10000;
		try {
			if (runDateList != null && runDateList.size() > 0) {
				int i = 0;
				for (String[] runDate:runDateList) {
					i++;
					Date startDate = stringToDate(runDate[0]);
					if (maxDbDate.after(startDate) && i <= runDateList.size() - weeks) {
						continue;
					}
					int days = (int)((stringToDate(runDate[1]).getTime() - startDate.getTime())/(1000*3600*24) + 1);
					for (int index:gaMarketingDataMap.keySet()) {
						getGaData(gaMarketingDataMap.get(index), runDate, index, days, pagenum);
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
			String[] runDate, int index, int days, int pagenum) throws ParseException {		
		// init
		// sessions
		gaMarketingData.setSessions(0);
		// pageviews
		gaMarketingData.setPageviews(0);
		
		String startDate = runDate[0];
		
		Calendar calendar = Calendar.getInstance();
		int times = 3;
		
		for(int past = 0 ; past < days; past++) {
			calendar.setTime(format.parse(startDate));
			boolean hasDbData = false;
			int preSessions = gaMarketingData.getSessions();
			int prePageviews = gaMarketingData.getPageviews();
			if (maxDbDate != null && !maxDbDate.before(calendar.getTime())) {
				hasDbData = getGaDataFromDb(gaMarketingData, startDate, index);
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
						String title = GAReportForGoogleDrive.class.getSimpleName() + "_" + index + "_" + startDate;
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
			startDate = format.format(calendar.getTime());
		}
		
		logger.info("Date is from " + runDate[0] + " To " + runDate[1] + ", index=" + index + ", total sessions="
		+ gaMarketingData.getSessions() + ", total pageviews=" + gaMarketingData.getPageviews());
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

	private int writeReport(XSSFWorkbook xwb, String sheetName,
			List<GAReportInfo> reportInfoList) {
		if (reportInfoList.size() == 0) {
			logger.error("size Error: reportInfoList size=" + reportInfoList.size());
			return -1;
		}
		GAMarketingReportWorkbook gaReport = new GAMarketingReportWorkbook(xwb);
		// j is level cells
		int j = 0;
		XSSFSheet sheet = null;
		XSSFRow row;
		String previouYear = "";
		String year = "";
		boolean firtSetting = false;
		for (int k = 0; k < reportInfoList.size(); k++) {
			GAReportInfo reportInfo = reportInfoList.get(k);
			String dateTo = reportInfo.getDateTo();
			year = dateTo.substring(0, 4);
			
			if (ComUtil.isNotEmpty(previouYear) && !previouYear.equals(year)) {
				row = sheet.getRow(0);
				if (row == null || row.getCell(j+1) == null) {
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
				} else {
					// MTD
					j++;
					gaReport.updateMTDColumn(sheet, j, monthWeeksMap.get(previouYear));
					
					// YTD
					j++;
					gaReport.updateYTDColumn(sheet, j);
					
					// Notes
					j++;
					
					// hidden rows and resize
					gaReport.hiddenRowsAndResize(sheet, j);
				}
			}
			
			if (!previouYear.equals(year)) {
				j = 0;
				String sheetNameStr = sheetName + "(" + year + ")";
				sheet = xwb.getSheet(sheetNameStr);
				if (sheet == null) {
					sheet = xwb.createSheet(sheetNameStr);
					XSSFSheet sheet2 = xwb.getSheet(sheetName + "(" + (Integer.parseInt(year)-1) + ")");
					if (sheet2 != null) {
						xwb.setSheetOrder(sheetNameStr, xwb.getSheetIndex(sheet2)+1);
					}
				}
				
				row = sheet.getRow(0);
				if (row == null || row.getCell(j) == null) {
					gaReport.addHeaderColumn(sheet, j);
				}
			}
			
			j++;
			row = sheet.getRow(0);
			if (row == null || row.getCell(j) == null) {
				gaReport.addDetailColumn(sheet, j, dateTo, reportInfo);
			} else {
				if (!firtSetting) {
					j = yearWeeksMap.get(year);
					if (weeks > 0) {
						j = j - weeks + 1;
					}
					firtSetting = true;
				}
				
				XSSFCell cell = gaReport.getSheetCell(gaReport.getSheetRow(sheet, 0), j);
				String dateStr = "";
				if (cell != null && cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
					try {
						dateStr = format.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
					} catch (Exception e) {
					}
				}
				boolean newColumn = false;
				if (!dateTo.equals(dateStr)) {
					ComExcelUtil.insertOneBlankColumn(sheet, j);
					newColumn = true;
				}
				gaReport.insertDetailColumn(sheet, j, dateTo, reportInfo, newColumn);
			}
			
			previouYear = year;
		}
		
		if (ComUtil.isNotEmpty(previouYear)) {
			row = sheet.getRow(0);
			if (row == null || row.getCell(j+1) == null) {
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
			} else {
				// MTD
				j++;
				gaReport.updateMTDColumn(sheet, j, monthWeeksMap.get(previouYear));
				
				// YTD
				j++;
				gaReport.updateYTDColumn(sheet, j);
				
				// Notes
				j++;
				
				// hidden rows and resize
				gaReport.hiddenRowsAndResize(sheet, j);
			}
			
			// the compare sheet
			XSSFSheet pastYearSheet = null;
			XSSFSheet currentYearSheet = null;
			previouYear = "";
			firtSetting = false;
			for (int k = 0; k < reportInfoList.size(); k++) {
				GAReportInfo reportInfo = reportInfoList.get(k);
				String dateTo = reportInfo.getDateTo();
				year = dateTo.substring(0, 4);
				int thisYear = Integer.parseInt(year);
				if (thisYear < 2015) {
					continue;
				}
				if (ComUtil.isNotEmpty(previouYear) && !previouYear.equals(year)) {					
						// hidden rows and resize
						gaReport.hiddenRowsAndResizeForCompareSheet(sheet, j);
				}
				
				if (!previouYear.equals(year)) {
					j = 0;
					String sheetNameStr = thisYear + " vs " + (thisYear-1);
					String currentYearSheetName = sheetName + "(" + year + ")";
					currentYearSheet = xwb.getSheet(currentYearSheetName);
					if (currentYearSheet == null) {
						logger.error("Can't get this year Sheet. name=" + currentYearSheetName);
						return -1;
					}
					String pastYearSheetName = sheetName + "(" + (thisYear-1) + ")";
					pastYearSheet = xwb.getSheet(pastYearSheetName);
					if (pastYearSheet == null) {
						logger.error("Can't get past year Sheet. name=" + pastYearSheetName);
						return -1;
					}
					sheet = xwb.getSheet(sheetNameStr);
					if (sheet == null) {
						sheet = xwb.createSheet(sheetNameStr);
						xwb.setSheetOrder(sheetNameStr, xwb.getSheetIndex(currentYearSheetName)+1);
					}
					
					row = sheet.getRow(0);
					if (row == null || row.getCell(j) == null) {
						gaReport.addCompareSheetHeaderColumn(sheet, currentYearSheet);
					}
				}
				
				j++;
				row = sheet.getRow(0);
				if (row != null && row.getCell(j) != null && !firtSetting) {
					j = yearWeeksMap.get(year);
					if (weeks > 0) {
						j = j - weeks + 1;
					}
					firtSetting = true;
				}
				
				XSSFCell cell = gaReport.getSheetCell(gaReport.getSheetRow(sheet, 0), j);
				String dateStr = "";
				if (cell != null && cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
					try {
						dateStr = format.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
					} catch (Exception e) {
					}
				}
				if (!dateTo.equals(dateStr)) {
					ComExcelUtil.insertOneBlankColumn(sheet, j);
				}
				
				gaReport.addCompareSheetDetailColumn(sheet, j, dateTo, pastYearSheet, currentYearSheet);
				
				previouYear = year;
			}
			
			if (ComUtil.isNotEmpty(previouYear)) {
				// hidden rows and resize
				gaReport.hiddenRowsAndResizeForCompareSheet(sheet, j);
			}
		}
		
		return 0;
	}

	/**
	 * @param args[0] args[1]
	 * 1: getFileList ($filename)
	 * 2: $runDate $weeks
	 * 3: $weeks $runDate
	 */
	public static void main(String[] args) {
		long startTime = System.currentTimeMillis();
		GAReportForGoogleDrive gaReport = new GAReportForGoogleDrive();
		List<String[]> runDateList = null;
		String runDateStr = null;
		
		// getFileList
		if (args != null && args.length > 0 && "getFileList".equalsIgnoreCase(args[0])) {
			if (args.length > 1) {
				gaReport.upload.printFileList(args[1]);
			} else {
				gaReport.upload.printFileList(null);
			}
			
			return;
		}
		
		if (args != null && args.length == 1) {
			if (!ComUtil.isDate(args[0], "yyyy-MM-dd")) {
				try {
					gaReport.weeks = Integer.parseInt(args[0]);
				} catch (Exception e) {
					logger.error("the parameter is not a day string(yyyy-MM-dd) or int!");
					return;
				}
			} else {
				runDateStr = args[0];
			}
		}
		
		if (args != null && args.length >= 2) {
			if (!ComUtil.isDate(args[0], "yyyy-MM-dd")) {
				try {
					gaReport.weeks = Integer.parseInt(args[0]);
				} catch (Exception e) {
					logger.error("the parameter is not a day string(yyyy-MM-dd) or int!");
					return;
				}
				runDateStr = args[1];
				if (!ComUtil.isDate(args[1], "yyyy-MM-dd")) {
					logger.error("the parameter is not a day string(yyyy-MM-dd) or int!");
					return;
				}
			} else {
				runDateStr = args[0];
				try {
					gaReport.weeks = Integer.parseInt(args[1]);
				} catch (Exception e) {
					logger.error("the parameter is not a day string(yyyy-MM-dd) or int!");
					return;
				}
			}
		}
		
		logger.info("prameter: runDateStr=" + runDateStr + ", weeks=" + gaReport.weeks);
		
		runDateList = gaReport.editRunDateList(runDateStr);
		
		// print runDateList
		for (String[] runDate:runDateList) {
			logger.info("dateFrom=" + runDate[0] + ", dateTo=" + runDate[1]);
		}
		
		String suffix = "";
		if (StringUtils.isNotEmpty(gaReport.fileNameDate)) {
			suffix = gaReport.dateStringformat(gaReport.fileNameDate);
		} else {
			logger.error("no current week fileName!");
			return;
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
		
		gaReport.outPutReport(outPutFile, runDateList);
		
		long endTime = System.currentTimeMillis();
		logger.debug("GAMarketingDataRequestReport total used time:" + (endTime-startTime) + "ms");
	}

}
