package com.ask.inv.google.analytics.report;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.ask.inv.google.analytics.model.GAReportInfo;

public class GAMarketingReportWorkbook {
	private static Logger logger = Logger.getLogger(GAMarketingReportWorkbook.class);
	
	public GAMarketingReportWorkbook(XSSFWorkbook xwb) {
		Style.init(xwb);
	}

	public void addHeaderColumn(XSSFSheet sheet, int j) {
		XSSFRow row;
		XSSFCell cell;
		// i is vertical rows, j is level cells
		int i = 0;
		
		// Freeze
		sheet.createFreezePane(1,1);

		// Marketing Metrics
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineFont16Style);
		cell.setCellValue("Marketing Metrics");

		// FA Tool
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		cell.setCellValue("FA Tool");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Partners*");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerStyle);
		cell.setCellValue("Total");

		i++;
		// Total Marketing
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14Style2);
		cell.setCellValue("Total Marketing");
		sheet.createRow(i++).createCell(j).setCellValue("Contributors Content");
		sheet.createRow(i++).createCell(j).setCellValue("Contributors/Writers Signed");
		sheet.createRow(i++).createCell(j).setCellValue("  Contributors Original");
		sheet.createRow(i++).createCell(j).setCellValue("  Contributors Manual");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("  Contributors Feed");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerStyle);
		cell.setCellValue("Contributors Total");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("  Yahoo No Repub (PV)");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("  Yahoo Repub (PV)");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerStyle);
		cell.setCellValue("Yahoo Total (PV)");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Social");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("Other Syndication");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("External: Yahoo + Social + Syndication");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Internal: Contributors");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailBorderMediumStyle);
		cell.setCellValue("Total");

		i++;
		// Contributors
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14Style);
		cell.setCellValue("Contributors");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		cell.setCellValue("Contributors");
		sheet.createRow(i++).createCell(j).setCellValue("Prospects");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("Signed");

		i++;
		// Contributor Content
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		cell.setCellValue("Contributor Content");
		sheet.createRow(i++).createCell(j).setCellValue("MyBankTracker");
		sheet.createRow(i++).createCell(j).setCellValue("Benzinga");
		sheet.createRow(i++).createCell(j).setCellValue("Blackrock");
		sheet.createRow(i++).createCell(j).setCellValue("CarInsurance.com");
		sheet.createRow(i++).createCell(j).setCellValue("DennisMiller");
		sheet.createRow(i++).createCell(j).setCellValue("EconBrowser");
		sheet.createRow(i++).createCell(j).setCellValue("ETF Expert");
		sheet.createRow(i++).createCell(j).setCellValue("FMD Capital Mgmt");
		sheet.createRow(i++).createCell(j).setCellValue("Franklin Templeton");
		sheet.createRow(i++).createCell(j).setCellValue("Get Rich Slowly");
		sheet.createRow(i++).createCell(j).setCellValue("GOBankingRates");
		sheet.createRow(i++).createCell(j).setCellValue("Guggenheim");
		sheet.createRow(i++).createCell(j).setCellValue("HSH");
		sheet.createRow(i++).createCell(j).setCellValue("ImprovementCenter");
		sheet.createRow(i++).createCell(j).setCellValue("Insure.com");
		sheet.createRow(i++).createCell(j).setCellValue("james Gruber (Asia Confidential)");
		sheet.createRow(i++).createCell(j).setCellValue("Money Blue Book");
		sheet.createRow(i++).createCell(j).setCellValue("Moneyrates");
		sheet.createRow(i++).createCell(j).setCellValue("OnlineColleges");
		sheet.createRow(i++).createCell(j).setCellValue("Quinstreet");
		sheet.createRow(i++).createCell(j).setCellValue("Schools.com");
		sheet.createRow(i++).createCell(j).setCellValue("WSD");
		sheet.createRow(i++).createCell(j).setCellValue("WisePiggy");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("ShopRate");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerStyle);
		cell.setCellValue("Total");

		i++;
		// Contributor - Original Content
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		cell.setCellValue("Contributor - Original Content");
		sheet.createRow(i++).createCell(j).setCellValue("Blackrock");
		sheet.createRow(i++).createCell(j).setCellValue("Tood Gordon");
		sheet.createRow(i++).createCell(j).setCellValue("GoBankingRates");
		sheet.createRow(i++).createCell(j).setCellValue("Fox Business");
		sheet.createRow(i++).createCell(j).setCellValue("HomeAdvisor");
		sheet.createRow(i++).createCell(j).setCellValue("AmigoBulls");
		sheet.createRow(i++).createCell(j).setCellValue("FitSmallBusiness");
		sheet.createRow(i++).createCell(j).setCellValue("Global Futures");
		sheet.createRow(i++).createCell(j).setCellValue("Betterment");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("FMD Capital Mgmt");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerStyle);
		cell.setCellValue("Total Pieces of Content");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerStyle);
		cell.setCellValue("Total PVs ");

		i++;
		// Conitrbutor PVs Non Orig (Manual) 
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		cell.setCellValue("Conitrbutor PVs Non Orig (Manual)");
		sheet.createRow(i++).createCell(j).setCellValue("benzinga");
		sheet.createRow(i++).createCell(j).setCellValue("blackrock");
		sheet.createRow(i++).createCell(j).setCellValue("CarInsurance");
		sheet.createRow(i++).createCell(j).setCellValue("casey-research");
		sheet.createRow(i++).createCell(j).setCellValue("dennis-miller");
		sheet.createRow(i++).createCell(j).setCellValue("econbrowser");
		sheet.createRow(i++).createCell(j).setCellValue("etf-expert");
		sheet.createRow(i++).createCell(j).setCellValue("fmd-capital-management");
		sheet.createRow(i++).createCell(j).setCellValue("foxbusiness");
		sheet.createRow(i++).createCell(j).setCellValue("franklin-templeton");
		sheet.createRow(i++).createCell(j).setCellValue("get-rich-slowly");
		sheet.createRow(i++).createCell(j).setCellValue("gobankingrates");
		sheet.createRow(i++).createCell(j).setCellValue("guggenheim-partners");
		sheet.createRow(i++).createCell(j).setCellValue("hsh");
		sheet.createRow(i++).createCell(j).setCellValue("Improvementcventer");
		sheet.createRow(i++).createCell(j).setCellValue("insurance");
		sheet.createRow(i++).createCell(j).setCellValue("insure");
		sheet.createRow(i++).createCell(j).setCellValue("kapitall");
		sheet.createRow(i++).createCell(j).setCellValue("mauldin-economics");
		sheet.createRow(i++).createCell(j).setCellValue("moneyrates");
		sheet.createRow(i++).createCell(j).setCellValue("my-bank-tracker");
		sheet.createRow(i++).createCell(j).setCellValue("onlinecolleges");
		sheet.createRow(i++).createCell(j).setCellValue("quinstreet");
		sheet.createRow(i++).createCell(j).setCellValue("schools");
		sheet.createRow(i++).createCell(j).setCellValue("Schoolscom");
		sheet.createRow(i++).createCell(j).setCellValue("shoprate");
		sheet.createRow(i++).createCell(j).setCellValue("wall street daily");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("Wisepiggy");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerStyle);
		cell.setCellValue("Total PVs");

		i++;
		// Contributor PVs Non Orig (Feed)
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		cell.setCellValue("Contributor PVs Non Orig (Feed)");
		sheet.createRow(i++).createCell(j).setCellValue("Benzinga");
		sheet.createRow(i++).createCell(j).setCellValue("ETFDB");
		sheet.createRow(i++).createCell(j).setCellValue("MoneyShow");
		sheet.createRow(i++).createCell(j).setCellValue("StreetAuthority");
		sheet.createRow(i++).createCell(j).setCellValue("Zacks");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("ForexNews");
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerStyle);
		cell.setCellValue("Total PVs");

		i+=2;
		// vertical head title
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14Style);
		cell.setCellValue("Syndication");

		// Yahoo Sessions
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		cell.setCellValue("Yahoo Sessions");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("  Yahoo Finance");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("  Yahoo Other");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("  Yahoo Full Content");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("  Yahoo SA Quote Pages");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Yahoo No Repub");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("Yahoo Repub");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerStyle);
		cell.setCellValue("Yahoo Total");

		i++;
		// Yahoo PVs
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		cell.setCellValue("Yahoo PVs");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("  Yahoo Finance");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("  Yahoo Other");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("  Yahoo Full Content");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("  Yahoo SA Quote Pages");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Yahoo No Repub");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("Yahoo Repub");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerStyle);
		cell.setCellValue("Yahoo Total");

		i++;
		// Other Sessions
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		cell.setCellValue("Other Sessions");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Nasdaq");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Forbes");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Fox Business");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("MSN/NBC");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Google News");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Blackrock");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Ask");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("Other");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerStyle);
		cell.setCellValue("Other Total Sessions");

		i++;
		// Other PVs
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		cell.setCellValue("Other PVs");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Nasdaq");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Forbes");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Fox Business");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("MSN/NBC");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Google News");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Blackrock");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Ask");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("Other");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerStyle);
		cell.setCellValue("Other Total PVs");

		i += 3;
		// Social
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14Style);
		cell.setCellValue("Social");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		cell.setCellValue("Social (All Networks)");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("All Visits");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("All PVs");

		i++;
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("All Posts");

		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("All PVs from Posting");

		i++;
		// FB Followers
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("FB Followers");
		// FB Reach
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("FB Reach");
		// FB Engagement
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("FB Engagement");
		i++;
		// Followers - (TW, LI, YT, G+, TOD)
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Followers - (TW, LI, YT, G+, TOD)");
		// YouTube Video Views
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("YouTube Video Views");
		i++;
		// Partner On-Site Social Shares
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Partner On-Site Social Shares");
		// On-site Social Shares
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("On-site Social Shares");

		i+=3;
		// Customer Service
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14Style);
		cell.setCellValue("Customer Service");
		// Customer Marketing
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		cell.setCellValue("Customer Marketing");
		// Positive Comments
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Positive Comments");
		// Negative Comments
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Negative Comments");
		// Product Suggestions
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Product Suggestions");
		// Contributor Inquiries
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("Contributor Inquiries");
		// Total
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerStyle);
		cell.setCellValue("Total");

		i++;
		// Weekly Negative Comments 10/11/14:
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		cell.setCellValue("Weekly Negative Comments 10/11/14:");
		// 36% – Unsubscribes (8)
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("36% – Unsubscribes (8)");
		// 9% – Delete Account (2)
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("9% – Delete Account (2)");
		// 32% – Editorial (7)
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("32% – Editorial (7)");
		// 23% – Product (5)
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		cell.setCellValue("23% – Product (5)");

		i+=3;
		// Writers
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14Style);
		cell.setCellValue("Writers");
		// Writers
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		cell.setCellValue("Writers");
		// Prospects
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Prospects");
		// Signed
		row = sheet.createRow(i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Signed");
	}
	
	public void addDetailColumn(XSSFSheet sheet, int j, String dateTo, GAReportInfo reportInfo) {
		XSSFRow row;
		XSSFCell cell;
		// i is vertical rows, j is level cells
		int i = 0;
		
		//date
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellValue(stringToDate(dateTo));
		cell.setCellStyle(Style.headerDateStyle);
		
		// FA Tool
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldStyle);
		
		//blank
		i++;
		// Total Marketing
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14Style2);
		// Contributors Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		String ref1 = null;
		String ref2 = null;
//		String ref1 = getSheetCell(getSheetRow(sheet, i+44), j).getReference();
//		String ref2 = getSheetCell(getSheetRow(sheet, i+57), j).getReference();
//		cell.setCellFormula("(" + ref1 + "+" + ref2 + ")");
		// Contributors/Writers Signed
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
//		ref1 = getSheetCell(getSheetRow(sheet, i+16), j).getReference();
//		ref2 = getSheetCell(getSheetRow(sheet, i+182), j).getReference();
//		cell.setCellFormula(ref1 + "+" + ref2);
		// Contributors Original
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
//		ref1 = getSheetCell(getSheetRow(sheet, i+56), j).getReference();
//		cell.setCellFormula(ref1);
		// Contributors Manual
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
//		ref1 = getSheetCell(getSheetRow(sheet, i+86), j).getReference();
//		cell.setCellFormula(ref1);
		// Contributors Feed
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineStyle);
//		ref1 = getSheetCell(getSheetRow(sheet, i+94), j).getReference();
//		cell.setCellFormula(ref1);
		// Contributors Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldStyle);
//		ref1 = getSheetCell(getSheetRow(sheet, i-4), j).getReference();
//		ref2 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
//		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		// Yahoo No Repub (PV)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
//		ref1 = getSheetCell(getSheetRow(sheet, i+110), j).getReference();
//		cell.setCellFormula(ref1);
		// Yahoo Repub (PV)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineStyle);
//		ref1 = getSheetCell(getSheetRow(sheet, i+110), j).getReference();
//		cell.setCellFormula(ref1);
		// Yahoo Total (PV)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldStyle);
//		ref1 = getSheetCell(getSheetRow(sheet, i+110), j).getReference();
//		cell.setCellFormula(ref1);
		// Social
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
//		ref1 = getSheetCell(getSheetRow(sheet, i+138), j).getReference();
//		cell.setCellFormula(ref1);
		// Other Syndication
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineStyle);
//		ref1 = getSheetCell(getSheetRow(sheet, i+130), j).getReference();
//		cell.setCellFormula(ref1);
		// External: Yahoo + Social + Syndication
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
//		ref1 = getSheetCell(getSheetRow(sheet, i-4), j).getReference();
//		ref2 = getSheetCell(getSheetRow(sheet, i-3), j).getReference();
//		String ref3 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
//		cell.setCellFormula("SUM("+ref1 + "," + ref2 + "," + ref3 + ")");
		// Internal: Contributors
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
//		ref1 = getSheetCell(getSheetRow(sheet, i-8), j).getReference();
//		cell.setCellFormula(ref1);
		// Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBorderMediumStyle);
//		ref1 = getSheetCell(getSheetRow(sheet, i-3), j).getReference();
//		ref2 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
//		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		
		//blank
		i++;
		// Contributors
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14Style);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineStyle);
		
		//blank
		i++;
		// Contributor Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		for (int l = 1; l <= 23; l++) {
			row = getSheetRow(sheet, i++);
    		cell=getSheetCell(row, j);
    		cell.setCellStyle(Style.numberStyle);
		}
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineStyle);
		// Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldStyle);
		ref1 = getSheetCell(getSheetRow(sheet, i-25), j).getReference();
		ref2 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");

		//blank
		i++;
		// Contributor - Original Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		for (int l = 1; l <= 9; l++) {
			row = getSheetRow(sheet, i++);
    		cell=getSheetCell(row, j);
    		cell.setCellStyle(Style.numberStyle);
		}
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineStyle);
		// Total Pieces of Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldStyle);
		ref1 = getSheetCell(getSheetRow(sheet, i-11), j).getReference();
		ref2 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		// Total PVs 
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldStyle);
		
		//blank
		i++;
		// Conitrbutor PVs Non Orig (Manual) 
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		for (int l = 1; l <= 27; l++) {
			row = getSheetRow(sheet, i++);
    		cell=getSheetCell(row, j);
    		cell.setCellStyle(Style.numberStyle);
		}
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineStyle);
		// Total PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldStyle);
		ref1 = getSheetCell(getSheetRow(sheet, i-29), j).getReference();
		ref2 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		
		//blank
		i++;
		// Contributor PVs Non Orig (Feed)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		for (int l = 1; l <= 5; l++) {
			row = getSheetRow(sheet, i++);
    		cell=getSheetCell(row, j);
    		cell.setCellStyle(Style.numberStyle);
		}
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineStyle);
		// Total PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldStyle);
		ref1 = getSheetCell(getSheetRow(sheet, i-7), j).getReference();
		ref2 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");

		// blank
		i++;
		// blank
		i++;
		//Syndication
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14Style);
		
		// Yahoo Sessions
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		// Yahoo Finance
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getYahooFinance());
		// Yahoo Other
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getYahooOther());
		// Yahoo Full Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getYahooFullContent());
		// Yahoo SA Quote Pages
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineStyle);
		cell.setCellValue(reportInfo.getYahooSAQuotePages());
		// Yahoo No Repub
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		String startRef = getSheetCell(getSheetRow(sheet, i - 5), j).getReference();
		String endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellStyle(Style.numberStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo Repub
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineStyle);
		cell.setCellValue(reportInfo.getYahooRepub());
		// Yahoo Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldStyle);
		startRef = getSheetCell(getSheetRow(sheet, i - 3), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellFormula(startRef + "+" + endRef);
		
		//blank
		i++;
		// Yahoo PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		// Yahoo Finance
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getYahooFinancePvs());
		// Yahoo Other
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getYahooOtherPvs());
		// Yahoo Full Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getYahooFullContentPvs());
		// Yahoo SA Quote Pages
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineStyle);
		cell.setCellValue(reportInfo.getYahooSAQuotePagesPvs());
		// Yahoo No Repub
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(getSheetRow(sheet, i - 5), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellStyle(Style.numberStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo Repub
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineStyle);
		cell.setCellValue(reportInfo.getYahooRepubPvs());
		// Yahoo Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldStyle);
		startRef = getSheetCell(getSheetRow(sheet, i - 3), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellFormula(startRef + "+" + endRef);
		
		//blank
		i++;
		// Other Sessions
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		// Nasdaq
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getNasdaq());
		// Forbes
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getForbes());
		// Fox Business
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getFoxBusiness());
		// MSN/NBC
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getMsnNbc());
		// Google News
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getGoogleNews());
		// Blackrock
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getBlackrock());
		// Ask
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getAsk());
		// Other
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineStyle);
		cell.setCellValue(reportInfo.getOther());
		// Other Total Sessions
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldStyle);
		startRef = getSheetCell(getSheetRow(sheet, i - 9), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		
		//blank
		i++;
		// Other PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		// Nasdaq
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getNasdaqPvs());
		// Forbes
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getForbesPvs());
		// Fox Business
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getFoxBusinessPvs());
		// MSN/NBC
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getMsnNbcPvs());
		// Google News
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getGoogleNewsPvs());
		// Blackrock
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getBlackrockPvs());
		// Ask
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getAskPvs());
		// Other
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineStyle);
		cell.setCellValue(reportInfo.getOtherPvs());
		// Other Total PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldStyle);
		startRef = getSheetCell(getSheetRow(sheet, i - 9), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		
		// blank
		i+=3;
		// Social
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14Style);
		// Social (All Networks
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		// All Visits
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getAllVisits());
		// All PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getAllPvs());
		
		//blank
		i++;
		// All Posts
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getAllPosts());
		// All PVs from Posting
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		cell.setCellValue(reportInfo.getAllPVsFromPosting());

		//blank
		i++;
		// FB Followers
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		// FB Reach
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		// FB Engagement
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		//blank
		i++;
		// Followers - (TW, LI, YT, G+, TOD)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		// YouTube Video Views
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		//blank
		i++;
		// Partner On-Site Social Shares
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		// On-site Social Shares
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		
		// blank
		i+=3;
		// Customer Service
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14Style);
		// Customer Marketing
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		// Positive Comments
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		// Negative Comments
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		// Product Suggestions
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		// Contributor Inquiries
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineStyle);
		// Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldStyle);
		startRef = getSheetCell(getSheetRow(sheet, i - 5), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		
		//blank
		i++;
		// Weekly Negative Comments 10/11/14:
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		// 36% – Unsubscribes (8)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		// 9% – Delete Account (2)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		// 32% – Editorial (7)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		// 23% – Product (5)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineStyle);
		
		// blank
		i+=3;
		// Writers
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14Style);
		// Writers
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineStyle);
		// Prospects
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
		// Signed
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberStyle);
	}
	
    public void addMTDColumn(XSSFSheet sheet, int j, int currentMonthWeeks) {
    	XSSFRow row;
		XSSFCell cell;
		// i is vertical rows, j is level cells
		int i = 0;
		String ref1,ref2,ref3,startRef,endRef;
		//date
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("MTD");
		cell.setCellStyle(Style.headerBorderMediumBlStyle);

		// FA Tool
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBlStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
		cell.setCellStyle(Style.numberBlStyle);
		cell.setCellFormula("SUM(" + ref1 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// Total Marketing
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14BlStyle2);
		// Contributors Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+44), j).getReference();
		ref2 = getSheetCell(getSheetRow(sheet, i+57), j).getReference();
		cell.setCellStyle(Style.numberBlStyle);
		cell.setCellFormula("(" + ref1 + "+" + ref2 + ")");
		// Contributors/Writers Signed
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+16), j).getReference();
		ref2 = getSheetCell(getSheetRow(sheet, i+182), j).getReference();
		cell.setCellStyle(Style.numberBlStyle);
		cell.setCellFormula(ref1 + "+" + ref2);
		// Contributors Original
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+56), j).getReference();
		cell.setCellStyle(Style.numberBlStyle);
		cell.setCellFormula(ref1);
		// Contributors Manual
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+86), j).getReference();
		cell.setCellStyle(Style.numberBlStyle);
		cell.setCellFormula(ref1);
		// Contributors Feed
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+94), j).getReference();
		cell.setCellStyle(Style.numberLineBlStyle);
		cell.setCellFormula(ref1);
		// Contributors Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBlStyle);
		ref1 = getSheetCell(getSheetRow(sheet, i-4), j).getReference();
		ref2 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		// Yahoo No Repub (PV)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+110), j).getReference();
		cell.setCellStyle(Style.numberBlStyle);
		cell.setCellFormula(ref1);
		// Yahoo Repub (PV)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+110), j).getReference();
		cell.setCellStyle(Style.numberLineBlStyle);
		cell.setCellFormula(ref1);
		// Yahoo Total (PV)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBlStyle);
		ref1 = getSheetCell(getSheetRow(sheet, i+110), j).getReference();
		cell.setCellFormula(ref1);
		// Social
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+138), j).getReference();
		cell.setCellStyle(Style.numberBlStyle);
		cell.setCellFormula(ref1);
		// Other Syndication
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+130), j).getReference();
		cell.setCellStyle(Style.numberLineBlStyle);
		cell.setCellFormula(ref1);
		// External: Yahoo + Social + Syndication
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i-4), j).getReference();
		ref2 = getSheetCell(getSheetRow(sheet, i-3), j).getReference();
		ref3 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
		cell.setCellStyle(Style.numberBlStyle);
		cell.setCellFormula("SUM("+ref1 + "," + ref2 + "," + ref3 + ")");
		// Internal: Contributors
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i-8), j).getReference();
		cell.setCellStyle(Style.numberBlStyle);
		cell.setCellFormula(ref1);
		// Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i-3), j).getReference();
		ref2 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
		cell.setCellStyle(Style.numberBorderMediumBlStyle);
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// Contributors
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14BlStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBlStyle);
		// Prospects
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Signed
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// Contributor Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBlStyle);
		for (int k = 1; k <= 23; k++) {
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.numberBlStyle);
			ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
			ref2 = getSheetCell(row, j-1).getReference();
			cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		}
		// ShopRate
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// Contributor - Original Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBlStyle);
		for (int k = 1; k <= 9; k++) {
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.numberBlStyle);
			ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
			ref2 = getSheetCell(row, j-1).getReference();
			cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		}
		// FMD Capital Mgmt
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Total Pieces of Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Total PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// Conitrbutor PVs Non Orig (Manual) 
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBlStyle);
		for (int k = 1; k <= 27; k++) {
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.numberBlStyle);
			ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
			ref2 = getSheetCell(row, j-1).getReference();
			cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		}
		// Wisepiggy
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Total PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// Contributor PVs Non Orig (Feed)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBlStyle);
		for (int k = 1; k <= 5; k++) {
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.numberBlStyle);
			ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
			ref2 = getSheetCell(row, j-1).getReference();
			cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		}
		// ForexNews
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Total PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		//Syndication
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14BlStyle);
		// Yahoo Sessions
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBlStyle);
		// Yahoo Finance
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Yahoo Other
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Yahoo Full Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Yahoo SA Quote Pages
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Yahoo No Repub
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(getSheetRow(sheet, i - 5), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellStyle(Style.numberBlStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo Repub
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Yahoo Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBlStyle);
		startRef = getSheetCell(getSheetRow(sheet, i - 3), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// Yahoo PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBlStyle);
		// Yahoo Finance
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Yahoo Other
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Yahoo Full Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Yahoo SA Quote Pages
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Yahoo No Repub
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(getSheetRow(sheet, i - 5), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellStyle(Style.numberBlStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo Repub
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Yahoo Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBlStyle);
		startRef = getSheetCell(getSheetRow(sheet, i - 3), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// Other Sessions
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBlStyle);
		// Nasdaq
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Forbes
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Fox Business
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// MSN/NBC
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Google News
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Blackrock
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Ask
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Other
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Other Total Sessions
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBlStyle);
		startRef = getSheetCell(getSheetRow(sheet, i - 9), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// Other PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBlStyle);
		// Nasdaq
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Forbes
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Fox Business
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// MSN/NBC
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Google News
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Blackrock
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Ask
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Other
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Other Total PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBlStyle);
		startRef = getSheetCell(getSheetRow(sheet, i - 9), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");

		for (int k = 0; k < 3; k++) {
			// blank
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.blankBlStyle);
		}
		// Social
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14BlStyle);
		// Social (All Networks)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBlStyle);
		// All Visits
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// All PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// All Posts
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// All PVs from Posting
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// FB Followers
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// FB Reach
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// FB Engagement
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// Followers - (TW, LI, YT, G+, TOD)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// YouTube Video Views
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// Partner On-Site Social Shares
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// On-site Social Shares
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");

		for (int k = 0; k < 3; k++) {
			// blank
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.blankBlStyle);
		}
		// Customer Service
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14BlStyle);
		// Customer Marketing
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBlStyle);
		// Positive Comments
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Negative Comments
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Product Suggestions
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Contributor Inquiries
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// Weekly Negative Comments 10/11/14:
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBlStyle);
		// 36% – Unsubscribes (8)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// 9% – Delete Account (2)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// 32% – Editorial (7)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBlStyle);
		// 23% – Product (5)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineBlStyle);

		for (int k = 0; k < 3; k++) {
			// blank
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.blankBlStyle);
		}
		// Writers
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14BlStyle);
		// Writers
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBlStyle);
		// Prospects
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
		// Signed
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBlStyle);
		ref1 = getSheetCell(row, j-currentMonthWeeks).getReference();
		ref2 = getSheetCell(row, j-1).getReference();
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");
	}
    
    public void addYTDColumn(XSSFSheet sheet, int j) {
    	XSSFRow row;
		XSSFCell cell;
		// i is vertical rows, j is level cells
		int i = 0;
		String ref1,ref2,ref3,startRef,endRef;
		
		// date
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("YTD");
		cell.setCellStyle(Style.headerBorderMediumBrStyle);

		// FA Tool
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBrStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBrStyle);
		ref1 = getSheetCell(row, 1).getReference();
		ref2 = getSheetCell(row, j-2).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBrStyle);
		ref1 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
		cell.setCellFormula("SUM(" + ref1 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// Total Marketing
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14BrStyle2);
		// Contributors Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+44), j).getReference();
		ref2 = getSheetCell(getSheetRow(sheet, i+57), j).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("(" + ref1 + "+" + ref2 + ")");
		// Contributors/Writers Signed
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+16), j).getReference();
		ref2 = getSheetCell(getSheetRow(sheet, i+182), j).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula(ref1 + "+" + ref2);
		// Contributors Original
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+56), j).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula(ref1);
		// Contributors Manual
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+86), j).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula(ref1);
		// Contributors Feed
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+94), j).getReference();
		cell.setCellStyle(Style.numberLineBrStyle);
		cell.setCellFormula(ref1);
		// Contributors Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBrStyle);
		ref1 = getSheetCell(getSheetRow(sheet, i-4), j).getReference();
		ref2 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		// Yahoo No Repub (PV)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+110), j).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula(ref1);
		// Yahoo Repub (PV)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+110), j).getReference();
		cell.setCellStyle(Style.numberLineBrStyle);
		cell.setCellFormula(ref1);
		// Yahoo Total (PV)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+110), j).getReference();
		cell.setCellStyle(Style.numberBoldBrStyle);
		cell.setCellFormula(ref1);
		// Social
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+138), j).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula(ref1);
		// Other Syndication
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i+130), j).getReference();
		cell.setCellStyle(Style.numberLineBrStyle);
		cell.setCellFormula(ref1);
		// External: Yahoo + Social + Syndication
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i-4), j).getReference();
		ref2 = getSheetCell(getSheetRow(sheet, i-3), j).getReference();
		ref3 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM("+ref1 + "," + ref2 + "," + ref3 + ")");
		// Internal: Contributors
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i-8), j).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula(ref1);
		// Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		ref1 = getSheetCell(getSheetRow(sheet, i-3), j).getReference();
		ref2 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
		cell.setCellStyle(Style.numberBorderMediumBrStyle);
		cell.setCellFormula("SUM("+ref1 + ":" + ref2 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// Contributors
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14BrStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBrStyle);
		// Prospects
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBrStyle);
		ref1 = getSheetCell(row, 1).getReference();
		ref2 = getSheetCell(row, j-2).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		// Signed
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBrStyle);
		ref1 = getSheetCell(row, 1).getReference();
		ref2 = getSheetCell(row, j-2).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// Contributor Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBrStyle);
		for (int k = 1; k <= 23; k++) {
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.numberBrStyle);
			ref1 = getSheetCell(row, 1).getReference();
			ref2 = getSheetCell(row, j-2).getReference();
			cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		}
		// ShopRate
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBrStyle);
		ref1 = getSheetCell(row, 1).getReference();
		ref2 = getSheetCell(row, j-2).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		// Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBrStyle);
		ref1 = getSheetCell(getSheetRow(sheet, i-25), j).getReference();
		ref2 = getSheetCell(getSheetRow(sheet, i-2), j).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// Contributor - Original Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBrStyle);
		for (int k = 1; k <= 9; k++) {
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.numberBrStyle);
			ref1 = getSheetCell(row, 1).getReference();
			ref2 = getSheetCell(row, j-2).getReference();
			cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		}
		// FMD Capital Mgmt
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBrStyle);
		ref1 = getSheetCell(row, 1).getReference();
		ref2 = getSheetCell(row, j-2).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		// Total Pieces of Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBrStyle);
		ref1 = getSheetCell(row, 1).getReference();
		ref2 = getSheetCell(row, j-2).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		// Total PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBrStyle);
		ref1 = getSheetCell(row, 1).getReference();
		ref2 = getSheetCell(row, j-2).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// Conitrbutor PVs Non Orig (Manual) 
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBrStyle);
		for (int k = 1; k <= 27; k++) {
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.numberBrStyle);
			ref1 = getSheetCell(row, 1).getReference();
			ref2 = getSheetCell(row, j-2).getReference();
			cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		}
		// Wisepiggy
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBrStyle);
		ref1 = getSheetCell(row, 1).getReference();
		ref2 = getSheetCell(row, j-2).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		// Total PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBrStyle);
		ref1 = getSheetCell(row, 1).getReference();
		ref2 = getSheetCell(row, j-2).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// Contributor PVs Non Orig (Feed)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBrStyle);
		for (int k = 1; k <= 5; k++) {
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.numberBrStyle);
			ref1 = getSheetCell(row, 1).getReference();
			ref2 = getSheetCell(row, j-2).getReference();
			cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		}
		// ForexNews
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBrStyle);
		ref1 = getSheetCell(row, 1).getReference();
		ref2 = getSheetCell(row, j-2).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");
		// Total PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBrStyle);
		ref1 = getSheetCell(row, 1).getReference();
		ref2 = getSheetCell(row, j-2).getReference();
		cell.setCellFormula("SUM(" + ref1 + ":" + ref2 + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		//Syndication
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14BrStyle);
		// Yahoo Sessions
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBrStyle);
		// Yahoo Finance
		row = getSheetRow(sheet, i++);
		cell = getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo Other
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo Full Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo SA Quote Pages
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBrStyle);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo No Repub
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(getSheetRow(sheet, i - 5), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo Repub
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBrStyle);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBrStyle);
		startRef = getSheetCell(getSheetRow(sheet, i - 3), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// Yahoo PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBrStyle);
		// Yahoo Finance
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo Other
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo Full Content
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo SA Quote Pages
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBrStyle);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo No Repub
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(getSheetRow(sheet, i - 5), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo Repub
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBrStyle);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Yahoo Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBrStyle);
		startRef = getSheetCell(getSheetRow(sheet, i - 3), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// Other Sessions
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBrStyle);
		// Nasdaq
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Forbes
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Fox Business
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// MSN/NBC
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Google News
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Blackrock
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBrStyle);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Ask
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBrStyle);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Other
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBrStyle);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Other Total Sessions
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBrStyle);
		startRef = getSheetCell(getSheetRow(sheet, i - 9), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// Other PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBrStyle);
		// Nasdaq
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Forbes
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Fox Business
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// MSN/NBC
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Google News
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Blackrock
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBrStyle);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Ask
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBrStyle);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Other
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberLineBrStyle);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Other Total PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.numberBoldBrStyle);
		startRef = getSheetCell(getSheetRow(sheet, i - 9), j).getReference();
		endRef = getSheetCell(getSheetRow(sheet, i - 2), j).getReference();
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// Social
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14BrStyle);
		// Social (All Networks
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBrStyle);
		// All Visits
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// All PVs
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// All Posts
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// All PVs from Posting
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// FB Followers
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// FB Reach
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// FB Engagement
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// Followers - (TW, LI, YT, G+, TOD)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// YouTube Video Views
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// Partner On-Site Social Shares
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// On-site Social Shares
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// Customer Service
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14BrStyle);
		// Customer Marketing
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBrStyle);
		// Positive Comments
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Negative Comments
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Product Suggestions
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Contributor Inquiries
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberLineBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Total
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBoldBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// Weekly Negative Comments 10/11/14:
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBrStyle);
		// 36% – Unsubscribes (8)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// 9% – Delete Account (2)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// 32% – Editorial (7)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// 23% – Product (5)
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.detailLineBrStyle);

		// blank
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.blankBrStyle);
		// Writers
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerBorderMediumFont14BrStyle);
		// Writers
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellStyle(Style.headerLineBrStyle);
		// Prospects
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
		// Signed
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		startRef = getSheetCell(row, 1).getReference();
		endRef = getSheetCell(row, j - 2).getReference();
		cell.setCellStyle(Style.numberBrStyle);
		cell.setCellFormula("SUM(" + startRef + ":" + endRef + ")");
    }
	
	public void addNotesColumn(XSSFSheet sheet, int j) {
		XSSFRow row;
		XSSFCell cell;
		// i is vertical rows, j is level cells
		int i = 0;

		// date
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Notes");
		cell.setCellStyle(Style.headerBorderMediumBrnStyle);
		for (int k = 0; k < 22; k++) {
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.blankBrnStyle);
		}
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Combine Contributors & Writers?");
		cell.setCellStyle(Style.blankBrnStyle);
		for (int k = 0; k < 111; k++) {
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.blankBrnStyle);
		}
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Roll in Fox");
		cell.setCellStyle(Style.blankBrnStyle);
		for (int k = 0; k < 10; k++) {
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.blankBrnStyle);
		}
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("Roll in Fox");
		cell.setCellStyle(Style.blankBrnStyle);
		for (int k = 0; k < 19; k++) {
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.blankBrnStyle);
		}
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("ShareThis");
		cell.setCellStyle(Style.blankBrnStyle);
		row = getSheetRow(sheet, i++);
		cell=getSheetCell(row, j);
		cell.setCellValue("ShareThis");
		cell.setCellStyle(Style.blankBrnStyle);
		for (int k = 0; k < 23; k++) {
			row = getSheetRow(sheet, i++);
			cell=getSheetCell(row, j);
			cell.setCellStyle(Style.blankBrnStyle);
		}
	}
	
	public void hiddenRowsAndResize(XSSFSheet sheet, int j) {
		XSSFRow row;
		// hidden rows
		row = getSheetRow(sheet, 1);
		row.setZeroHeight(true);
		row = getSheetRow(sheet, 2);
		row.setZeroHeight(true);
		row = getSheetRow(sheet, 3);
		row.setZeroHeight(true);
		
		for (short k = 0; k <= j; k++) {
			sheet.autoSizeColumn(k);
		}
		short k = 0;
		for (k = 0; k <= j; k++) {
			if (k != 0) {
				if (k == j - 1) {
					sheet.setColumnWidth(k, sheet.getColumnWidth(k) + 2048);
				} else if (k == j - 2) {
					sheet.setColumnWidth(k, sheet.getColumnWidth(k) + 1536);
				} else if (k == j) {
					sheet.setColumnWidth(k, sheet.getColumnWidth(k) + 2048);
				} else {
					sheet.setColumnWidth(k, sheet.getColumnWidth(k) + 512);
				}
			}
		}
		
		sheet.setZoom(80);
	}
	public XSSFRow getSheetRow(XSSFSheet sheet, int rownum) {
    	if (sheet == null) {
    		return null;
    	}
    	XSSFRow row = sheet.getRow(rownum);
    	if (row == null) {
    		row = sheet.createRow(rownum);
    	}
    	return row;
    }
	
	public XSSFCell getSheetCell(XSSFRow row, int cellnum) {
    	if (row == null) {
    		return null;
    	}
    	XSSFCell cell = row.getCell(cellnum);
    	if (cell == null) {
    		cell = row.createCell(cellnum);
    	}
    	return cell;
    }
	
	public Date stringToDate(String dateStr) {
		DateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
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
}
