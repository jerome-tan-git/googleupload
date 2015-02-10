package com.ask.inv.Util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.TimeZone;

public class ComUtil {

	public static boolean isDate(String dateString, String format) {
		if (dateString == null || dateString.length() == 0) {
			return false;
		}
		SimpleDateFormat sdf = new SimpleDateFormat(format);
        try {
			if (dateString.equals(sdf.format(sdf.parse(dateString)))) {  
			    return true;  
			} else {  
			    return false;
			}
		} catch (ParseException e) {
			return false;
		}
	}
	
	public static boolean dateAfter(String dateBegin, String dateEnd) {
    	boolean flag = false;
    	SimpleDateFormat df = new SimpleDateFormat("MM/dd/yyyy");
    	try {
			Date date1 = df.parse(dateBegin);
			Date date2 = df.parse(dateEnd);
			flag = date2.after(date1);
		} catch (ParseException e) {
			flag = false;
		}
    	
    	return flag;
    }
	
	public static int dateCompare(String dateBegin, String dateEnd) {
    	int flag = -2;
    	SimpleDateFormat df = new SimpleDateFormat("MM/dd/yyyy");
    	try {
			Date date1 = df.parse(dateBegin);
			Date date2 = df.parse(dateEnd);
			flag = date1.compareTo(date2);
		} catch (ParseException e) {
			e.printStackTrace();
		}
    	
    	return flag;
    }
	
	public static Date stringToDate(String dateStr) {
		Date date = null;
		if (isEmpty(dateStr)) {
			return null;
		}
    	SimpleDateFormat df = new SimpleDateFormat("MM/dd/yyyy");
    	try {
			date = df.parse(dateStr);
		} catch (ParseException e) {
			e.printStackTrace();
		}
    	
    	return date;
    }
	
	public static Date stringToDate(String dateStr, String format) {
		Date date = null;
		if (isEmpty(dateStr)) {
			return null;
		}
    	SimpleDateFormat df = new SimpleDateFormat(format);
    	try {
			date = df.parse(dateStr);
		} catch (ParseException e) {
			e.printStackTrace();
		}
    	
    	return date;
    }
	
	public static String dateTostring(Date date, String format) {
		String dateStr = null;
		if (date == null) {
			return null;
		}
    	SimpleDateFormat df = new SimpleDateFormat(format);
    	try {
    		dateStr = df.format(date);
		} catch (Exception e) {
			e.printStackTrace();
		}
    	
    	return dateStr;
    }
	
	public static Date stringToDateHHmmss(String dateStr) {
		Date date = null;
		if (isEmpty(dateStr)) {
			return null;
		}
    	SimpleDateFormat df = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
    	df.setTimeZone(TimeZone.getTimeZone("UTC"));
    	try {
			date = df.parse(dateStr);
		} catch (ParseException e) {
			e.printStackTrace();
		}
    	
    	return date;
    }
	
	public static Date stringToDatehhmmss(String dateStr) {
		Date date = null;
		if (ComUtil.isEmpty(dateStr)) {
			return null;
		}
    	SimpleDateFormat df = new SimpleDateFormat("MM/dd/yyyy hh:mm:ss a");
    	df.setTimeZone(TimeZone.getTimeZone("UTC"));
    	try {
			date = df.parse(dateStr);
		} catch (ParseException e) {
			e.printStackTrace();
		}
    	
    	return date;
    }
	
	public static String dateFormatString(String dateStr) {
		SimpleDateFormat df = new SimpleDateFormat("MM/dd/yyyy");
		String dateString = "";
		try {
			Date date = df.parse(dateStr);
			df = new SimpleDateFormat("yyyyMMdd");
			dateString = df.format(date);
		} catch (ParseException e) {
			e.printStackTrace();
		}
		
		return dateString;
	}
	
	public static String dateStringFormat(String dateStr, String formatFrom, String formatTo) {
		SimpleDateFormat df = new SimpleDateFormat(formatFrom);
		String dateString = "";
		try {
			Date date = df.parse(dateStr);
			df = new SimpleDateFormat(formatTo);
			dateString = df.format(date);
		} catch (ParseException e) {
			e.printStackTrace();
		}
		
		return dateString;
	}
	
	public static boolean isEmpty(String str) {
		return str == null || str.length() == 0;
	}
	
	public static boolean isNotEmpty(String str) {
		return str != null && str.length() != 0;
	}
	
	public static String replaceSinglequotes(String str) {
    	if (str == null) {
    		return str;
    	} else {
    		return str.replace("'", "''");
    	}
    }
	
	public static int calculateWordCount(String str) {
		if (str == null || str.length() == 0) {
			return 0;
		}
		
		str = str.replaceAll("\r\n|\n|\r", " ")
                .replaceAll("^\\s+|\\s+$", "")
                .replaceAll("&nbsp;", " ");

		str = stripHTMLTag(str);

       String[] words = str.split("\\s+");
       int wordCount = 0;
        for (int wordIndex = words.length - 1; wordIndex >= 0; wordIndex--) {
            if (!words[wordIndex].matches("^([\\s\t\r\n]*)$")) {
            	wordCount++;
            }
        }

		return wordCount;
	}
	
	public static String stripHTMLTag(String str) {
		if (str == null || str.length() == 0) {
			return "";
		}
		str = str.replaceAll("</p\\s*>", "\r\n");
		str = str.replaceAll("<br\\s*/?>", "\r\n");
		str = str.replaceAll("<[^>]+>", "");
    	return str;
	}
	
	public static int charCountOfHTML(String str) {
		if (str == null || str.length() == 0) {
			return 0;
		}
		str = str.replaceAll("\r\n|\n|\r", "")
				.replaceAll("^\\s+|\\s+$", "")
				.replaceAll("&nbsp;", "")
				.replaceAll("\\s", "");

		str = stripHTMLTag(str).replaceAll("^([\\s\t\r\n]*)$", "");

		return str.length();
	}
}
