package com.ask.inv.Util;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class ComExcelUtil {
	public static Object getCellValue(XSSFCell cell) {
    	Object value = null;
    	DecimalFormat df = new DecimalFormat("0");
    	SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    	DecimalFormat nf = new DecimalFormat("0.00");
    	switch (cell.getCellType()) {
    	case XSSFCell.CELL_TYPE_STRING:
    		value = cell.getStringCellValue();
    		break;
    	case XSSFCell.CELL_TYPE_NUMERIC:
    		if("@".equals(cell.getCellStyle().getDataFormatString())){
    			value = df.format(cell.getNumericCellValue());
    		} else if("General".equals(cell.getCellStyle().getDataFormatString())){
    			value = nf.format(cell.getNumericCellValue());
    		} else {
    			value = sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
    		}
    		break;
    	case XSSFCell.CELL_TYPE_BOOLEAN:
    		value = cell.getBooleanCellValue();
    		break;
    	case XSSFCell.CELL_TYPE_FORMULA:
    		value = cell.getCellFormula();
    		break;
    	case XSSFCell.CELL_TYPE_ERROR:
    		value = cell.getErrorCellValue();
			break;
    	case XSSFCell.CELL_TYPE_BLANK:
    		value = "";
    		break;
    	default:
    		value = cell.toString();
    	}
    	return value;
    }
	
	public static void insertOneBlankColumn(XSSFSheet sheet, int insertPosition) {
		if (sheet == null || insertPosition < 0) return;
		int rows = sheet.getPhysicalNumberOfRows();
		if (rows <= 0) {
			return;
		}
		int columns = 0;
		Map<Integer, Integer> columnWidthMap = new HashMap<Integer, Integer>();
		for (int i = 0; i < rows; i++) {
			XSSFRow row = sheet.getRow(i);
			columns = row.getLastCellNum();
			for (int j = columns; j >= insertPosition + 1; j--) {
				XSSFCell cell = getSheetCell(row, j-1);
				XSSFCell newCell = getSheetCell(row, j);
				newCell.setCellStyle(cell.getCellStyle());
				
				if (!columnWidthMap.containsKey(j)) {
					columnWidthMap.put(j, sheet.getColumnWidth(j-1));
				}
				
				switch (cell.getCellType()) {
				case XSSFCell.CELL_TYPE_STRING:
					newCell.setCellValue(cell.getStringCellValue());
					break;
				case XSSFCell.CELL_TYPE_NUMERIC:
					newCell.setCellValue(cell.getNumericCellValue());
					break;
				case XSSFCell.CELL_TYPE_BOOLEAN:
					newCell.setCellValue(cell.getBooleanCellValue());
					break;
				case XSSFCell.CELL_TYPE_FORMULA:
					newCell.setCellFormula(cell.getCellFormula());
					break;
				case XSSFCell.CELL_TYPE_ERROR:
					newCell.setCellErrorValue(cell.getErrorCellValue());
					break;
				case XSSFCell.CELL_TYPE_BLANK:
					newCell.setCellValue("");
					break;
				default:
					newCell.setCellValue(cell.toString());
				}
				newCell.setCellType(cell.getCellType());
				
				row.removeCell(cell);
			}
		}
		
		for (int j:columnWidthMap.keySet()) {
			sheet.setColumnWidth(j, columnWidthMap.get(j));
		}
		columnWidthMap.clear();
		columnWidthMap = null;
	}
	
	public static XSSFRow getSheetRow(XSSFSheet sheet, int rownum) {
    	if (sheet == null) {
    		return null;
    	}
    	XSSFRow row = sheet.getRow(rownum);
    	if (row == null) {
    		row = sheet.createRow(rownum);
    	}
    	return row;
    }
	
	public static XSSFCell getSheetCell(XSSFRow row, int cellnum) {
    	if (row == null) {
    		return null;
    	}
    	XSSFCell cell = row.getCell(cellnum);
    	if (cell == null) {
    		cell = row.createCell(cellnum);
    	}
    	return cell;
    }
}
