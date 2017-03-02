package utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileExportUtils {
	public static final int CELL_TYPE_DATE = 0;
	public static final int CELL_TYPE_CURRENCY = 1;
	
	public static final int CELL_STYLE_BOLD = 100;

	// Constants for production request export template
	public static final CellReference PRODUCTION_REQUEST_ID_REF = new CellReference("H3");
	public static final CellReference PRODUCTION_REQUEST_PROJECT_ID_REF = new CellReference("D4");
	public static final CellReference PRODUCTION_REQUEST_PROJECT_NAME_REF = new CellReference("F4");
	public static final CellReference PRODUCTION_REQUEST_SUBMITTED_DATE_REF = new CellReference("D5");
	public static final CellReference PRODUCTION_REQUEST_DUE_DATE_REF = new CellReference("F5");
	public static final CellReference PRODUCTION_REQUEST_USE_DATE_REF = new CellReference("H5");
	public static final CellReference PRODUCTION_REQUEST_CLIENT_NAME_REF = new CellReference("D6");
	public static final CellReference PRODUCTION_REQUEST_CLIENT_TYPE_REF = new CellReference("F6");
	public static final CellReference PRODUCTION_REQUEST_PRIORITY_REF = new CellReference("H6");
	public static final CellReference PRODUCTION_REQUEST_COMMENTS = new CellReference("D13");
	public static final CellReference PRODUCTION_REQUEST_ITEMS_START_REF = new CellReference("B10");
	public static final CellReference PRODUCTION_REQUEST_RECIPIENT_NAME_REF = new CellReference("D11");
	public static final CellReference PRODUCTION_REQUEST_RECIPIENT_CONTACT_REF = new CellReference("F11");
	public static final CellReference PRODUCTION_REQUEST_DELIVERY_METHOD_REF = new CellReference("H11");
	public static final CellReference PRODUCTION_REQUEST_RECIPIENT_ORG_REF = new CellReference("D12");
	public static final CellReference PRODUCTION_REQUEST_RECIPIENT_ADDRESS_REF = new CellReference("F12");

	public static boolean isRowEmpty(XSSFRow row) {
		return (row == null);
	}

	public static boolean isCellEmpty(XSSFCell cell) {
		if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
			return true;
		}

		if (cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getStringCellValue().isEmpty()) {
			return true;
		}

		return false;
	}

	// shortcut for inserting new row
	public static void insertRow(XSSFSheet sheet, String... args) {
		XSSFRow row;
		if (sheet.getPhysicalNumberOfRows() == 0) {
			row = sheet.createRow(sheet.getLastRowNum());
		} else {
			row = sheet.createRow(sheet.getLastRowNum() + 1);
		}
		
		setRow(row, args);
	}
	
	// shortcut for setting row values
	public static void setRow(XSSFRow row, String... args) {
		int columnCounter = 0;
		for (String arg : args) {
			XSSFCell cell = row.createCell(columnCounter);
			cell.setCellValue(arg);
			columnCounter++;
		}
	}
	
	// shortcut for inserting a new cell
	public static XSSFCell insertCell(XSSFRow row, int columnCounter, String value) {
		XSSFCell cell = row.createCell(columnCounter);
		cell.setCellValue(value);
		return cell;
	}
	
//	// shortcut for inserting new row for a table only
//	public static void insertTableRow(XSSFSheet sheet, String... args) {
//		XSSFRow row;
//		if (sheet.getPhysicalNumberOfRows() == 0) {
//			row = sheet.createRow(sheet.getLastRowNum());
//		} else {
//			row = sheet.createRow(sheet.getLastRowNum() + 1);
//		}
//		
//		setTableRow(row, args);
//	}
	
//	// shortcut for setting row values
//	public static void setTableRow(XSSFRow row, String... args) {
//		int columnCounter = 0;
//		for (String arg : args) {
//			XSSFCell cell = row.createCell(columnCounter);
//			cell.setCellValue(arg);
//			columnCounter++;
//		}
//	}
	
	public static String sanitizeForFilename(String s) {
		return s.replaceAll("[^a-zA-Z0-9.-]", "_");
	}
	
	// autosize wont work until we install asian fonts. 
	public static void autoSizeColumns(XSSFSheet sheet) {
		// assumes that header includes all columns
		XSSFRow row = sheet.getRow(0);
		if (row != null) {
			int columnCount = row.getLastCellNum();
			for(int i = 0; i <= columnCount; i++) {
				sheet.autoSizeColumn(i);
			}
		}
	}
	
	public static void setRowFormat(XSSFRow row, int format) {
		if (row == null) return;
		
	    for(int i = 0; i < row.getLastCellNum(); i++){
	    	setCellFormat(row.getCell(i), format);
	    }
	}
	
	public static void setCellFormat(XSSFCell cell, int format) {
		if (cell == null) return;
		XSSFWorkbook workbook = cell.getRow().getSheet().getWorkbook();
		
		CellStyle cellStyle = workbook.createCellStyle();
		CreationHelper createHelper = workbook.getCreationHelper();
		
		switch (format) {
			case CELL_TYPE_DATE:
				cellStyle.setDataFormat(
				    createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
				break;
			case CELL_TYPE_CURRENCY:
				break;
			case CELL_STYLE_BOLD:
				Font font = workbook.createFont();
			    font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			    cellStyle.setFont(font);
				break;
			default:
				break;
		}
		
		cell.setCellStyle(cellStyle);
	}
	
	public static CellStyle createHyperlinkStyle(XSSFWorkbook workbook) {
		CellStyle hlink_style = workbook.createCellStyle();
		Font hlink_font = workbook.createFont();
		hlink_font.setUnderline(Font.U_SINGLE);
		hlink_font.setColor(IndexedColors.BLUE.getIndex());
		hlink_style.setFont(hlink_font);
		return hlink_style;
	}
	
}
