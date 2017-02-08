/* 
 * Copyright(c)2010-2014 WUDAOSOFT.COM
 * 
 * Email:changsoul.wu@gmail.com
 * 
 * QQ:275100589
 */ 
 
package com.wudaosoft.reports.excel;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.text.SimpleDateFormat;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.NumberFormat;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableHyperlink;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/** 
 * @author Changsoul Wu
 * 
 */
public class ExcelUtils {
	
	static final int TITLE_FONT_SIZE = 11;
	
	static final int DEFAULT_FONT_SIZE = 10;
	
	static final Colour TITLE_COLOR = Colour.BLUE2;
	
	static final Colour COLUMN_NAME_COLOR = Colour.YELLOW2;
	
	static final Colour CONTENT_COLOR = Colour.ORANGE;
	
	static final String DEFAULT_DATE_PATTERN = "yyyy-MM-dd HH:mm:ss";
	
	static final String DEFAULT_MONEY_FORMAT  = "#,##0.00";
	
	public static void addTitleName(int row, int column, WritableSheet sheet, String text) throws RowsExceededException, WriteException {
		
		
		WritableCellFormat cf = new WritableCellFormat(new WritableFont(WritableFont.createFont("宋体"), TITLE_FONT_SIZE, WritableFont.BOLD));
		
		cf.setAlignment(Alignment.CENTRE);
		cf.setVerticalAlignment(VerticalAlignment.CENTRE);
		
		cf.setBackground(TITLE_COLOR);
		cf.setBorder(Border.ALL, BorderLineStyle.THIN);
		
		Label label = new Label(column, row, text, cf);
		
		sheet.addCell(label);
	}
	
	public static void addColumnName(int row, int column, WritableSheet sheet, String text) throws RowsExceededException, WriteException {
		
		
		WritableCellFormat cf = new WritableCellFormat(new WritableFont(WritableFont.createFont("宋体"), TITLE_FONT_SIZE, WritableFont.BOLD));
		
		cf.setAlignment(Alignment.CENTRE);
		cf.setVerticalAlignment(VerticalAlignment.CENTRE);
		
		cf.setBackground(COLUMN_NAME_COLOR);
		cf.setBorder(Border.ALL, BorderLineStyle.THIN);
		
		Label label = new Label(column, row, text, cf);
		
//		CellView cv = new CellView();
//		cv.setAutosize(true);
//		sheet.setColumnView(column, cv);
		
		sheet.addCell(label);
	}
	
	public static void addStringCell(int row, int column, WritableSheet sheet, String text, Column c) throws RowsExceededException, WriteException {
		
		addStringCell(row, column, sheet, text, c, null);
	}
	
	public static void addStringCell(int row, int column, WritableSheet sheet, String text, Column c, Colour color) throws RowsExceededException, WriteException {
		
		WritableCellFormat cf = null;
		
		if(color == null) {
			cf = new WritableCellFormat(new WritableFont(WritableFont.createFont("宋体"), DEFAULT_FONT_SIZE));
		}else {
			cf = new WritableCellFormat(new WritableFont(WritableFont.createFont("宋体"), DEFAULT_FONT_SIZE, WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE, color));
		}
		
		cf.setAlignment(Alignment.getAlignment(c.getAlign().ordinal()));
		cf.setVerticalAlignment(VerticalAlignment.CENTRE);
		
		cf.setBackground(CONTENT_COLOR);
		cf.setBorder(Border.ALL, BorderLineStyle.THIN);
		
		Label label = new Label(column, row, text, cf);
		
		sheet.addCell(label);
	}
	
	public static void addMoneyCell(int row, int column, WritableSheet sheet, double value, Column c) throws RowsExceededException, WriteException {
		
		NumberFormat nf = new NumberFormat(DEFAULT_MONEY_FORMAT);
		
		WritableCellFormat cf = new WritableCellFormat(new WritableFont(WritableFont.createFont("宋体"), DEFAULT_FONT_SIZE), nf);
		
		cf.setBackground(CONTENT_COLOR);
		cf.setBorder(Border.ALL, BorderLineStyle.THIN);
		cf.setAlignment(Alignment.getAlignment(c.getAlign().ordinal()));
		cf.setVerticalAlignment(VerticalAlignment.CENTRE);
		
		jxl.write.Number number = new jxl.write.Number(column, row, value, cf);
		
		sheet.addCell(number);
	}

	public static void addNumberCell(int row, int column, WritableSheet sheet, Object value, String pattern,
			int scale, Column c) throws RowsExceededException, WriteException {
		
		String ptt = "0";
		
		if(pattern != null && !"".equals(pattern)) {
			ptt = pattern;
		}else if(scale > 0) {
			ptt = genScale(scale);
		}
		
		NumberFormat nf = new NumberFormat(ptt);
		
		WritableCellFormat cf = new WritableCellFormat(new WritableFont(WritableFont.createFont("宋体"), DEFAULT_FONT_SIZE), nf);
		
		cf.setBackground(CONTENT_COLOR);
		cf.setBorder(Border.ALL, BorderLineStyle.THIN);
		cf.setAlignment(Alignment.getAlignment(c.getAlign().ordinal()));
		cf.setVerticalAlignment(VerticalAlignment.CENTRE);
		
		jxl.write.Number number = new jxl.write.Number(column, row, Double.valueOf(value.toString().trim()).doubleValue(), cf);
		
		sheet.addCell(number);
	}
	
//	public static void addDateCell(int row, int column, WritableSheet sheet, Date value, String pattern) throws RowsExceededException, WriteException {
//		if(pattern == null || pattern.equals(""))
//			pattern = DEFAULT_DATE_PATTERN;
//		
//		jxl.write.DateFormat df = new jxl.write.DateFormat(pattern);
//		
//		WritableCellFormat cf = new WritableCellFormat(DEFAULT_FONT, df);
//		
//		cf.setBackground(CONTENT_COLOR);
//		cf.setBorder(Border.ALL, BorderLineStyle.THIN);
//		cf.setAlignment(Alignment.RIGHT);
//		cf.setVerticalAlignment(VerticalAlignment.CENTRE);
//		
//		jxl.write.DateTime cell = new jxl.write.DateTime(column, row, value, cf);
//		
//		sheet.addCell(cell);
//	}
	
	public static void addDateCell(int row, int column, WritableSheet sheet, Object value, String pattern, Column c) throws RowsExceededException, WriteException {
		if(pattern == null || pattern.equals(""))
			pattern = DEFAULT_DATE_PATTERN;
		
		SimpleDateFormat sdf = new SimpleDateFormat(pattern);
		
		addStringCell(row, column, sheet, sdf.format(value), c);
	}
	
	public static void addBooleanCell(int row, int column, WritableSheet sheet, boolean value, Column c) throws RowsExceededException, WriteException {
//		if(pattern == null || pattern.equals(""))
//			pattern = DEFAULT_DATE_PATTERN;
//		
//		jxl.write.DateFormat df = new jxl.write.DateFormat(pattern);
		
		WritableCellFormat cf = new WritableCellFormat(new WritableFont(WritableFont.createFont("宋体"), DEFAULT_FONT_SIZE));
		
		cf.setBackground(CONTENT_COLOR);
		cf.setBorder(Border.ALL, BorderLineStyle.THIN);
		cf.setAlignment(Alignment.getAlignment(c.getAlign().ordinal()));
		cf.setVerticalAlignment(VerticalAlignment.CENTRE);
		
		jxl.write.Boolean cell = new jxl.write.Boolean(column, row, value, cf);
		
		sheet.addCell(cell);
	}
	
	public static void addHyperlinkCell(int row, int column, WritableSheet sheet, Object value, String linkText, Column c) {
		
		if(linkText == null || linkText.equals(""))
			linkText = value.toString();
		
		try {
			
			//sheet.mergeCells(column, row, column + 1, row);
			
			WritableHyperlink whl = new WritableHyperlink(column, row, new URL(value.toString()));
			whl.setDescription(linkText);
			sheet.addHyperlink(whl);
			
			WritableCellFormat cf = new WritableCellFormat(new WritableFont(WritableFont.createFont("宋体"), DEFAULT_FONT_SIZE));
			
			cf.setAlignment(Alignment.CENTRE);
			cf.setVerticalAlignment(VerticalAlignment.CENTRE);
			
			cf.setBackground(CONTENT_COLOR);
			cf.setBorder(Border.ALL, BorderLineStyle.THIN);
			
			sheet.getWritableCell(column, row).setCellFormat(cf);
		} catch (MalformedURLException e) {
		} catch (RowsExceededException e) {
		} catch (WriteException e) {
		}
	}
	
	public static WritableWorkbook createWorkbook(File file) throws IOException {
		WritableWorkbook book = Workbook.createWorkbook(file);
		
		book.setColourRGB(Colour.BLUE2, 219, 229, 241);
		book.setColourRGB(Colour.YELLOW2, 184, 204, 228);
		book.setColourRGB(Colour.ORANGE, 242, 242, 242);
		return book;
	}
	
	public static WritableWorkbook createWorkbook(OutputStream out) throws IOException {
		WritableWorkbook book = Workbook.createWorkbook(out);
		
		book.setColourRGB(Colour.BLUE2, 219, 229, 241);
		book.setColourRGB(Colour.YELLOW2, 184, 204, 228);
		book.setColourRGB(Colour.ORANGE, 242, 242, 242);
		return book;
	}

	public static String genScale(int scale) {
		if(scale < 1)
			return "";
		
		StringBuilder sb = new StringBuilder("0.");
		
		for(int i = 0; i < scale; i++) {
			sb.append("0");
		}
		
		return sb.toString();
	}
}
