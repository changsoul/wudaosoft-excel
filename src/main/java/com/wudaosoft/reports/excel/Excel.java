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
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;

import jxl.Cell;
import jxl.CellView;
import jxl.format.Colour;
import jxl.format.Font;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/** 
 * @author Changsoul Wu
 * 
 */
public class Excel {
	
	private String title = "data";

	private boolean writeTitle = false;
	
	private List<Column> columns = new ArrayList<Column>(0);
	
	private Collection<? extends Object> dataList;
	
	public Excel(Collection<? extends Object> list) {
		this(list, false);
	}
	
	public Excel(Collection<? extends Object> list, boolean writeTitle) {
		this.dataList = list;
		this.writeTitle = writeTitle;
		
		processAnnotation();
	}
	
	public void setTitle(String title) {
		this.title = title;
	}
	
	public String getTitle() {
		return title;
	}
	
	public boolean generateExcel(File file) throws RowsExceededException, WriteException, IllegalArgumentException, IllegalAccessException, InvocationTargetException {
		if(file == null || dataList == null || dataList.size() == 0)
			return false;
		
		try {
			WritableWorkbook book = ExcelUtils.createWorkbook(file);
			
			processData(book);
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		}
		
		return true;
	}
	
	public boolean generateExcel(OutputStream out) throws RowsExceededException, WriteException, IllegalArgumentException, IllegalAccessException, InvocationTargetException {
		if(out == null || dataList == null || dataList.size() == 0)
			return false;
		
		try {
			WritableWorkbook book = ExcelUtils.createWorkbook(out);
			
			processData(book);
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		}
		
		return true;
	}

	private void processAnnotation() {
		
		if(dataList == null || dataList.size() == 0)
			return;
		
		Class<?> clazz = dataList.iterator().next().getClass();

		if (clazz.isAnnotationPresent(ExcelEnity.class)) {

			ExcelEnity annotation = clazz.getAnnotation(ExcelEnity.class);

			if("data".equals(title))
				title = annotation.value();
		}
		
		Field[] fields = clazz.getDeclaredFields();
		
		columns = new ArrayList<Column>(150);
		
		for(Field field : fields) {
			if(field.isAnnotationPresent(ExcelColumn.class)) {
				ExcelColumn ec = field.getAnnotation(ExcelColumn.class);
				
				Column c = new Column();
				
				c.setName(ec.name());
				c.setField(field.getName());
				c.setLength(ec.length());
				c.setIndex(ec.index());
				c.setScale(ec.scale());
				c.setIsDateColumn(ec.isDate());
				c.setIsNumber(ec.isNumber());
				c.setIsMoneyColumn(ec.isMoney());
				c.setIsHyperlink(ec.isHyperlink());
				c.setPattern(ec.pattern());
				c.setLinkText(ec.linkText());
				c.setTrueText(ec.trueText());
				c.setFalseText(ec.falseText());
				c.setRoundingMode(ec.roundingMode());
				c.setAlign(ec.align());
				
				Method m = null;
				try {
					m = clazz.getMethod(getMethodName(field));
				} catch (SecurityException e) {
				} catch (NoSuchMethodException e) {
				}
				
				if(m == null) {
					try {
						m = clazz.getMethod(field.getName());
					} catch (SecurityException e) {
					} catch (NoSuchMethodException e) {
					}
				}
				
				c.setGetMethod(m);
				columns.add(c);
			}
		}
		
		if(!columns.isEmpty()) {
			
			Collections.sort(columns, new Comparator<Column>(){

				@Override
				public int compare(Column o1, Column o2) {
					return Integer.valueOf(o1.getIndex()).compareTo(Integer.valueOf(o2.getIndex())) ;
				}
				
			});
			
		}
	}
	
	private void processData(WritableWorkbook book) throws RowsExceededException, WriteException, IllegalArgumentException, IllegalAccessException, InvocationTargetException, IOException {
		
		WritableSheet sheet = book.createSheet(title, 0);
		
		int rowIndex = 0;
		
		if(writeTitle) {
			ExcelUtils.addTitleName(rowIndex, 0, sheet, title);
			
			sheet.mergeCells(0, 0, columns.size()-1, 0);
			
			sheet.setRowView(rowIndex, 500);
			rowIndex++;
		}
		
		int[] columnsWidth = new int[columns.size()];
		
		for(int i =0; i < columns.size(); i++) {
			Column column = columns.get(i);
			ExcelUtils.addColumnName(rowIndex, i, sheet, column.getName());
			
			int width = column.getLength() > 0 ? genWidth(sheet.getCell(i, rowIndex), column.getLength()) : Math.max(columnsWidth[i], getCellWidth(sheet.getCell(i, rowIndex)));
			columnsWidth[i] = width;
		}
		
		sheet.setRowView(rowIndex, 340);
		
		rowIndex++;
		
		Iterator<?> iter = dataList.iterator();
		while (iter.hasNext()) {
			Object data = (Object) iter.next();
			
			for(int j =0; j < columns.size(); j++) {
				Column c = columns.get(j);
				
				Method m = c.getGetMethod();
				
				Object value = m.invoke(data);
				
				if(value != null) {
					if(c.isDateColumn()) {
						
						ExcelUtils.addDateCell(rowIndex, j, sheet, value, c.getPattern(), c);
					}  else if(c.isNumber()) {
						
						ExcelUtils.addNumberCell(rowIndex, j, sheet, value, c.getPattern(), c.getScale(), c);
					}  else if(c.isHyperlink()) {
						
						ExcelUtils.addHyperlinkCell(rowIndex, j, sheet, value, c.getLinkText(), c);
					}  else if(c.isMoneyColumn()) {
						
						double money = 0;
						int scale = c.getScale() > 0 ? c.getScale() : 2;
						
						if(value instanceof BigDecimal) {
							money = ((BigDecimal)value).setScale(scale, c.getRoundingMode()).doubleValue();
						}else {
							money = new BigDecimal(value.toString()).setScale(scale, c.getRoundingMode()).doubleValue();
						}
						
						ExcelUtils.addMoneyCell(rowIndex, j, sheet, money, c);
					} else if(value instanceof Boolean) {
						
						boolean b = ((Boolean)value).booleanValue();
						
						if(b) {
							ExcelUtils.addStringCell(rowIndex, j, sheet, c.getTrueText(), c);
						}
						else{
							ExcelUtils.addStringCell(rowIndex, j, sheet, c.getFalseText(), c, Colour.RED);
						}
						
						//ExcelUtils.addStringCell(rowIndex, j, sheet, v, c);
					} else {
						ExcelUtils.addStringCell(rowIndex, j, sheet, value.toString(), c);
					}
				}else {
					ExcelUtils.addStringCell(rowIndex, j, sheet, "", c);
				}
				
				columnsWidth[j] = Math.max(columnsWidth[j], getCellWidth(sheet.getCell(j, rowIndex)));
			}
			
			rowIndex++;
		}
		
		for(int i =0; i < columns.size(); i++) {
			CellView cv = new CellView();
			cv.setSize(columnsWidth[i] + 512);
			sheet.setColumnView(i, cv);
		}
		
		book.write();
        book.close();
	}
	
	private int getCellWidth(Cell cell) {
		Font defaultFont = WritableWorkbook.NORMAL_STYLE.getFont();
		String contents = cell.getContents();
		// Font font = cell.getCellFormat().getFont();
		//
		// Font activeFont = font.equals(defaultFont) ? columnFont : font;
		Font activeFont = cell.getCellFormat().getFont();

		int pointSize = activeFont.getPointSize();
		int numChars = contents.length();

		if (activeFont.isItalic() || activeFont.getBoldWeight() > 400) {
			numChars += 2;
		}

		int points = numChars * pointSize;

		return (int) ((points * 256) / defaultFont.getPointSize());
	}
	
	private int genWidth(Cell cell, int fontLength) {
		
		return (int) ((fontLength * cell.getCellFormat().getFont().getPointSize() * 256) / WritableWorkbook.NORMAL_STYLE.getFont().getPointSize());
	}
	
	private String getMethodName(Field field) {
		byte[] items = field.getName().getBytes();
		items[0] = (byte) ((char) items[0] - 'a' + 'A');
		return "get" + new String(items);
	}

}
