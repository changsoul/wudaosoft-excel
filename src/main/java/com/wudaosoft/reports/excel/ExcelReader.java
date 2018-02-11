/**
 *    Copyright 2009-2018 Wudao Software Studio(wudaosoft.com)
 * 
 *    Licensed under the Apache License, Version 2.0 (the "License");
 *    you may not use this file except in compliance with the License.
 *    You may obtain a copy of the License at
 * 
 *        http://www.apache.org/licenses/LICENSE-2.0
 * 
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 */
package com.wudaosoft.reports.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 * @author changsoul.wu
 *
 */
public class ExcelReader {
	
	private List<Column> columns = new ArrayList<Column>(0);
	private Map<Object, Column> columnMap = new HashMap<Object, Column>();
	
	public <T> List<T> readExcel(String filePath, Class<T> clazz) throws BiffException, IOException, InstantiationException, IllegalAccessException  {
		return readExcel(new FileInputStream(filePath), clazz);
	}
	
	public <T> List<T> readExcel(File file, Class<T> clazz) throws BiffException, IOException, InstantiationException, IllegalAccessException  {
		return readExcel(new FileInputStream(file), clazz);
	}

	public <T> List<T> readExcel(InputStream stream, Class<T> clazz) throws BiffException, IOException, InstantiationException, IllegalAccessException  {

		columns.clear();
		columnMap.clear();
		
		// 获取Excel文件对象
		Workbook rwb = Workbook.getWorkbook(stream);
		// 获取文件的指定工作表 默认的第一个
		Sheet sheet = rwb.getSheet(0);
		// 行数(表头的目录不需要，从1开始)
		
		int rows = sheet.getRows();
		int cols = sheet.getColumns();
		
		if(rows <= 1 || cols == 0)
			return Collections.emptyList();
		
		processAnnotation(clazz);
		
		for(int i = 0; i < cols; i++) {
			
			Cell cell = sheet.getCell(i, 0);
			
			if(cell == null || cell.getContents() == null)
				continue;
			
			for(Column column : columns) {
				
				if(column.getName().equals(cell.getContents().trim()))
					columnMap.put(i, column);
			}
		}
		
		List<T> list = new ArrayList<T>();
		
		for (int i = 1; i < rows; i++) {

			T obj = clazz.newInstance();
			
			// 列数
			for (int j = 0; j < cols; j++) {
				Column column = columnMap.get(j);
				
				if(column == null)
					continue;
				
				Cell cell = sheet.getCell(j, i);
				
				String val = "";
				if(cell != null)
					val = cell.getContents();
				
				try {
					
					Method method = column.getSetMethod();
					
					if(method == null)
						continue;
					
					String type = method.getParameterTypes()[0].getName();
					
					if (type.equals("java.lang.String")) {
						column.getSetMethod().invoke(obj, val);
					} else if(type.equals("java.lang.Integer")) {
						column.getSetMethod().invoke(obj, Integer.valueOf(val.trim()));
					} else if(type.equals("java.lang.Float")) {
						column.getSetMethod().invoke(obj, Float.valueOf(val.trim()));
					} else if(type.equals("java.lang.Double")) {
						column.getSetMethod().invoke(obj, Double.valueOf(val.trim()));
					} else if(type.equals("java.lang.Long")) {
						column.getSetMethod().invoke(obj, Long.valueOf(val.trim()));
					} else if(type.equals("java.lang.Short")) {
						column.getSetMethod().invoke(obj, Short.valueOf(val.trim()));
					} else if(type.equals("java.util.Date")) {
						column.getSetMethod().invoke(obj, new SimpleDateFormat(column.getPattern()).parse(val.trim()));
					} else if(type.equals("java.math.BigDecimal")) {
						column.getSetMethod().invoke(obj, new BigDecimal(val.trim()));
					} else {
						column.getSetMethod().invoke(obj, val);
					}
				} catch (IllegalArgumentException e) {
				} catch (InvocationTargetException e) {
				} catch (NullPointerException e) {
				} catch (ParseException e) {
				}
			}
			list.add(obj);
		}
		return list;
	}

	private void processAnnotation(Class<?> clazz) {
		
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
					m = clazz.getMethod(getSetMethodName(field));
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
				
				c.setSetMethod(m);
				columns.add(c);
			}
		}
	}

	private String getSetMethodName(Field field) {
		byte[] items = field.getName().getBytes();
		items[0] = (byte) ((char) items[0] - 'a' + 'A');
		return "set" + new String(items);
	}
}
