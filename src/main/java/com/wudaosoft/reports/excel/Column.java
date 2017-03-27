/* 
 * Copyright(c)2010-2014 WUDAOSOFT.COM
 * 
 * Email:changsoul.wu@gmail.com
 * 
 * QQ:275100589
 */ 
 
package com.wudaosoft.reports.excel;

import java.lang.reflect.Method;
import java.math.RoundingMode;

/** 
 * @author Changsoul Wu
 * 
 */
public class Column {

	private String name;
	
	private String field;
	
	private String pattern;
	
	private String linkText;
	
	private String trueText;
	
	private String falseText;
	
	private Method getMethod;
	
	private int length = -1;
	
	private int index = 9999;
	
	private int scale = 0;
	
	private boolean isDateColumn = false;
	
	private boolean isMoneyColumn = false;
	
	private boolean isNumber = false;
	
	private boolean isHyperlink = false;
	
	private RoundingMode roundingMode = RoundingMode.HALF_UP;
	
	private Align align = Align.CENTRE;;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getField() {
		return field;
	}

	public void setField(String field) {
		this.field = field;
	}

	public String getPattern() {
		return pattern;
	}

	public void setPattern(String pattern) {
		this.pattern = pattern;
	}

	public Method getGetMethod() {
		return getMethod;
	}

	public void setGetMethod(Method getMethod) {
		this.getMethod = getMethod;
	}

	public int getLength() {
		return length;
	}

	public void setLength(int length) {
		this.length = length;
	}

	public int getIndex() {
		return index;
	}

	public void setIndex(int index) {
		this.index = index;
	}

	public boolean isDateColumn() {
		return isDateColumn;
	}

	public void setIsDateColumn(boolean isDateColumn) {
		this.isDateColumn = isDateColumn;
	}

	public boolean isMoneyColumn() {
		return isMoneyColumn;
	}

	public void setIsMoneyColumn(boolean isMoneyColumn) {
		this.isMoneyColumn = isMoneyColumn;
	}

	public boolean isHyperlink() {
		return isHyperlink;
	}

	public void setIsHyperlink(boolean isHyperlink) {
		this.isHyperlink = isHyperlink;
	}

	public String getLinkText() {
		return linkText;
	}

	public void setLinkText(String linkText) {
		this.linkText = linkText;
	}

	public String getTrueText() {
		return trueText;
	}

	public void setTrueText(String trueText) {
		this.trueText = trueText;
	}

	public String getFalseText() {
		return falseText;
	}

	public void setFalseText(String falseText) {
		this.falseText = falseText;
	}

	public int getScale() {
		return scale;
	}

	public void setScale(int scale) {
		this.scale = scale;
	}

	public boolean isNumber() {
		return isNumber;
	}

	public void setIsNumber(boolean isNumber) {
		this.isNumber = isNumber;
	}

	public RoundingMode getRoundingMode() {
		return roundingMode;
	}

	public void setRoundingMode(RoundingMode roundingMode) {
		this.roundingMode = roundingMode;
	}

	public Align getAlign() {
		return align;
	}

	public void setAlign(Align align) {
		this.align = align;
	}
	
}
