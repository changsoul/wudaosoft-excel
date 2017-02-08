/* 
 * Copyright(c)2010-2014 WUDAOSOFT.COM
 * 
 * Email:changsoul.wu@gmail.com
 * 
 * QQ:275100589
 */

package com.wudaosoft.reports.excel;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.math.RoundingMode;

/**
 * @author Changsoul Wu
 * 
 */
@Documented
@Target(value = { ElementType.FIELD, ElementType.METHOD })
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {
	
	String name() default "";
	
	boolean isDate() default false;
	
	boolean isMoney() default false;
	
	boolean isNumber() default false;
	
	boolean isHyperlink() default false;
	
	String pattern() default "";
	
	String linkText() default "";
	
	int index() default 9999;
	
	int length() default -1;
	
	int scale() default 0;
	
	RoundingMode roundingMode() default RoundingMode.HALF_UP;
	
	Align align() default Align.CENTRE;
	
}
