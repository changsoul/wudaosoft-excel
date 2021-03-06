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
	
	String trueText() default "是";
	
	String falseText() default "否";
	
	int index() default 9999;
	
	int length() default -1;
	
	int scale() default 0;
	
	RoundingMode roundingMode() default RoundingMode.HALF_UP;
	
	Align align() default Align.CENTRE;
	
}
