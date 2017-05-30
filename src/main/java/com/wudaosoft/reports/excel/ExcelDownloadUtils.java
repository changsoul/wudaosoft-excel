/* 
 * Copyright(c)2010-2014 WUDAOSOFT.COM
 * 
 * Email:changsoul.wu@gmail.com
 * 
 * QQ:275100589
 */

package com.wudaosoft.reports.excel;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.net.URLEncoder;
import java.util.Collection;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * @author Changsoul Wu
 * 
 */
public class ExcelDownloadUtils {

	public static void downloadExcel(Collection<? extends Object> list, HttpServletRequest req, HttpServletResponse resp)
			throws IOException, RowsExceededException, WriteException, IllegalArgumentException, IllegalAccessException,
			InvocationTargetException {
		downloadExcel(list, null, req, resp);
	}

	public static void downloadExcel(Collection<? extends Object> list, String filename, HttpServletRequest req,
			HttpServletResponse resp) throws IOException, RowsExceededException, WriteException,
			IllegalArgumentException, IllegalAccessException, InvocationTargetException {
		downloadExcel(list, filename, false, req, resp);
	}

	public static void downloadExcel(Collection<? extends Object> list, String filename, boolean writeTitle,
			HttpServletRequest req, HttpServletResponse resp) throws IOException, RowsExceededException, WriteException,
			IllegalArgumentException, IllegalAccessException, InvocationTargetException {

		Excel xls = new Excel(list, writeTitle);

		if (filename == null)
			filename = xls.getTitle();

		String fileName = filename + ".xls";
		String newFileName = URLEncoder.encode(fileName, "UTF8");

		String userAgent = req.getHeader("User-Agent");

		// 如果没有UA，则默认使用IE的方式进行编码，因为毕竟IE还是占多数的
		String rtn = "filename=\"" + newFileName + "\"";

		if (userAgent != null) {

			userAgent = userAgent.toLowerCase();

			// IE浏览器，只能采用URLEncoder编码
			if (userAgent.indexOf("msie") != -1) {
				rtn = "filename=\"" + newFileName + "\"";
			}

			// Opera浏览器只能采用filename*
			else if (userAgent.indexOf("opera") != -1) {
				rtn = "filename*=UTF-8''" + newFileName;
			}

			// Safari浏览器，只能采用ISO编码的中文输出
			else if (userAgent.indexOf("safari") != -1) {
				rtn = "filename=\"" + new String(fileName.getBytes("UTF-8"), "ISO-8859-1") + "\"";
			}

			// Chrome浏览器，只能采用MimeUtility编码或ISO编码的中文输出
			else if (userAgent.indexOf("applewebkit") != -1) {
				rtn = "filename=\"" + new String(fileName.getBytes("UTF-8"), "ISO-8859-1") + "\"";
			}

			// FireFox浏览器，可以使用MimeUtility或filename*或ISO编码的中文输出
			else if (userAgent.indexOf("mozilla") != -1) {
				rtn = "filename*=UTF-8''" + newFileName;
			}

		}

		resp.reset();
		resp.setContentType("application/vnd.ms-excel;charset=utf-8");
		resp.addHeader("Content-Disposition", "attachment;" + rtn);
		// response.addHeader("Content-Length", "" + file.length());

		xls.generateExcel(resp.getOutputStream());
	}
}
