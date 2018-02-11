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

		String rtn = "filename=\"" + newFileName + "\"";

		if (userAgent != null) {

			userAgent = userAgent.toLowerCase();

			if (userAgent.indexOf("msie") != -1) {
				rtn = "filename=\"" + newFileName + "\"";
			}

			else if (userAgent.indexOf("opera") != -1) {
				rtn = "filename*=UTF-8''" + newFileName;
			}
			else if (userAgent.indexOf("safari") != -1) {
				rtn = "filename=\"" + new String(fileName.getBytes("UTF-8"), "ISO-8859-1") + "\"";
			}
			else if (userAgent.indexOf("applewebkit") != -1) {
				rtn = "filename=\"" + new String(fileName.getBytes("UTF-8"), "ISO-8859-1") + "\"";
			}
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
