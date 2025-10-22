/**
 * Copyright 2013-2015 JEECG (jeecgos@163.com)
 *   
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.jeecgframework.poi.excel.view;

import javax.servlet.http.HttpServletRequest;

import org.springframework.web.servlet.view.AbstractView;

/**
 * 基础抽象Excel View
 * 
 * @author JEECG
 * @date 2015年2月28日 下午1:41:05
 */
public abstract class MiniAbstractExcelView extends AbstractView {

	private static final String CONTENT_TYPE = "application/vnd.ms-excel";

	protected static final String HSSF = ".xls";
	protected static final String XSSF = ".xlsx";

	public MiniAbstractExcelView() {
		setContentType(CONTENT_TYPE);
	}

	protected boolean isIE(HttpServletRequest request) {
		return (request.getHeader("USER-AGENT").toLowerCase().indexOf("msie") > 0 || request.getHeader("USER-AGENT").toLowerCase().indexOf("rv:11.0") > 0) ? true : false;
	}

}
