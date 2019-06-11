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
package org.jeecgframework.poi.excel.entity.result;

import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * 导入返回类
 * 
 * @author JEECG
 * @date 2014年6月29日 下午5:12:10
 */
public class ExcelImportResult<T> {

	/**
	 * 结果集
	 */
	private List<T> list;

	/**
	 * 是否存在校验失败
	 */
	private boolean verfiyFail;

	/**
	 * 数据源
	 */
	private Workbook workbook;

	public ExcelImportResult() {

	}

	public ExcelImportResult(List<T> list, boolean verfiyFail, Workbook workbook) {
		this.list = list;
		this.verfiyFail = verfiyFail;
		this.workbook = workbook;
	}

	public List<T> getList() {
		return list;
	}

	public Workbook getWorkbook() {
		return workbook;
	}

	public boolean isVerfiyFail() {
		return verfiyFail;
	}

	public void setList(List<T> list) {
		this.list = list;
	}

	public void setVerfiyFail(boolean verfiyFail) {
		this.verfiyFail = verfiyFail;
	}

	public void setWorkbook(Workbook workbook) {
		this.workbook = workbook;
	}

}
