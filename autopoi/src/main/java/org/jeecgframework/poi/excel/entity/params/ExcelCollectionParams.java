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
package org.jeecgframework.poi.excel.entity.params;

import java.util.Map;

/**
 * Excel 对于的 Collection
 * 
 * @author JEECG
 * @date 2013-9-26
 * @version 1.0
 */
public class ExcelCollectionParams {

	/**
	 * 集合对应的名称
	 */
	private String name;
	/**
	 * Excel 列名称
	 */
	private String excelName;
	/**
	 * 实体对象
	 */
	private Class<?> type;
	/**
	 * 这个list下面的参数集合实体对象
	 */
	private Map<String, ExcelImportEntity> excelParams;

	public Map<String, ExcelImportEntity> getExcelParams() {
		return excelParams;
	}

	public String getName() {
		return name;
	}

	public Class<?> getType() {
		return type;
	}

	public void setExcelParams(Map<String, ExcelImportEntity> excelParams) {
		this.excelParams = excelParams;
	}

	public void setName(String name) {
		this.name = name;
	}

	public void setType(Class<?> type) {
		this.type = type;
	}

	public String getExcelName() {
		return excelName;
	}

	public void setExcelName(String excelName) {
		this.excelName = excelName;
	}
}
