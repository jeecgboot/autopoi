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
package org.jeecgframework.poi.handler.impl;

import java.util.Map;

import org.jeecgframework.poi.handler.inter.IExcelDataHandler;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
/**
 * 数据处理默认实现,返回空
 * 
 * @author JEECG
 * @date 2014年6月20日 上午12:11:52
 */
public abstract class ExcelDataHandlerDefaultImpl implements IExcelDataHandler {
	/**
	 * 需要处理的字段
	 */
	private String[] needHandlerFields;

	@Override
	public Object exportHandler(Object obj, String name, Object value) {
		return value;
	}

	@Override
	public String[] getNeedHandlerFields() {
		return needHandlerFields;
	}

	@Override
	public Object importHandler(Object obj, String name, Object value) {
		return value;
	}

	@Override
	public void setNeedHandlerFields(String[] needHandlerFields) {
		this.needHandlerFields = needHandlerFields;
	}

	@Override
	public void setMapValue(Map<String, Object> map, String originKey, Object value) {
		map.put(originKey, value);
	}

	@Override
	public Hyperlink getHyperlink(CreationHelper creationHelper, Object obj, String name, Object value) {
		return null;
	}
}
