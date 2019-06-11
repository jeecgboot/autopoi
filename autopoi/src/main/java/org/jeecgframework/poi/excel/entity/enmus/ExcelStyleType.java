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
package org.jeecgframework.poi.excel.entity.enmus;

import org.jeecgframework.poi.excel.export.styler.ExcelExportStylerBorderImpl;
import org.jeecgframework.poi.excel.export.styler.ExcelExportStylerColorImpl;
import org.jeecgframework.poi.excel.export.styler.ExcelExportStylerDefaultImpl;

/**
 * 插件提供的几个默认样式
 * 
 * @author JEECG
 * @date 2015年1月9日 下午9:02:24
 */
public enum ExcelStyleType {

	NONE("默认样式", ExcelExportStylerDefaultImpl.class), BORDER("边框样式", ExcelExportStylerBorderImpl.class), COLOR("间隔行样式", ExcelExportStylerColorImpl.class);

	private String name;
	private Class<?> clazz;

	ExcelStyleType(String name, Class<?> clazz) {
		this.name = name;
		this.clazz = clazz;
	}

	public Class<?> getClazz() {
		return clazz;
	}

	public String getName() {
		return name;
	}

}
