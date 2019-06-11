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
package org.jeecgframework.poi.excel.entity.sax;

import org.jeecgframework.poi.excel.entity.enmus.CellValueType;

/**
 * Cell 对象
 * 
 * @author JEECG
 * @date 2014年12月29日 下午10:12:57
 */
public class SaxReadCellEntity {
	/**
	 * 值类型
	 */
	private CellValueType cellType;
	/**
	 * 值
	 */
	private Object value;

	public SaxReadCellEntity(CellValueType cellType, Object value) {
		this.cellType = cellType;
		this.value = value;
	}

	public CellValueType getCellType() {
		return cellType;
	}

	public void setCellType(CellValueType cellType) {
		this.cellType = cellType;
	}

	public Object getValue() {
		return value;
	}

	public void setValue(Object value) {
		this.value = value;
	}

	@Override
	public String toString() {
		return "[type=" + cellType.toString() + ",value=" + value + "]";
	}

}
