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
package org.jeecgframework.poi.excel.imports.sax.parse;

import java.util.List;

import org.jeecgframework.poi.excel.entity.sax.SaxReadCellEntity;

public interface ISaxRowRead {
	/**
	 * 获取返回数据
	 * 
	 * @param <T>
	 * @return
	 */
	public <T> List<T> getList();

	/**
	 * 解析数据
	 * 
	 * @param index
	 * @param datas
	 */
	public void parse(int index, List<SaxReadCellEntity> datas);

	//update-begin---author:lvdandan ---date:20261020  for：[issues/189]SAX导入用IExcelReadRowHanlder回调时，最后一行数据丢失------------
	/**
	 * 一个sheet的所有行都解析完成时调用，用于处理还没有输出的最后一条数据。
	 * 默认不做处理，已有的自定义实现不需要改动。
	 */
	public default void endSheet() {
	}
	//update-end---author:lvdandan ---date:20261020  for：[issues/189]SAX导入用IExcelReadRowHanlder回调时，最后一行数据丢失------------

}
