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

}
