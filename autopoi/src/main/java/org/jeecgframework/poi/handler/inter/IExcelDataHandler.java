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
package org.jeecgframework.poi.handler.inter;

import java.util.Map;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
/**
 * Excel 导入导出 数据处理接口
 * 
 * @author JEECG
 * @date 2014年6月19日 下午11:59:45
 */
public interface IExcelDataHandler {

	/**
	 * 导出处理方法
	 * 
	 * @param obj
	 *            当前对象
	 * @param name
	 *            当前字段名称
	 * @param value
	 *            当前值
	 * @return
	 */
	public Object exportHandler(Object obj, String name, Object value);

	/**
	 * 获取需要处理的字段,导入和导出统一处理了, 减少书写的字段
	 * 
	 * @return
	 */
	public String[] getNeedHandlerFields();

	/**
	 * 导入处理方法 当前对象,当前字段名称,当前值
	 * 
	 * @param obj
	 *            当前对象
	 * @param name
	 *            当前字段名称
	 * @param value
	 *            当前值
	 * @return
	 */
	public Object importHandler(Object obj, String name, Object value);

	/**
	 * 设置需要处理的属性列表
	 * 
	 * @param fields
	 */
	public void setNeedHandlerFields(String[] fields);

	/**
	 * 设置Map导入,自定义 put
	 * 
	 * @param map
	 * @param originKey
	 * @param value
	 */
	public void setMapValue(Map<String, Object> map, String originKey, Object value);
	/**
	 * 获取这个字段的 Hyperlink ,07版本需要,03版本不需要
	 * @param creationHelper
	 * @param obj
	 * @param name
	 * @param value
	 * @return
	 */
	public Hyperlink getHyperlink(CreationHelper creationHelper, Object obj, String name, Object value);

}
