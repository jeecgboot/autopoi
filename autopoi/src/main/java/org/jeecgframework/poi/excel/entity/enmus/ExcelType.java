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

/**
 * Excel 文件格式类型枚举
 * <p>用于指定导出/导入的 Excel 文件格式版本</p>
 * 
 * @author JEECG
 * @date 2014年12月29日 下午9:08:21
 */
public enum ExcelType {

	/**
	 * HSSF 格式 - Excel 97-2003 版本 (.xls)
	 * <ul>
	 *   <li>文件扩展名：.xls</li>
	 *   <li>最大行数：65,536 行（2^16）</li>
	 *   <li>最大列数：256 列（2^8）</li>
	 *   <li>适用场景：兼容老版本 Excel，数据量较小的场景</li>
	 *   <li>对应 POI 类：HSSFWorkbook</li>
	 * </ul>
	 */
	HSSF,
	
	/**
	 * XSSF 格式 - Excel 2007+ 版本 (.xlsx)
	 * <ul>
	 *   <li>文件扩展名：.xlsx</li>
	 *   <li>最大行数：1,048,576 行（2^20）</li>
	 *   <li>最大列数：16,384 列（2^14）</li>
	 *   <li>适用场景：现代 Excel 版本，大数据量导出，推荐使用</li>
	 *   <li>对应 POI 类：XSSFWorkbook</li>
	 *   <li>优势：支持更大数据量，文件压缩比更高，功能更丰富</li>
	 * </ul>
	 * <p><b>注意：</b>导出时请确保文件扩展名与格式类型匹配，避免文件损坏</p>
	 */
	XSSF;

}
