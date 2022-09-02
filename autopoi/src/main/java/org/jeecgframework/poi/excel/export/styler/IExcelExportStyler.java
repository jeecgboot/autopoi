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
package org.jeecgframework.poi.excel.export.styler;

import org.apache.poi.ss.usermodel.CellStyle;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.jeecgframework.poi.excel.entity.params.ExcelForEachParams;

/**
 * Excel导出样式接口
 * 
 * @author JEECG
 * @date 2015年1月9日 下午5:32:30
 */
public interface IExcelExportStyler {

	/**
	 * 列表头样式
	 * 
	 * @param headerColor
	 * @return
	 */
	public CellStyle getHeaderStyle(short headerColor);

	/**
	 * 标题样式
	 * 
	 * @param color
	 * @return
	 */
	public CellStyle getTitleStyle(short color);

	/**
	 * 获取样式方法
	 * 
	 * @param noneStyler
	 * @param entity
	 * @return
	 */
	public CellStyle getStyles(boolean noneStyler, ExcelExportEntity entity);
	/**
	 * 模板使用的样式设置
	 */
	public CellStyle getTemplateStyles(boolean isSingle, ExcelForEachParams excelForEachParams);

}
