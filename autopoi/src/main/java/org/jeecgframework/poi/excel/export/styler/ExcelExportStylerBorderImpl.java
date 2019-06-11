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
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 带有边框的Excel样式
 * 
 * @author JEECG
 * @date 2015年1月9日 下午5:55:29
 */
public class ExcelExportStylerBorderImpl extends AbstractExcelExportStyler implements IExcelExportStyler {

	public ExcelExportStylerBorderImpl(Workbook workbook) {
		super.createStyles(workbook);
	}

	@Override
	public CellStyle getHeaderStyle(short color) {
		CellStyle titleStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setFontHeightInPoints((short) 12);
		titleStyle.setFont(font);
		titleStyle.setBorderLeft((short) 1); // 左边框
		titleStyle.setBorderRight((short) 1); // 右边框
		titleStyle.setBorderBottom((short) 1);
		titleStyle.setBorderTop((short) 1);
		titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
		titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		return titleStyle;
	}

	@Override
	public CellStyle stringNoneStyle(Workbook workbook, boolean isWarp) {
		CellStyle style = workbook.createCellStyle();
		style.setBorderLeft((short) 1); // 左边框
		style.setBorderRight((short) 1); // 右边框
		style.setBorderBottom((short) 1);
		style.setBorderTop((short) 1);
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setDataFormat(STRING_FORMAT);
		if (isWarp) {
			style.setWrapText(true);
		}
		return style;
	}

	@Override
	public CellStyle getTitleStyle(short color) {
		CellStyle titleStyle = workbook.createCellStyle();
		titleStyle.setBorderLeft((short) 1); // 左边框
		titleStyle.setBorderRight((short) 1); // 右边框
		titleStyle.setBorderBottom((short) 1);
		titleStyle.setBorderTop((short) 1);
		titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
		titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		titleStyle.setWrapText(true);
		return titleStyle;
	}

	@Override
	public CellStyle stringSeptailStyle(Workbook workbook, boolean isWarp) {
		return isWarp ? stringNoneWrapStyle : stringNoneStyle;
	}

}
