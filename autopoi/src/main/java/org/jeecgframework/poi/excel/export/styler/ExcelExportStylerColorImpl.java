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

import org.apache.poi.ss.usermodel.*;

/**
 * 带有样式的导出服务
 * 
 * @author JEECG
 * @date 2015年1月9日 下午4:54:15
 */
public class ExcelExportStylerColorImpl extends AbstractExcelExportStyler implements IExcelExportStyler {

	public ExcelExportStylerColorImpl(Workbook workbook) {
		super.createStyles(workbook);
	}

	@Override
	public CellStyle getHeaderStyle(short headerColor) {
		CellStyle titleStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setFontHeightInPoints((short) 24);
		titleStyle.setFont(font);
		titleStyle.setFillForegroundColor(headerColor);
		titleStyle.setAlignment(HorizontalAlignment.CENTER);
		titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		return titleStyle;
	}

	@Override
	public CellStyle stringNoneStyle(Workbook workbook, boolean isWarp) {
		CellStyle style = workbook.createCellStyle();
		style.setBorderLeft(BorderStyle.THIN); // 左边框
		style.setBorderRight(BorderStyle.THIN); // 右边框
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setDataFormat(STRING_FORMAT);
		if (isWarp) {
			style.setWrapText(true);
		}
		return style;
	}

	@Override
	public CellStyle getTitleStyle(short color) {
		CellStyle titleStyle = workbook.createCellStyle();
		titleStyle.setFillForegroundColor(color); // 填充的背景颜色
		titleStyle.setAlignment(HorizontalAlignment.CENTER);
		titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); // 填充图案
		titleStyle.setWrapText(true);
		return titleStyle;
	}

	@Override
	public CellStyle stringSeptailStyle(Workbook workbook, boolean isWarp) {
		CellStyle style = workbook.createCellStyle();
		style.setBorderLeft(BorderStyle.THIN); // 左边框
		style.setBorderRight(BorderStyle.THIN); // 右边框
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setFillForegroundColor((short) 41); // 填充的背景颜色
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND); // 填充图案
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setDataFormat(STRING_FORMAT);
		if (isWarp) {
			style.setWrapText(true);
		}
		return style;
	}

}
