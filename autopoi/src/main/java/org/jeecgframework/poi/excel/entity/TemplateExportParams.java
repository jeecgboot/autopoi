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
package org.jeecgframework.poi.excel.entity;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jeecgframework.poi.excel.export.styler.ExcelExportStylerDefaultImpl;

import java.io.IOException;
import java.io.InputStream;

/**
 * 模板导出参数设置
 * 
 * @author JEECG
 * @date 2013-10-17
 * @version 1.0
 */
public class TemplateExportParams extends ExcelBaseParams {

	/**
	 * 输出全部的sheet
	 */
	private boolean scanAllsheet = false;
	/**
	 * 模板的路径
	 */
	private String templateUrl;
	/**
	 * 模板
	 */
	private Workbook templateWb;

	/**
	 * 需要导出的第几个 sheetNum,默认是第0个
	 */
	private Integer[] sheetNum = new Integer[] { 0 };

	/**
	 * 这只sheetName 不填就使用原来的
	 */
	private String[] sheetName;

	/**
	 * 表格列标题行数,默认1
	 */
	private int headingRows = 1;

	/**
	 * 表格列标题开始行,默认1
	 */
	private int headingStartRow = 1;
	/**
	 * 设置数据源的NUM
	 */
	private int dataSheetNum = 0;
	/**
	 * Excel 导出style
	 */
	private Class<?> style = ExcelExportStylerDefaultImpl.class;
	/**
	 * FOR EACH 用到的局部变量
	 */
	private String tempParams = "t";
    //列循环支持
	private boolean   colForEach      = false;

	/**
	 * 默认构造器
	 */
	public TemplateExportParams() {

	}

	/**
	 * 构造器
	 * 
	 * @param templateUrl
	 *            模板路径
	 * @param scanAllsheet
	 *            是否输出全部的sheet
	 * @param sheetName
	 *            sheet的名称,可不填
	 */
	public TemplateExportParams(String templateUrl, boolean scanAllsheet, String... sheetName) {
		this.templateUrl = templateUrl;
		this.scanAllsheet = scanAllsheet;
		if (sheetName != null && sheetName.length > 0) {
			this.sheetName = sheetName;

		}
	}

	/**
	 * 构造器
	 * 
	 * @param templateUrl
	 *            模板路径
	 * @param sheetNum
	 *            sheet 的位置,可不填
	 */
	public TemplateExportParams(String templateUrl, Integer... sheetNum) {
		this.templateUrl = templateUrl;
		if (sheetNum != null && sheetNum.length > 0) {
			this.sheetNum = sheetNum;
		}
	}

	/**
	 * 单个sheet输出构造器
	 * 
	 * @param templateUrl
	 *            模板路径
	 * @param sheetName
	 *            sheet的名称
	 * @param sheetNum
	 *            sheet的位置,可不填
	 */
	public TemplateExportParams(String templateUrl, String sheetName, Integer... sheetNum) {
		this.templateUrl = templateUrl;
		this.sheetName = new String[] { sheetName };
		if (sheetNum != null && sheetNum.length > 0) {
			this.sheetNum = sheetNum;
		}
	}
	//update-begin-author:liusq---date:2024-09-03--for: [issues/7048]TemplateExportParams类建议增加传入模板文件InputStream的方式
	/**
	 * 构造器
	 * @param inputStream 输入流
	 * @param scanAllsheet 是否输出全部的sheet
	 * @param sheetName    sheet的名称,可不填
	 */
	public TemplateExportParams(InputStream inputStream, boolean scanAllsheet, String... sheetName) throws IOException {
		this.templateWb = WorkbookFactory.create(inputStream);
		this.scanAllsheet = scanAllsheet;
		if (sheetName != null && sheetName.length > 0) {
			this.sheetName = sheetName;
		}
	}
	/**
	 * 构造器
	 * @param inputStream 输入流
	 * @param sheetNum    sheet 的位置,可不填
	 */
	public TemplateExportParams(InputStream inputStream, Integer... sheetNum) throws IOException {
		this.templateWb = WorkbookFactory.create(inputStream);
		if (sheetNum != null && sheetNum.length > 0) {
			this.sheetNum = sheetNum;
		}
	}

	/**
	 * 单个sheet输出构造器
	 * @param inputStream 输入流
	 * @param sheetName   sheet的名称
	 * @param sheetNum    sheet的位置,可不填
	 */
	public TemplateExportParams(InputStream inputStream, String sheetName, Integer... sheetNum) throws IOException {
		this.templateWb = WorkbookFactory.create(inputStream);
		this.sheetName = new String[] { sheetName };
		if (sheetNum != null && sheetNum.length > 0) {
			this.sheetNum = sheetNum;
		}
	}
	//update-end-author:liusq---date:2024-09-03--for: [issues/7048]TemplateExportParams类建议增加传入模板文件InputStream的方式
	public int getHeadingRows() {
		return headingRows;
	}

	public int getHeadingStartRow() {
		return headingStartRow;
	}

	public String[] getSheetName() {
		return sheetName;
	}

	public Integer[] getSheetNum() {
		return sheetNum;
	}

	public String getTemplateUrl() {
		return templateUrl;
	}

	public void setHeadingRows(int headingRows) {
		this.headingRows = headingRows;
	}

	public void setHeadingStartRow(int headingStartRow) {
		this.headingStartRow = headingStartRow;
	}

	public void setSheetName(String[] sheetName) {
		this.sheetName = sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = new String[] { sheetName };
	}

	public void setSheetNum(Integer[] sheetNum) {
		this.sheetNum = sheetNum;
	}

	public void setSheetNum(Integer sheetNum) {
		this.sheetNum = new Integer[] { sheetNum };
	}

	public void setTemplateUrl(String templateUrl) {
		this.templateUrl = templateUrl;
	}

	public Class<?> getStyle() {
		return style;
	}

	public void setStyle(Class<?> style) {
		this.style = style;
	}

	public int getDataSheetNum() {
		return dataSheetNum;
	}

	public void setDataSheetNum(int dataSheetNum) {
		this.dataSheetNum = dataSheetNum;
	}

	public boolean isScanAllsheet() {
		return scanAllsheet;
	}

	public void setScanAllsheet(boolean scanAllsheet) {
		this.scanAllsheet = scanAllsheet;
	}

	public String getTempParams() {
		return tempParams;
	}

	public void setTempParams(String tempParams) {
		this.tempParams = tempParams;
	}

	public boolean isColForEach() {
		return colForEach;
	}

	public void setColForEach(boolean colForEach) {
		this.colForEach = colForEach;
	}
	public Workbook getTemplateWb() {
		return templateWb;
	}

	public void setTemplateWb(Workbook templateWb) {
		this.templateWb = templateWb;
	}
}
