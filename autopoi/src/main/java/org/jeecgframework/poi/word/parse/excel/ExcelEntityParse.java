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
package org.jeecgframework.poi.word.parse.excel;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.jeecgframework.poi.excel.annotation.ExcelTarget;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.jeecgframework.poi.excel.export.base.ExportBase;
import org.jeecgframework.poi.exception.word.WordExportException;
import org.jeecgframework.poi.exception.word.enmus.WordExportEnum;
import org.jeecgframework.poi.util.PoiPublicUtil;
import org.jeecgframework.poi.word.entity.params.ExcelListEntity;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 解析实体类对象 复用注解
 * 
 * @author JEECG
 * @date 2014年8月9日 下午10:30:57
 */
public class ExcelEntityParse extends ExportBase {

	private static final Logger LOGGER = LoggerFactory.getLogger(ExcelEntityParse.class);

	private static void checkExcelParams(ExcelListEntity entity) {
		if (entity.getList() == null || entity.getClazz() == null) {
			throw new WordExportException(WordExportEnum.EXCEL_PARAMS_ERROR);
		}

	}

	private int createCells(int index, Object t, List<ExcelExportEntity> excelParams, XWPFTable table, short rowHeight) throws Exception {
		ExcelExportEntity entity;
		XWPFTableRow row = table.createRow();
		row.setHeight(rowHeight);
		int maxHeight = 1, cellNum = 0;
		for (int k = 0, paramSize = excelParams.size(); k < paramSize; k++) {
			entity = excelParams.get(k);
			if (entity.getList() != null) {
				Collection<?> list = (Collection<?>) entity.getMethod().invoke(t, new Object[] {});
				int listC = 0;
				for (Object obj : list) {
					createListCells(index + listC, cellNum, obj, entity.getList(), table);
					listC++;
				}
				cellNum += entity.getList().size();
				if (list != null && list.size() > maxHeight) {
					maxHeight = list.size();
				}
			} else {
				Object value = getCellValue(entity, t);
				if (entity.getType() == 1) {
					setCellValue(row, value, cellNum++);
				}
			}
		}
		// 合并需要合并的单元格
		cellNum = 0;
		for (int k = 0, paramSize = excelParams.size(); k < paramSize; k++) {
			entity = excelParams.get(k);
			if (entity.getList() != null) {
				cellNum += entity.getList().size();
			} else if (entity.isNeedMerge()) {
				table.setCellMargins(index, index + maxHeight - 1, cellNum, cellNum);
				cellNum++;
			}
		}
		return maxHeight;
	}

	/**
	 * 创建List之后的各个Cells
	 * 
	 * @param styles
	 */
	public void createListCells(int index, int cellNum, Object obj, List<ExcelExportEntity> excelParams, XWPFTable table) throws Exception {
		ExcelExportEntity entity;
		XWPFTableRow row;
		if (table.getRow(index) == null) {
			row = table.createRow();
			row.setHeight(getRowHeight(excelParams));
		} else {
			row = table.getRow(index);
		}
		for (int k = 0, paramSize = excelParams.size(); k < paramSize; k++) {
			entity = excelParams.get(k);
			Object value = getCellValue(entity, obj);
			if (entity.getType() == 1) {
				setCellValue(row, value, cellNum++);
			}
		}
	}

	/**
	 * 获取表头数据
	 * 
	 * @param table
	 * @param index
	 * @return
	 */
	private Map<String, Integer> getTitleMap(XWPFTable table, int index, int headRows) {
		if (index < headRows) {
			throw new WordExportException(WordExportEnum.EXCEL_NO_HEAD);
		}
		Map<String, Integer> map = new HashMap<String, Integer>();
		String text;
		for (int j = 0; j < headRows; j++) {
			List<XWPFTableCell> cells = table.getRow(index - j - 1).getTableCells();
			for (int i = 0; i < cells.size(); i++) {
				text = cells.get(i).getText();
				if (StringUtils.isEmpty(text)) {
					throw new WordExportException(WordExportEnum.EXCEL_HEAD_HAVA_NULL);
				}
				map.put(text, i);
			}
		}
		return map;
	}

	/**
	 * 解析上一行并生成更多行
	 * 
	 * @param table
	 * @param i
	 * @param listobj
	 */
	public void parseNextRowAndAddRow(XWPFTable table, int index, ExcelListEntity entity) {
		checkExcelParams(entity);
		// 获取表头数据
		Map<String, Integer> titlemap = getTitleMap(table, index, entity.getHeadRows());
		try {
			// 得到所有字段
			Field fileds[] = PoiPublicUtil.getClassFields(entity.getClazz());
			ExcelTarget etarget = entity.getClazz().getAnnotation(ExcelTarget.class);
			String targetId = null;
			if (etarget != null) {
				targetId = etarget.value();
			}
			// 获取实体对象的导出数据
			List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
			getAllExcelField(null, targetId, fileds, excelParams, entity.getClazz(), null);
			// 根据表头进行筛选排序
			sortAndFilterExportField(excelParams, titlemap);
			short rowHeight = getRowHeight(excelParams);
			Iterator<?> its = entity.getList().iterator();
			while (its.hasNext()) {
				Object t = its.next();
				index += createCells(index, t, excelParams, table, rowHeight);
			}
		} catch (Exception e) {
			LOGGER.error(e.getMessage(), e);
		}
	}

	private void setCellValue(XWPFTableRow row, Object value, int cellNum) {
		if (row.getCell(cellNum++) != null) {
			row.getCell(cellNum - 1).setText(value == null ? "" : value.toString());
		} else {
			row.createCell().setText(value == null ? "" : value.toString());
		}
	}

	/**
	 * 对导出序列进行排序和塞选
	 * 
	 * @param excelParams
	 * @param titlemap
	 * @return
	 */
	private void sortAndFilterExportField(List<ExcelExportEntity> excelParams, Map<String, Integer> titlemap) {
		for (int i = excelParams.size() - 1; i >= 0; i--) {
			if (excelParams.get(i).getList() != null && excelParams.get(i).getList().size() > 0) {
				sortAndFilterExportField(excelParams.get(i).getList(), titlemap);
				if (excelParams.get(i).getList().size() == 0) {
					excelParams.remove(i);
				} else {
					excelParams.get(i).setOrderNum(i);
				}
			} else {
				if (titlemap.containsKey(excelParams.get(i).getName())) {
					excelParams.get(i).setOrderNum(i);
				} else {
					excelParams.remove(i);
				}
			}
		}
		sortAllParams(excelParams);
	}

}
