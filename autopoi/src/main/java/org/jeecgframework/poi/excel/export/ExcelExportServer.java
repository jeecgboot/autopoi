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
package org.jeecgframework.poi.excel.export;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jeecgframework.poi.excel.annotation.ExcelTarget;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.jeecgframework.poi.excel.entity.vo.PoiBaseConstants;
import org.jeecgframework.poi.excel.export.base.ExcelExportBase;
import org.jeecgframework.poi.excel.export.styler.IExcelExportStyler;
import org.jeecgframework.poi.exception.excel.ExcelExportException;
import org.jeecgframework.poi.exception.excel.enums.ExcelExportEnum;
import org.jeecgframework.poi.util.PoiPublicUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Excel导出服务
 * 
 * @author JEECG
 * @date 2014年6月17日 下午5:30:54
 */
public class ExcelExportServer extends ExcelExportBase {

	private final static Logger LOGGER = LoggerFactory.getLogger(ExcelExportServer.class);

	// 最大行数,超过自动多Sheet
	private int MAX_NUM = 60000;

	private int createHeaderAndTitle(ExportParams entity, Sheet sheet, Workbook workbook, List<ExcelExportEntity> excelParams) {
		int rows = 0, feildWidth = getFieldWidth(excelParams);
		if (entity.getTitle() != null) {
			rows += createHeaderRow(entity, sheet, workbook, feildWidth);
		}
		rows += createTitleRow(entity, sheet, workbook, rows, excelParams);
		sheet.createFreezePane(0, rows, 0, rows);
		return rows;
	}

	/**
	 * 创建 表头改变
	 * 
	 * @param entity
	 * @param sheet
	 * @param workbook
	 * @param feildWidth
	 */
	public int createHeaderRow(ExportParams entity, Sheet sheet, Workbook workbook, int feildWidth) {
		Row row = sheet.createRow(0);
		row.setHeight(entity.getTitleHeight());
		createStringCell(row, 0, entity.getTitle(), getExcelExportStyler().getHeaderStyle(entity.getHeaderColor()), null);
		for (int i = 1; i <= feildWidth; i++) {
			createStringCell(row, i, "", getExcelExportStyler().getHeaderStyle(entity.getHeaderColor()), null);
		}
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, feildWidth));
		if (entity.getSecondTitle() != null) {
			row = sheet.createRow(1);
			row.setHeight(entity.getSecondTitleHeight());
			CellStyle style = workbook.createCellStyle();
			style.setAlignment(CellStyle.ALIGN_RIGHT);
			createStringCell(row, 0, entity.getSecondTitle(), style, null);
			for (int i = 1; i <= feildWidth; i++) {
				createStringCell(row, i, "", getExcelExportStyler().getHeaderStyle(entity.getHeaderColor()), null);
			}
			sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, feildWidth));
			return 2;
		}
		return 1;
	}

	public void createSheet(Workbook workbook, ExportParams entity, Class<?> pojoClass, Collection<?> dataSet, String[] exportFields) {
		if (LOGGER.isDebugEnabled()) {
			LOGGER.debug("Excel export start ,class is {}", pojoClass);
			LOGGER.debug("Excel version is {}", entity.getType().equals(ExcelType.HSSF) ? "03" : "07");
		}
		if (workbook == null || entity == null || pojoClass == null || dataSet == null) {
			throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
		}
		super.type = entity.getType();
		if (type.equals(ExcelType.XSSF)) {
			MAX_NUM = 1000000;
		}
		Sheet sheet = null;
		try {
			sheet = workbook.createSheet(entity.getSheetName());
		} catch (Exception e) {
			// 重复遍历,出现了重名现象,创建非指定的名称Sheet
			sheet = workbook.createSheet();
		}
		try {
			dataHanlder = entity.getDataHanlder();
			if (dataHanlder != null) {
				needHanlderList = Arrays.asList(dataHanlder.getNeedHandlerFields());
			}
			// 创建表格样式
			setExcelExportStyler((IExcelExportStyler) entity.getStyle().getConstructor(Workbook.class).newInstance(workbook));
			Drawing patriarch = sheet.createDrawingPatriarch();
			List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
			if (entity.isAddIndex()) {
				excelParams.add(indexExcelEntity(entity));
			}
			// 得到所有字段
			Field fileds[] = PoiPublicUtil.getClassFields(pojoClass);

            //---update-begin-----autor:scott------date:20191016-------for:导出字段支持自定义--------
            //支持自定义导出字段
            if (exportFields != null) {
                List<Field> list = new ArrayList<Field>(Arrays.asList(fileds));
                for (int i = 0; i < list.size(); i++) {
                    if (!Arrays.asList(exportFields).contains(list.get(i).getName())) {
                        list.remove(i);
                        i--;
                    }
                }

                if (list != null && list.size() > 0) {
                    fileds = list.toArray(new Field[0]);
                } else {
                    fileds = null;
                }
            }
            //---update-end-----autor:scott------date:20191016-------for:导出字段支持自定义--------

			ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
			String targetId = etarget == null ? null : etarget.value();
			getAllExcelField(entity.getExclusions(), targetId, fileds, excelParams, pojoClass, null);
			//update-begin-author:taoyan date:20200304 for:在此方法循环内设置一下图片磁盘目录，便于导出
			reConfigExcelExportParams(excelParams,entity);
			//update-end-author:taoyan date:20200304 for:在此方法循环内设置一下图片磁盘目录，便于导出
			int index = entity.isCreateHeadRows() ? createHeaderAndTitle(entity, sheet, workbook, excelParams) : 0;
			int titleHeight = index;
			setCellWith(excelParams, sheet);
			short rowHeight = getRowHeight(excelParams);
			setCurrentIndex(1);
			Iterator<?> its = dataSet.iterator();
			List<Object> tempList = new ArrayList<Object>();
			while (its.hasNext()) {
				Object t = its.next();
				index += createCells(patriarch, index, t, excelParams, sheet, workbook, rowHeight);
				tempList.add(t);
				if (index >= MAX_NUM)
					break;
			}
			mergeCells(sheet, excelParams, titleHeight);

			if (entity.getFreezeCol() != 0) {
				sheet.createFreezePane(entity.getFreezeCol(), 0, entity.getFreezeCol(), 0);
			}

			its = dataSet.iterator();
			for (int i = 0, le = tempList.size(); i < le; i++) {
				its.next();
				its.remove();
			}
			// 创建合计信息
			addStatisticsRow(getExcelExportStyler().getStyles(true, null), sheet);

			// 发现还有剩余list 继续循环创建Sheet
			if (dataSet.size() > 0) {
				createSheet(workbook, entity, pojoClass, dataSet,exportFields);
			}

		} catch (Exception e) {
			LOGGER.error(e.getMessage(), e);
			throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e.getCause());
		}
	}

	public void createSheetForMap(Workbook workbook, ExportParams entity, List<ExcelExportEntity> entityList, Collection<? extends Map<?, ?>> dataSet) {
		if (LOGGER.isDebugEnabled()) {
			LOGGER.debug("Excel version is {}", entity.getType().equals(ExcelType.HSSF) ? "03" : "07");
		}
		if (workbook == null || entity == null || entityList == null || dataSet == null) {
			throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
		}
		super.type = entity.getType();
		if (type.equals(ExcelType.XSSF)) {
			MAX_NUM = 1000000;
		}
		Sheet sheet = null;
		try {
			sheet = workbook.createSheet(entity.getSheetName());
		} catch (Exception e) {
			// 重复遍历,出现了重名现象,创建非指定的名称Sheet
			sheet = workbook.createSheet();
		}
		try {
			dataHanlder = entity.getDataHanlder();
			if (dataHanlder != null) {
				needHanlderList = Arrays.asList(dataHanlder.getNeedHandlerFields());
			}
			// 创建表格样式
			setExcelExportStyler((IExcelExportStyler) entity.getStyle().getConstructor(Workbook.class).newInstance(workbook));
			Drawing patriarch = sheet.createDrawingPatriarch();
			List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
			if (entity.isAddIndex()) {
				excelParams.add(indexExcelEntity(entity));
			}
			excelParams.addAll(entityList);
			sortAllParams(excelParams);
			int index = entity.isCreateHeadRows() ? createHeaderAndTitle(entity, sheet, workbook, excelParams) : 0;
			int titleHeight = index;
			setCellWith(excelParams, sheet);
			short rowHeight = getRowHeight(excelParams);
			setCurrentIndex(1);
			Iterator<?> its = dataSet.iterator();
			List<Object> tempList = new ArrayList<Object>();
			while (its.hasNext()) {
				Object t = its.next();
				index += createCells(patriarch, index, t, excelParams, sheet, workbook, rowHeight);
				tempList.add(t);
				if (index >= MAX_NUM)
					break;
			}
			if (entity.getFreezeCol() != 0) {
				sheet.createFreezePane(entity.getFreezeCol(), 0, entity.getFreezeCol(), 0);
			}

			mergeCells(sheet, excelParams, titleHeight);

			its = dataSet.iterator();
			for (int i = 0, le = tempList.size(); i < le; i++) {
				its.next();
				its.remove();
			}
			// 发现还有剩余list 继续循环创建Sheet
			if (dataSet.size() > 0) {
				createSheetForMap(workbook, entity, entityList, dataSet);
			}

		} catch (Exception e) {
			LOGGER.error(e.getMessage(), e);
			throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e.getCause());
		}
	}

	/**
	 * 创建表头
	 * 
	 * @param title
	 * @param index
	 */
	private int createTitleRow(ExportParams title, Sheet sheet, Workbook workbook, int index, List<ExcelExportEntity> excelParams) {
		Row row = sheet.createRow(index);
		int rows = getRowNums(excelParams);
		row.setHeight((short) 450);
		Row listRow = null;
		if (rows == 2) {
			listRow = sheet.createRow(index + 1);
			listRow.setHeight((short) 450);
		}
		int cellIndex = 0;
		CellStyle titleStyle = getExcelExportStyler().getTitleStyle(title.getColor());
		for (int i = 0, exportFieldTitleSize = excelParams.size(); i < exportFieldTitleSize; i++) {
			ExcelExportEntity entity = excelParams.get(i);
			//update-begin-author:taoyan date:20200319 for:建议autoPoi升级，优化数据返回List Map格式下的复合表头导出excel的体验 #873
			if(entity.isColspan()){
				List<String> subList = entity.getSubColumnList();
				if(subList==null || subList.size()==0){
					continue;
				}else{
					entity.initSubExportEntity(excelParams);
				}
			}
			//update-end-author:taoyan date:20200319 for:建议autoPoi升级，优化数据返回List Map格式下的复合表头导出excel的体验 #873
			if (StringUtils.isNotBlank(entity.getName())) {
				createStringCell(row, cellIndex, entity.getName(), titleStyle, entity);
			}
			if (entity.getList() != null) {
				List<ExcelExportEntity> sTitel = entity.getList();
				if (StringUtils.isNotBlank(entity.getName())) {
					sheet.addMergedRegion(new CellRangeAddress(index, index, cellIndex, cellIndex + sTitel.size() - 1));
				}
				for (int j = 0, size = sTitel.size(); j < size; j++) {
					createStringCell(rows == 2 ? listRow : row, cellIndex, sTitel.get(j).getName(), titleStyle, entity);
					cellIndex++;
				}
				cellIndex--;
			} else if (rows == 2) {
				createStringCell(listRow, cellIndex, "", titleStyle, entity);
				sheet.addMergedRegion(new CellRangeAddress(index, index + 1, cellIndex, cellIndex));
			}
			cellIndex++;
		}
		return rows;

	}

	/**
	 * 判断表头是只有一行还是两行
	 * 
	 * @param excelParams
	 * @return
	 */
	private int getRowNums(List<ExcelExportEntity> excelParams) {
		for (int i = 0; i < excelParams.size(); i++) {
			//update-begin-author:taoyan date:20200319 for:建议autoPoi升级，优化数据返回List Map格式下的复合表头导出excel的体验 #873
			ExcelExportEntity temp = excelParams.get(i);
			if ((temp.getList() != null || temp.isColspan()) && StringUtils.isNotBlank(temp.getName())) {
				return 2;
			}
			//update-end-author:taoyan date:20200319 for:建议autoPoi升级，优化数据返回List Map格式下的复合表头导出excel的体验 #873
		}
		return 1;
	}

	private ExcelExportEntity indexExcelEntity(ExportParams entity) {
		ExcelExportEntity exportEntity = new ExcelExportEntity();
		exportEntity.setOrderNum(0);
		exportEntity.setName(entity.getIndexName());
		exportEntity.setWidth(10);
		exportEntity.setFormat(PoiBaseConstants.IS_ADD_INDEX);
		return exportEntity;
	}

}
