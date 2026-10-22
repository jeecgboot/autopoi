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

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.commons.lang3.StringUtils;
import org.jeecgframework.poi.excel.annotation.ExcelTarget;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.jeecgframework.poi.excel.entity.params.ExcelCollectionParams;
import org.jeecgframework.poi.excel.entity.params.ExcelImportEntity;
import org.jeecgframework.poi.excel.entity.sax.SaxReadCellEntity;
import org.jeecgframework.poi.excel.imports.CellValueServer;
import org.jeecgframework.poi.excel.imports.base.ImportBaseService;
import org.jeecgframework.poi.exception.excel.ExcelImportException;
import org.jeecgframework.poi.handler.inter.IExcelReadRowHanlder;
import org.jeecgframework.poi.util.PoiPublicUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.collect.Lists;

/**
 * 当行读取数据
 * 
 * @author JEECG
 * @param <T>
 * @date 2015年1月1日 下午7:59:39
 */
@SuppressWarnings({ "rawtypes", "unchecked" })
public class SaxRowRead extends ImportBaseService implements ISaxRowRead {

	private static final Logger LOGGER = LoggerFactory.getLogger(SaxRowRead.class);
	/** 需要返回的数据 **/
	private List list;
	/** 导出的对象 **/
	private Class<?> pojoClass;
	/** 导入参数 **/
	private ImportParams params;
	/** 列表头对应关系 **/
	private Map<Integer, String> titlemap = new HashMap<Integer, String>();
	/** 当前的对象 **/
	private Object object = null;

	private Map<String, ExcelImportEntity> excelParams = new HashMap<String, ExcelImportEntity>();

	private List<ExcelCollectionParams> excelCollection = new ArrayList<ExcelCollectionParams>();

	private String targetId;

	private CellValueServer cellValueServer;

	private IExcelReadRowHanlder hanlder;

	public SaxRowRead(Class<?> pojoClass, ImportParams params, IExcelReadRowHanlder hanlder) {
		list = Lists.newArrayList();
		this.params = params;
		this.pojoClass = pojoClass;
		cellValueServer = new CellValueServer();
		this.hanlder = hanlder;
		initParams(pojoClass, params);
	}

	private void initParams(Class<?> pojoClass, ImportParams params) {
		//update-begin---author:lvdandan ---date:20261020  for：[issues/140]SAX导入以Map.class作为目标类型时，每行返回一个Map，没有@Excel字段需要解析------------
		if (Map.class.equals(pojoClass)) {
			return;
		}
		//update-end---author:lvdandan ---date:20261020  for：[issues/140]SAX导入以Map.class作为目标类型时，每行返回一个Map，没有@Excel字段需要解析------------
		try {

			Field fileds[] = PoiPublicUtil.getClassFields(pojoClass);
			ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
			if (etarget != null) {
				targetId = etarget.value();
			}
			getAllExcelField(targetId, fileds, excelParams, excelCollection, pojoClass, null);
			//update-begin---author:lvdandan ---date:20261020  for：[issues/189]SAX导入fixedIndex不生效，没有表头行(headRows=0)时也要生效------------
			applyFixedIndex();
			//update-end---author:lvdandan ---date:20261020  for：[issues/189]SAX导入fixedIndex不生效，没有表头行(headRows=0)时也要生效------------
		} catch (Exception e) {
			LOGGER.error(e.getMessage(), e);
			throw new ExcelImportException(e.getMessage());
		}

	}

	@Override
	public <T> List<T> getList() {
		return list;
	}

	//update-begin---author:lvdandan ---date:20261020  for：[issues/189]SAX导入用IExcelReadRowHanlder回调时，最后一行数据丢失------------
	@Override
	public void endSheet() {
		// 回调模式下，一条数据要等读到下一条才会输出，sheet读完后需要把最后一条补输出
		if (object != null && hanlder != null) {
			hanlder.hanlder(object);
		}
		// 下一个sheet的第一行不应该被当成上一个sheet最后一条数据的集合元素
		object = null;
	}
	//update-end---author:lvdandan ---date:20261020  for：[issues/189]SAX导入用IExcelReadRowHanlder回调时，最后一行数据丢失------------

	@Override
	public void parse(int index, List<SaxReadCellEntity> datas) {
		try {
			if (datas == null || datas.size() == 0) {
				return;
			}
			// 标题行跳过
			if (index < params.getTitleRows()) {
				return;
			}
			// 表头行
			if (index < params.getTitleRows() + params.getHeadRows()) {
				addHeadData(datas);
			} else if (index < params.getTitleRows() + params.getHeadRows() + params.getStartRows()) {
				//update-begin---author:lvdandan ---date:20261020  for：[issues/137]ImportParams.startRows 不生效，表头后需要跳过startRows行------------
				return;
				//update-end---author:lvdandan ---date:20261020  for：[issues/137]ImportParams.startRows 不生效，表头后需要跳过startRows行------------
			} else {
				addListData(datas);
			}
		} catch (Exception e) {
			LOGGER.error(e.getMessage(), e);
			throw new ExcelImportException(e.getMessage());
		}
	}

	//update-begin---author:lvdandan ---date:20261020  for：[issues/140]SAX导入以Map.class作为目标类型时，每行返回一个Map，key是表头，value是单元格的值------------
	/**
	 * 目标类型是 Map：每一行生成一个 Map，key 是表头，value 是单元格的值（与 DOM 导入一致）。
	 * 按列顺序放入 LinkedHashMap，每个表头都有对应的 key，空单元格的值是 null。
	 * 值保持 SAX 读到的类型：文本是 String，数字是 BigDecimal，日期是 Date。
	 * Map 没有一对多的子行，每一行都是一条新数据，所以直接输出。
	 *
	 * @param datas
	 */
	private void addMapData(List<SaxReadCellEntity> datas) {
		Map<String, Object> map = new LinkedHashMap<String, Object>();
		for (Map.Entry<Integer, String> title : new TreeMap<Integer, String>(titlemap).entrySet()) {
			int col = title.getKey();
			SaxReadCellEntity cell = col < datas.size() ? datas.get(col) : null;
			map.put(title.getValue(), cell == null ? null : cell.getValue());
		}
		if (hanlder != null) {
			hanlder.hanlder(map);
		} else {
			list.add(map);
		}
	}
	//update-end---author:lvdandan ---date:20261020  for：[issues/140]SAX导入以Map.class作为目标类型时，每行返回一个Map，key是表头，value是单元格的值------------

	/**
	 * 集合元素处理
	 *
	 * @param datas
	 */
	private void addListData(List<SaxReadCellEntity> datas) throws Exception {
		//update-begin---author:lvdandan ---date:20261020  for：[issues/140]SAX导入以Map.class作为目标类型时，每行返回一个Map，key是表头，value是单元格的值------------
		if (Map.class.equals(pojoClass)) {
			addMapData(datas);
			return;
		}
		//update-end---author:lvdandan ---date:20261020  for：[issues/140]SAX导入以Map.class作为目标类型时，每行返回一个Map，key是表头，value是单元格的值------------
		// 判断是集合元素还是不是集合元素,如果是就继续加入这个集合,不是就创建新的对象
		//update-begin---author:lvdandan ---date:20261020  for：[issues/190]SAX导入空单元格用null占位，主键列为空(含该行没有这一列)时是集合元素------------
		// 与DOM导入一致：只有实体含集合(一对多)时，主键列为空才是上一条数据的集合元素，否则是一条新数据
		if (excelCollection.size() > 0 && isEmptyCell(datas, params.getKeyIndex()) && object != null) {
		//update-end---author:lvdandan ---date:20261020  for：[issues/190]SAX导入空单元格用null占位，主键列为空(含该行没有这一列)时是集合元素------------
			for (ExcelCollectionParams param : excelCollection) {
				addListContinue(object, param, datas, titlemap, targetId, params);
			}
		} else {
			if (object != null && hanlder != null) {
				hanlder.hanlder(object);
			}
			object = PoiPublicUtil.createObject(pojoClass, targetId);
			SaxReadCellEntity entity;
			for (int i = 0, le = datas.size(); i < le; i++) {
				entity = datas.get(i);
				String titleString = (String) titlemap.get(i);
				//update-begin---author:lvdandan ---date:20261020  for：[issues/190]SAX导入空单元格用null占位，跳过空单元格------------
				if (entity != null && excelParams.containsKey(titleString)) {
				//update-end---author:lvdandan ---date:20261020  for：[issues/190]SAX导入空单元格用null占位，跳过空单元格------------
					saveFieldValue(params, object, entity, excelParams, titleString);
				}
			}
			for (ExcelCollectionParams param : excelCollection) {
				addListContinue(object, param, datas, titlemap, targetId, params);
			}
			if (hanlder == null) {
				list.add(object);
			}
		}

	}

	/***
	 * 向List里面继续添加元素
	 * 
	 * @param exclusions
	 * @param object
	 * @param param
	 * @param datas
	 * @param titlemap
	 * @param targetId
	 * @param params
	 */
	private void addListContinue(Object object, ExcelCollectionParams param, List<SaxReadCellEntity> datas, Map<Integer, String> titlemap, String targetId, ImportParams params) throws Exception {
		Collection collection = (Collection) PoiPublicUtil.getMethod(param.getName(), object.getClass()).invoke(object, new Object[] {});
		Object entity = PoiPublicUtil.createObject(param.getType(), targetId);
		boolean isUsed = false;// 是否需要加上这个对象
		for (int i = 0; i < datas.size(); i++) {
			String titleString = (String) titlemap.get(i);
			//update-begin---author:lvdandan ---date:20261020  for：[issues/190]SAX导入空单元格用null占位，跳过空单元格------------
			if (datas.get(i) != null && param.getExcelParams().containsKey(titleString)) {
			//update-end---author:lvdandan ---date:20261020  for：[issues/190]SAX导入空单元格用null占位，跳过空单元格------------
				saveFieldValue(params, entity, datas.get(i), param.getExcelParams(), titleString);
				isUsed = true;
			}
		}
		if (isUsed) {
			collection.add(entity);
		}
	}

	/**
	 * 设置值
	 * 
	 * @param params
	 * @param object
	 * @param entity
	 * @param excelParams
	 * @param titleString
	 * @throws Exception
	 */
	private void saveFieldValue(ImportParams params, Object object, SaxReadCellEntity entity, Map<String, ExcelImportEntity> excelParams, String titleString) throws Exception {
		Object value = cellValueServer.getValue(params.getDataHanlder(), object, entity, excelParams, titleString);
		setValues(excelParams.get(titleString), object, value);
	}

	/**
	 * put 表头数据
	 * 
	 * @param datas
	 */
	//update-begin---author:lvdandan ---date:20261020  for：[issues/190]SAX导入空单元格用null占位------------
	/**
	 * 指定列的单元格是否为空：该行没有这一列、该列是空单元格(null占位)、或值为空字符串
	 */
	private boolean isEmptyCell(List<SaxReadCellEntity> datas, int index) {
		if (index < 0 || index >= datas.size() || datas.get(index) == null) {
			return true;
		}
		return StringUtils.isEmpty(String.valueOf(datas.get(index).getValue()));
	}
	//update-end---author:lvdandan ---date:20261020  for：[issues/190]SAX导入空单元格用null占位------------

	private void addHeadData(List<SaxReadCellEntity> datas) {
		for (int i = 0; i < datas.size(); i++) {
			//update-begin---author:lvdandan ---date:20261020  for：[issues/190]SAX导入空单元格用null占位，跳过空单元格------------
			if (datas.get(i) != null && StringUtils.isNotEmpty(String.valueOf(datas.get(i).getValue()))) {
			//update-end---author:lvdandan ---date:20261020  for：[issues/190]SAX导入空单元格用null占位，跳过空单元格------------
				titlemap.put(i, String.valueOf(datas.get(i).getValue()));
			}
		}
		//update-begin---author:lvdandan ---date:20261020  for：[issues/189]SAX导入fixedIndex不生效，表头文字不能覆盖固定列的映射------------
		applyFixedIndex();
		//update-end---author:lvdandan ---date:20261020  for：[issues/189]SAX导入fixedIndex不生效，表头文字不能覆盖固定列的映射------------
	}

	//update-begin---author:lvdandan ---date:20261020  for：[issues/189]SAX导入fixedIndex不生效------------
	/**
	 * 把 @Excel(fixedIndex = n) 的字段固定映射到第 n 列，不依赖表头文字（与 DOM 导入一致）
	 */
	private void applyFixedIndex() {
		for (String key : excelParams.keySet()) {
			if (key.startsWith("FIXED_")) {
				titlemap.put(Integer.parseInt(key.substring("FIXED_".length())), key);
			}
		}
	}
	//update-end---author:lvdandan ---date:20261020  for：[issues/189]SAX导入fixedIndex不生效------------
}
