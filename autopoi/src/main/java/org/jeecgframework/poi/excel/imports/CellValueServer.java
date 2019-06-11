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
package org.jeecgframework.poi.excel.imports;

import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.sql.Time;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.jeecgframework.poi.excel.entity.params.ExcelImportEntity;
import org.jeecgframework.poi.excel.entity.sax.SaxReadCellEntity;
import org.jeecgframework.poi.exception.excel.ExcelImportException;
import org.jeecgframework.poi.exception.excel.enums.ExcelImportEnum;
import org.jeecgframework.poi.handler.inter.IExcelDataHandler;
import org.jeecgframework.poi.util.ExcelUtil;
import org.jeecgframework.poi.util.PoiPublicUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Cell 取值服务 判断类型处理数据 1.判断Excel中的类型 2.根据replace替换值 3.handler处理数据 4.判断返回类型转化数据返回
 * 
 * @author JEECG
 * @date 2014年6月26日 下午10:42:28
 */
public class CellValueServer {

	private static final Logger LOGGER = LoggerFactory.getLogger(CellValueServer.class);

	private List<String> hanlderList = null;

	/**
	 * 获取单元格内的值
	 * 
	 * @param xclass
	 * @param cell
	 * @param entity
	 * @return
	 */
	private Object getCellValue(String xclass, Cell cell, ExcelImportEntity entity) {
		if (cell == null) {
			return "";
		}
		Object result = null;
		// 日期格式比较特殊,和cell格式不一致
		if ("class java.util.Date".equals(xclass) || ("class java.sql.Time").equals(xclass)) {
			if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
				// 日期格式
				result = cell.getDateCellValue();
			} else {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				result = getDateData(entity, cell.getStringCellValue());
			}
			if (("class java.sql.Time").equals(xclass)) {
				result = new Time(((Date) result).getTime());
			}
		} else if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
			result = cell.getNumericCellValue();
		} else if (Cell.CELL_TYPE_BOOLEAN == cell.getCellType()) {
			result = cell.getBooleanCellValue();
		} else {
			result = cell.getStringCellValue();
		}
		return result;
	}

	/**
	 * 获取日期类型数据
	 * 
	 * @Author JEECG
	 * @date 2013年11月26日
	 * @param entity
	 * @param value
	 * @return
	 */
	private Date getDateData(ExcelImportEntity entity, String value) {
		if (StringUtils.isNotEmpty(entity.getFormat()) && StringUtils.isNotEmpty(value)) {
			SimpleDateFormat format = new SimpleDateFormat(entity.getFormat());
			try {
				return format.parse(value);
			} catch (ParseException e) {
				LOGGER.error("时间格式化失败,格式化:{},值:{}", entity.getFormat(), value);
				throw new ExcelImportException(ExcelImportEnum.GET_VALUE_ERROR);
			}
		}
		return null;
	}

	/**
	 * 获取cell的值
	 * 
	 * @param object
	 * @param excelParams
	 * @param cell
	 * @param titleString
	 */
	public Object getValue(IExcelDataHandler dataHanlder, Object object, Cell cell, Map<String, ExcelImportEntity> excelParams, String titleString) throws Exception {
		ExcelImportEntity entity = excelParams.get(titleString);
		String xclass = "class java.lang.Object";
		if (!(object instanceof Map)) {
			Method setMethod = entity.getMethods() != null && entity.getMethods().size() > 0 ? entity.getMethods().get(entity.getMethods().size() - 1) : entity.getMethod();
			Type[] ts = setMethod.getGenericParameterTypes();
			xclass = ts[0].toString();
		}
		Object result = getCellValue(xclass, cell, entity);
		if (entity != null) {
			result = hanlderSuffix(entity.getSuffix(), result);
			//update-begin-author:taoYan date:20180807 for:多值替换
			result = replaceValue(entity.getReplace(), result,entity.isMultiReplace());
			//update-end-author:taoYan date:20180807 for:多值替换
		}
		result = hanlderValue(dataHanlder, object, result, titleString);
		return getValueByType(xclass, result, entity);
	}

	/**
	 * 获取cell值
	 * 
	 * @param dataHanlder
	 * @param object
	 * @param entity
	 * @param excelParams
	 * @param titleString
	 * @return
	 */
	public Object getValue(IExcelDataHandler dataHanlder, Object object, SaxReadCellEntity cellEntity, Map<String, ExcelImportEntity> excelParams, String titleString) {
		ExcelImportEntity entity = excelParams.get(titleString);
		Method setMethod = entity.getMethods() != null && entity.getMethods().size() > 0 ? entity.getMethods().get(entity.getMethods().size() - 1) : entity.getMethod();
		Type[] ts = setMethod.getGenericParameterTypes();
		String xclass = ts[0].toString();
		Object result = cellEntity.getValue();
		result = hanlderSuffix(entity.getSuffix(), result);
		//update-begin-auhtor:taoyan date:20180807 for:多值替换
		result = replaceValue(entity.getReplace(), result,entity.isMultiReplace());
		//update-end-auhtor:taoyan date:20180807 for:多值替换
		result = hanlderValue(dataHanlder, object, result, titleString);
		return getValueByType(xclass, result, entity);
	}

	/**
	 * 把后缀删除掉
	 * 
	 * @param result
	 * @param suffix
	 * @return
	 */
	private Object hanlderSuffix(String suffix, Object result) {
		if (StringUtils.isNotEmpty(suffix) && result != null && result.toString().endsWith(suffix)) {
			String temp = result.toString();
			return temp.substring(0, temp.length() - suffix.length());
		}
		return result;
	}

	/**
	 * 根据返回类型获取返回值
	 * 
	 * @param xclass
	 * @param result
	 * @param entity
	 * @return
	 */
	private Object getValueByType(String xclass, Object result, ExcelImportEntity entity) {
		try {
			//update-begin-author:scott date:20180711 for:TASK #2950 【bug】excel 导入报错，空指针问题
			if(result==null || "".equals(String.valueOf(result))){
				return null;
			}
			//update-end-author:scott date:20180711 for:TASK #2950 【bug】excel 导入报错，空指针问题
			if ("class java.util.Date".equals(xclass)) {
				return result;
			}
			if ("class java.lang.Boolean".equals(xclass) || "boolean".equals(xclass)) {
				return Boolean.valueOf(String.valueOf(result));
			}
			if ("class java.lang.Double".equals(xclass) || "double".equals(xclass)) {
				return Double.valueOf(String.valueOf(result));
			}
			if ("class java.lang.Long".equals(xclass) || "long".equals(xclass)) {
				return Long.valueOf(ExcelUtil.remove0Suffix(String.valueOf(result)));
			}
			if ("class java.lang.Float".equals(xclass) || "float".equals(xclass)) {
				return Float.valueOf(String.valueOf(result));
			}
			if ("class java.lang.Integer".equals(xclass) || "int".equals(xclass)) {
				return Integer.valueOf(ExcelUtil.remove0Suffix(String.valueOf(result)));
			}
			if ("class java.math.BigDecimal".equals(xclass)) {
				return new BigDecimal(String.valueOf(result));
			}
			if ("class java.lang.String".equals(xclass)) {
				// 针对String 类型,但是Excel获取的数据却不是String,比如Double类型,防止科学计数法
				if (result instanceof String) {
					return result;
				}
				// double类型防止科学计数法
				if (result instanceof Double) {
					return PoiPublicUtil.doubleToString((Double) result);
				}
				return String.valueOf(result);
			}
			return result;
		} catch (Exception e) {
			LOGGER.error(e.getMessage(), e);
			throw new ExcelImportException(ExcelImportEnum.GET_VALUE_ERROR);
		}
	}

	/**
	 * 调用处理接口处理值
	 * 
	 * @param dataHanlder
	 * @param object
	 * @param result
	 * @param titleString
	 * @return
	 */
	private Object hanlderValue(IExcelDataHandler dataHanlder, Object object, Object result, String titleString) {
		if (dataHanlder == null || dataHanlder.getNeedHandlerFields() == null || dataHanlder.getNeedHandlerFields().length == 0) {
			return result;
		}
		if (hanlderList == null) {
			hanlderList = Arrays.asList(dataHanlder.getNeedHandlerFields());
		}
		if (hanlderList.contains(titleString)) {
			return dataHanlder.importHandler(object, titleString, result);
		}
		return result;
	}

	//update-begin-author:taoyan date:20180807 for:导入多值替换--
	/**
	 * 导入支持多值替换
	 * @param replace 数据库中字典查询出来的数组
	 * @param result excel单元格获取的值
	 * @param multiReplace 是否支持多值替换
	 * @author taoYan
	 * @since 2018年8月7日
	 */
	private Object replaceValue(String[] replace, Object result,boolean multiReplace) {
		if(result == null){
			return "";
		}
		if(replace == null || replace.length<=0){
			return result;
		}
		String temp = String.valueOf(result);
		String backValue = "";
		if(temp.indexOf(",")>0 && multiReplace){
			//原值中带有逗号，认为他是多值的
			String multiReplaces[] = temp.split(",");
			for (String str : multiReplaces) {
				backValue = backValue.concat(replaceSingleValue(replace, str)+",");
			}
			if(backValue.equals("")){
				backValue = temp;
			}else{
				backValue = backValue.substring(0, backValue.length()-1);
			}
		}else{
			backValue = replaceSingleValue(replace, temp);
		}
		return backValue;
	}
	/**
	 * 单值替换 ,若没找到则原值返回
	 */
	private String replaceSingleValue(String[] replace, String temp){
		String[] tempArr;
		for (int i = 0; i < replace.length; i++) {
			tempArr = replace[i].split("_");
			if (temp.equals(tempArr[0])) {
				return tempArr[1];
			}
		}
		return temp;
	}
	//update-end-author:taoyan date:20180807 for:导入多值替换--
}
