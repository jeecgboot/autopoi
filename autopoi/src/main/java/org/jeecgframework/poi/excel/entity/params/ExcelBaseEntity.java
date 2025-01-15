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
package org.jeecgframework.poi.excel.entity.params;

import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.List;

/**
 * Excel 导入导出基础对象类
 * 
 * @author JEECG
 * @date 2014年6月20日 下午2:26:09
 */
public class ExcelBaseEntity {
	/**
	 * 对应name
	 */
	protected String name;
	/**
	 * 对应type
	 */
	private int type = 1;
	/**
	 * 数据库格式
	 */
	private String databaseFormat;
	/**
	 * 导出日期格式
	 */
	private String format;

	/**
	 * 数字格式化,参数是Pattern,使用的对象是DecimalFormat
	 */
	private String numFormat;
	/**
	 * 替换值表达式 ："男_1","女_0"
	 */
	private String[] replace;
	/**
	 * 替换是否是替换多个值
	 */
	private boolean multiReplace;

	/**
	 * 表头组名称
	 */
	private String groupName;
	
	/**
	 * set/get方法
	 */
	private Method method;
	/**
	 * 固定的列
	 */
	private Integer      fixedIndex;
	/**
	 * 字典名称
	 */
	private String       dict;
	/**
	 * 这个是不是超链接,如果是需要实现接口返回对象
	 */
	private boolean     hyperlink;

	private List<Method> methods;

	private HashMap<String,String> replaceMap;


	public HashMap<String, String> getReplaceMap() {
		return replaceMap;
	}

	public void setReplaceMap(HashMap<String, String> replaceMap) {
		this.replaceMap = replaceMap;
	}

	public String getDatabaseFormat() {
		return databaseFormat;
	}

	public String getFormat() {
		return format;
	}

	public Method getMethod() {
		return method;
	}

	public List<Method> getMethods() {
		return methods;
	}

	public String getName() {
		return name;
	}

	public String[] getReplace() {
		return replace;
	}

	public int getType() {
		return type;
	}

	public void setDatabaseFormat(String databaseFormat) {
		this.databaseFormat = databaseFormat;
	}

	public void setFormat(String format) {
		this.format = format;
	}

	public void setMethod(Method method) {
		this.method = method;
	}

	public void setMethods(List<Method> methods) {
		this.methods = methods;
	}

	public void setName(String name) {
		this.name = name;
	}

	public void setReplace(String[] replace) {
		this.replace = replace;
	}

	public void setType(int type) {
		this.type = type;
	}

	public boolean isMultiReplace() {
		return multiReplace;
	}
	public void setMultiReplace(boolean multiReplace) {
		this.multiReplace = multiReplace;
	}

	public String getNumFormat() {
		return numFormat;
	}

	public void setNumFormat(String numFormat) {
		this.numFormat = numFormat;
	}

	public String getGroupName() {
		return groupName;
	}

	public void setGroupName(String groupName) {
		this.groupName = groupName;
	}

	public Integer getFixedIndex() {
		return fixedIndex;
	}

	public void setFixedIndex(Integer fixedIndex) {
		this.fixedIndex = fixedIndex;
	}

	public String getDict() {
		return dict;
	}

	public void setDict(String dict) {
		this.dict = dict;
	}

	public boolean isHyperlink() {
		return hyperlink;
	}

	public void setHyperlink(boolean hyperlink) {
		this.hyperlink = hyperlink;
	}
}
