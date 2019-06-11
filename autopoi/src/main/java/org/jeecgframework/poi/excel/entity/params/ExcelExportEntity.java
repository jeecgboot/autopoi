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

import java.util.List;

/**
 * excel 导出工具类,对cell类型做映射
 * 
 * @author JEECG
 * @version 1.0 2013年8月24日
 */
public class ExcelExportEntity extends ExcelBaseEntity implements Comparable<ExcelExportEntity> {

	/**
	 * 如果是MAP导出,这个是map的key
	 */
	private Object key;

	private double width = 10;

	private double height = 10;

	/**
	 * 图片的类型,1是文件,2是数据库
	 */
	private int exportImageType = 0;

	/**
	 * 排序顺序
	 */
	private int orderNum = 0;

	/**
	 * 是否支持换行
	 */
	private boolean isWrap;

	/**
	 * 是否需要合并
	 */
	private boolean needMerge;
	/**
	 * 单元格纵向合并
	 */
	private boolean mergeVertical;
	/**
	 * 合并依赖
	 */
	private int[] mergeRely;
	/**
	 * 后缀
	 */
	private String suffix;
	/**
	 * 统计
	 */
	private boolean isStatistics;

	private List<ExcelExportEntity> list;

	public ExcelExportEntity() {

	}

	public ExcelExportEntity(String name) {
		super.name = name;
	}

	public ExcelExportEntity(String name, Object key) {
		super.name = name;
		this.key = key;
	}

	public ExcelExportEntity(String name, Object key, int width) {
		super.name = name;
		this.width = width;
		this.key = key;
	}

	public int getExportImageType() {
		return exportImageType;
	}

	public double getHeight() {
		return height;
	}

	public Object getKey() {
		return key;
	}

	public List<ExcelExportEntity> getList() {
		return list;
	}

	public int[] getMergeRely() {
		return mergeRely == null ? new int[0] : mergeRely;
	}

	public int getOrderNum() {
		return orderNum;
	}

	public double getWidth() {
		return width;
	}

	public boolean isMergeVertical() {
		return mergeVertical;
	}

	public boolean isNeedMerge() {
		return needMerge;
	}

	public boolean isWrap() {
		return isWrap;
	}

	public void setExportImageType(int exportImageType) {
		this.exportImageType = exportImageType;
	}

	public void setHeight(double height) {
		this.height = height;
	}

	public void setKey(Object key) {
		this.key = key;
	}

	public void setList(List<ExcelExportEntity> list) {
		this.list = list;
	}

	public void setMergeRely(int[] mergeRely) {
		this.mergeRely = mergeRely;
	}

	public void setMergeVertical(boolean mergeVertical) {
		this.mergeVertical = mergeVertical;
	}

	public void setNeedMerge(boolean needMerge) {
		this.needMerge = needMerge;
	}

	public void setOrderNum(int orderNum) {
		this.orderNum = orderNum;
	}

	public void setWidth(double width) {
		this.width = width;
	}

	public void setWrap(boolean isWrap) {
		this.isWrap = isWrap;
	}

	public String getSuffix() {
		return suffix;
	}

	public void setSuffix(String suffix) {
		this.suffix = suffix;
	}

	public boolean isStatistics() {
		return isStatistics;
	}

	public void setStatistics(boolean isStatistics) {
		this.isStatistics = isStatistics;
	}

	@Override
	public int compareTo(ExcelExportEntity prev) {
		return this.getOrderNum() - prev.getOrderNum();
	}

}
