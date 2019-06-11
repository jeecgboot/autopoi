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
package org.jeecgframework.poi.word.entity.params;

/**
 * Excel 对象导出结构
 * 
 * @author JEECG
 * @date 2014年7月26日 下午11:14:48
 */
public class ListParamEntity {
	// 唯一值,在遍历中重复使用
	public static final String SINGLE = "single";
	// 属于数组类型
	public static final String LIST = "list";
	/**
	 * 属性名称
	 */
	private String name;
	/**
	 * 目标
	 */
	private String target;
	/**
	 * 当是唯一值的时候直接求出值
	 */
	private Object value;
	/**
	 * 数据类型,SINGLE || LIST
	 */
	private String type;

	public ListParamEntity() {

	}

	public ListParamEntity(String name, Object value) {
		this.name = name;
		this.value = value;
		this.type = LIST;
	}

	public ListParamEntity(String name, String target) {
		this.name = name;
		this.target = target;
		this.type = LIST;
	}

	public String getName() {
		return name;
	}

	public String getTarget() {
		return target;
	}

	public String getType() {
		return type;
	}

	public Object getValue() {
		return value;
	}

	public void setName(String name) {
		this.name = name;
	}

	public void setTarget(String target) {
		this.target = target;
	}

	public void setType(String type) {
		this.type = type;
	}

	public void setValue(Object value) {
		this.value = value;
	}
}
