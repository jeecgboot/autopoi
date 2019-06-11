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
package org.jeecgframework.poi.exception.word.enmus;

/**
 * 导出异常枚举
 * 
 * @author JEECG
 * @date 2014年8月9日 下午10:34:58
 */
public enum WordExportEnum {

	EXCEL_PARAMS_ERROR("Excel 导出 参数错误"), EXCEL_HEAD_HAVA_NULL("Excel 表头 有的字段为空"), EXCEL_NO_HEAD("Excel 没有表头");

	private String msg;

	WordExportEnum(String msg) {
		this.msg = msg;
	}

	public String getMsg() {
		return msg;
	}

	public void setMsg(String msg) {
		this.msg = msg;
	}

}
