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
package org.jeecgframework.poi.excel.entity.result;

/**
 * Excel导入处理返回结果
 * 
 * @author JEECG
 * @date 2014年6月23日 下午11:03:29
 */
public class ExcelVerifyHanlderResult {
	/**
	 * 是否正确
	 */
	private boolean success;
	/**
	 * 错误信息
	 */
	private String msg;

	public ExcelVerifyHanlderResult() {

	}

	public ExcelVerifyHanlderResult(boolean success) {
		this.success = success;
	}

	public ExcelVerifyHanlderResult(boolean success, String msg) {
		this.success = success;
		this.msg = msg;
	}

	public String getMsg() {
		return msg;
	}

	public boolean isSuccess() {
		return success;
	}

	public void setMsg(String msg) {
		this.msg = msg;
	}

	public void setSuccess(boolean success) {
		this.success = success;
	}

}
