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

import org.jeecgframework.poi.handler.inter.IExcelDataHandler;

/**
 * 基础参数
 * 
 * @author JEECG
 * @date 2014年6月20日 下午1:56:52
 */
public class ExcelBaseParams {

	/**
	 * 数据处理接口,以此为主,replace,format都在这后面
	 */
	private IExcelDataHandler dataHanlder;

	public IExcelDataHandler getDataHanlder() {
		return dataHanlder;
	}

	public void setDataHanlder(IExcelDataHandler dataHanlder) {
		this.dataHanlder = dataHanlder;
	}

}
