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
package org.jeecgframework.poi.excel.imports.verifys;

import org.apache.commons.lang3.StringUtils;
import org.jeecgframework.poi.excel.entity.params.ExcelVerifyEntity;
import org.jeecgframework.poi.excel.entity.result.ExcelVerifyHanlderResult;
import org.jeecgframework.poi.handler.inter.IExcelVerifyHandler;

/**
 * 校验服务
 * 
 * @author JEECG
 * @date 2014年6月29日 下午4:37:56
 */
public class VerifyHandlerServer {

	private final static ExcelVerifyHanlderResult DEFAULT_RESULT = new ExcelVerifyHanlderResult(true);

	private void addVerifyResult(ExcelVerifyHanlderResult hanlderResult, ExcelVerifyHanlderResult result) {
		if (!hanlderResult.isSuccess()) {
			result.setSuccess(false);
			result.setMsg((StringUtils.isEmpty(result.getMsg()) ? "" : result.getMsg() + " , ") + hanlderResult.getMsg());
		}
	}

	/**
	 * 校驗數據
	 * 
	 * @param object
	 * @param value
	 * @param titleString
	 * @param verify
	 * @param excelVerifyHandler
	 */
	public ExcelVerifyHanlderResult verifyData(Object object, Object value, String name, ExcelVerifyEntity verify, IExcelVerifyHandler excelVerifyHandler) {
		if (verify == null) {
			return DEFAULT_RESULT;
		}
		ExcelVerifyHanlderResult result = new ExcelVerifyHanlderResult(true, "");
		if (verify.isNotNull()) {
			addVerifyResult(BaseVerifyHandler.notNull(name, value), result);
		}
		if (verify.isEmail()) {
			addVerifyResult(BaseVerifyHandler.isEmail(name, value), result);
		}
		if (verify.isMobile()) {
			addVerifyResult(BaseVerifyHandler.isMobile(name, value), result);
		}
		if (verify.isTel()) {
			addVerifyResult(BaseVerifyHandler.isTel(name, value), result);
		}
		if (verify.getMaxLength() != -1) {
			addVerifyResult(BaseVerifyHandler.maxLength(name, value, verify.getMaxLength()), result);
		}
		if (verify.getMinLength() != -1) {
			addVerifyResult(BaseVerifyHandler.minLength(name, value, verify.getMinLength()), result);
		}
		if (StringUtils.isNotEmpty(verify.getRegex())) {
			addVerifyResult(BaseVerifyHandler.regex(name, value, verify.getRegex(), verify.getRegexTip()), result);
		}
		if (verify.isInterHandler()) {
			addVerifyResult(excelVerifyHandler.verifyHandler(object, name, value), result);
		}
		return result;

	}
}
