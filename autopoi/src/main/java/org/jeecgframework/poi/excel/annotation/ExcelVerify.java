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
package org.jeecgframework.poi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel 导入校验
 * 
 * @author JEECG
 * @date 2014年6月23日 下午10:46:26
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelVerify {
	/**
	 * 接口校验
	 * 
	 * @return
	 */
	public boolean interHandler() default false;

	/**
	 * 是电子邮件
	 * 
	 * @return
	 */
	public boolean isEmail() default false;

	/**
	 * 是13位移动电话
	 * 
	 * @return
	 */
	public boolean isMobile() default false;

	/**
	 * 是座机号码
	 * 
	 * @return
	 */
	public boolean isTel() default false;

	/**
	 * 最大长度
	 * 
	 * @return
	 */
	public int maxLength() default -1;

	/**
	 * 最小长度
	 * 
	 * @return
	 */
	public int minLength() default -1;

	/**
	 * 不允许空
	 * 
	 * @return
	 */
	public boolean notNull() default false;

	/**
	 * 正在表达式
	 * 
	 * @return
	 */
	public String regex() default "";

	/**
	 * 正在表达式,错误提示信息
	 * 
	 * @return
	 */
	public String regexTip() default "数据不符合规范";

}
