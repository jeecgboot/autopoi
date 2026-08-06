package org.jeecgframework.dict.service;


import java.util.HashMap;

/**
 * 描述：
 * @author：scott
 * @since：2024-09-09
 * @version:1.0
 */
public interface AutoPoiDictMapServiceI {
	/**
 	 * 方法描述:  查询数据字典优化
 	 * 作    者： TestNet
 	 * @param dicTable
 	 * @param dicCode
 	 * @param dicText
 	 * @return 
 	 * 返回类型： HashMap<key,value>
 	 */
 	public HashMap<String,String> queryDict(String dicTable, String dicCode, String dicText, boolean isKeyValue);

}
