package org.jeecgframework.dict.service;


/**
 * 描述：
 * @author：scott
 * @since：2017-4-12 下午04:58:15
 * @version:1.0
 */
public interface AutoPoiDictServiceI{
	/**
 	 * 方法描述:  查询数据字典
 	 * 作    者： yiming.zhang
 	 * 日    期： 2014年5月11日-下午4:22:42
 	 * @param dicTable
 	 * @param dicCode
 	 * @param dicText
	 * @param dataSource for [issues/7736]@Excel 不支持分布式下表字典跨库查询 #7736
 	 * @return 
 	 * 返回类型： List<DictEntity>
 	 */
 	public String[] queryDict(String dicTable,String dicCode, String dicText, String dataSource) throws Exception;

}
