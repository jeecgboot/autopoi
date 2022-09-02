package org.jeecgframework.poi.handler.inter;

import java.util.List;

/**
 * 导出数据接口
 * @Description [LOWCOD-2521]【autopoi】大数据导出方法【全局】
 * @author liusq
 * @date  2022年1月4号
 */
public interface IExcelExportServer {
    /**
     * 查询数据接口
     *
     * @param queryParams 查询条件
     * @param page        当前页数从1开始
     * @data 2022年1月4号
     * @return
     */
    public List<Object> selectListForExcelExport(Object queryParams, int page);
}
