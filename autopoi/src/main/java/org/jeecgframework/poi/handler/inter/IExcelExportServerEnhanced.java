package org.jeecgframework.poi.handler.inter;

import java.util.List;

/**
 * 增强的导出数据接口 - 支持游标分页,解决大数据量导出性能问题
 * for [QQYUN-13964]演示系统数据量大，点击没反应
 * 
 * 推荐实现方式:
 * 1. 使用主键ID或其他有序字段作为游标
 * 2. 每次查询条件: WHERE id > lastId ORDER BY id LIMIT pageSize
 * 3. 避免使用 LIMIT offset, size 这种会随着offset增大而变慢的方式
 * 
 * @Description 解决40万+数据导出查询效率问题
 * @author chenrui
 * @date 2025-11-03
 */
public interface IExcelExportServerEnhanced<T> {
    
    /**
     * 基于游标的分页查询 - 高性能方案
     * 
     * 实现示例:
     * <pre>
     * public List<SysLog> selectListForExcelExport(Object queryParams, SysLog lastRecord, int pageSize) {
     *     Long lastId = lastRecord != null ? ((YourEntity)lastRecord).getId() : 0L;
     *     return mapper.selectList(new QueryWrapper<YourEntity>()
     *         .gt("id", lastId)
     *         .orderByAsc("id")
     *         .last("LIMIT " + pageSize));
     * }
     * </pre>
     * 
     * @param queryParams 查询条件
     * @param lastRecord  上一批次的最后一条记录(首次查询时为null)
     * @param pageSize    每批次查询数量
     * @return 当前批次的数据列表,返回null或空列表表示没有更多数据
     */
    List<T> selectListForExcelExport(Object queryParams, T lastRecord, int pageSize);
    
    /**
     * 获取默认的每批次查询数量
     * 可以根据业务情况调整,建议5000-20000之间
     * 
     * @return 每批次查询数量,默认10000
     */
    default int getPageSize() {
        return 10000;
    }
}

