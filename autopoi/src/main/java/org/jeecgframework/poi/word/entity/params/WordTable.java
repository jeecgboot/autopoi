package org.jeecgframework.poi.word.entity.params;

import java.util.List;
import java.util.Map;

/**
 * Word表格对象
 *
 * @author Pursuer
 * @version 1.0
 * @date 2023/4/23
 */
public class WordTable {
    /**
     * 表头（key：数据属性名，value：属性描述  例：{name:名称}）
     */
    private Map<String, String> headers;
    /**
     * 数据
     */
    private List<?> data;

    public WordTable() {
    }

    public WordTable(Map<String, String> headers, List<?> data) {
        this.headers = headers;
        this.data = data;
    }

    public Map<String, String> getHeaders() {
        return headers;
    }

    public void setHeaders(Map<String, String> headers) {
        this.headers = headers;
    }

    public List<?> getData() {
        return data;
    }

    public void setData(List<?> data) {
        this.data = data;
    }

    @Override
    public String toString() {
        return "WordTable{" +
                "headers=" + headers +
                ", data=" + data +
                '}';
    }
}