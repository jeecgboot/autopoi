package org.jeecgframework.poi.handler.inter;

import java.util.Collection;
/**
 * 大数据写出服务接口
 *
 * @Description [LOWCOD-2521]【autopoi】大数据导出方法【全局】
 * @author liusq
 * @date 2022年1月4号
 */
public interface IWriter<T> {
    /**
     * 获取输出对象
     *
     * @return
     */
    default public T get() {
        return null;
    }

    /**
     * 写入数据
     *
     * @param data
     * @return
     */
    public IWriter<T> write(Collection data);

    /**
     * 关闭流,完成业务
     *
     * @return
     */
    public T close();
}
