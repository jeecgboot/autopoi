package org.jeecgframework.poi.handler.inter;

/**
 * 字典翻译处理
 * @author liusq
 * @date 2022-05-27
 */
public interface IExcelDictHandler {

    /**
     * 从值翻译到名称
     *
     * @param dict  字典Key
     * @param obj   对象
     * @param name  属性名称
     * @param value 属性值
     * @return
     */
    public String toName(String dict, Object obj, String name, Object value);

    /**
     * 从名称翻译到值
     *
     * @param dict  字典Key
     * @param obj   对象
     * @param name  属性名称
     * @param value 属性值
     * @return
     */
    public String toValue(String dict, Object obj, String name, Object value);

}
