package org.jeecgframework.poi.excel.imports.base;

public interface ImportFileServiceI {

    /**
     * 上传文件 返回文件地址字符串
     * @param data
     * @return
     */
    String doUpload(byte[] data);

    /**
     * 上传文件 返回文件地址字符串
     * @param data
     * @param saveUrl 保存路径
     * @return
     */
     String doUpload(byte[] data,String saveUrl);

}
