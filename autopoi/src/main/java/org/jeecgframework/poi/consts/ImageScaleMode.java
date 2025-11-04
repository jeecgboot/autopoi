package org.jeecgframework.poi.consts;

/**
 * 图片缩放模式枚举
 * for [issues/8892] AutoPoi ImageEntity建议添加scale属性，控制图片导出缩放模式
 * 
 * @author chenrui
 * @date 2025-10-29
 */
public enum ImageScaleMode {
    
    /**
     * 拉伸填充
     */
    STRETCH(0, "拉伸填充"),
    
    /**
     * 等比例缩放适应
     */
    FIT(1, "等比例缩放适应"),
    
    /**
     * 不缩放（原始大小）
     */
    ORIGINAL(2, "不缩放（原始大小）");
    
    private final int code;
    private final String description;
    
    ImageScaleMode(int code, String description) {
        this.code = code;
        this.description = description;
    }
    
    public int getCode() {
        return code;
    }
    
    public String getDescription() {
        return description;
    }
    
    /**
     * 根据code获取枚举
     * 
     * @param code 代码值
     * @return 对应的枚举值，如果找不到则返回STRETCH
     */
    public static ImageScaleMode valueOf(int code) {
        for (ImageScaleMode mode : values()) {
            if (mode.code == code) {
                return mode;
            }
        }
        return STRETCH;
    }
}

