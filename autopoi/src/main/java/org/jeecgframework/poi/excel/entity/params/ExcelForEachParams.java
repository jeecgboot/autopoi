package org.jeecgframework.poi.excel.entity.params;

import java.io.Serializable;
import java.util.Stack;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 模板for each是的参数
 * @author JueYue
 * @date 2015年4月29日 下午9:22:48
 */
public class ExcelForEachParams implements Serializable {

    /**
     *
     */
    private static final long serialVersionUID = 1L;
    /**
     * key
     */
    private String            name;
    /**
     * key
     */
    private Stack<String> tempName;
    /**
     * 模板的cellStyle
     */
    private CellStyle         cellStyle;
    /**
     * 行高
     */
    private short             height;
    /**
     * 常量值
     */
    private String            constValue;
    /**
     * 列合并
     */
    private int               colspan          = 1;
    /**
     * 行合并
     */
    private int               rowspan          = 1;
    /**
     * 行合并
     */
    private              boolean       collectCell;

    public ExcelForEachParams() {

    }

    public ExcelForEachParams(String name, CellStyle cellStyle, short height) {
        this.name = name;
        this.cellStyle = cellStyle;
        this.height = height;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public short getHeight() {
        return height;
    }

    public void setHeight(short height) {
        this.height = height;
    }

    public String getConstValue() {
        return constValue;
    }

    public void setConstValue(String constValue) {
        this.constValue = constValue;
    }

    public int getColspan() {
        return colspan;
    }

    public void setColspan(int colspan) {
        this.colspan = colspan;
    }

    public int getRowspan() {
        return rowspan;
    }

    public void setRowspan(int rowspan) {
        this.rowspan = rowspan;
    }

    public boolean isCollectCell() {
        return collectCell;
    }

    public void setCollectCell(boolean collectCell) {
        this.collectCell = collectCell;
    }

    public Stack<String> getTempName() {
        return tempName;
    }

    public void setTempName(Stack<String> tempName) {
        this.tempName = tempName;
    }
}
