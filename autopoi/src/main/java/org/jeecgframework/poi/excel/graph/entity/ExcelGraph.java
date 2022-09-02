/**
 * 
 */
package org.jeecgframework.poi.excel.graph.entity;

import java.util.List;

/**
 * @Description Excel 图形构造服务
 * @author liusq
 * @date  2022年1月4号
 */
public interface ExcelGraph
{
	public ExcelGraphElement getCategory();
	public List<ExcelGraphElement> getValueList();
	public Integer getGraphType();
	public List<ExcelTitleCell> getTitleCell();
	public List<String> getTitle();
}
