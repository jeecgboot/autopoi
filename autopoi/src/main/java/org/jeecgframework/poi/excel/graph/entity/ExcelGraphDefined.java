/**
 * 
 */
package org.jeecgframework.poi.excel.graph.entity;

import org.jeecgframework.poi.excel.graph.constant.ExcelGraphType;

import java.util.ArrayList;
import java.util.List;


/**
 * @Description Excel 图形构造服务
 * @author liusq
 * @date  2022年1月4号
 */
public class ExcelGraphDefined implements ExcelGraph
{
	private ExcelGraphElement category;
	public List<ExcelGraphElement> valueList= new ArrayList<>();
	public List<ExcelTitleCell> titleCell= new ArrayList<>();
	private Integer graphType= ExcelGraphType.LINE_CHART;
	public List<String> title= new ArrayList<>();
	
	@Override
    public ExcelGraphElement getCategory()
	{
		return category;
	}
	public void setCategory(ExcelGraphElement category)
	{
		this.category = category;
	}
	@Override
    public List<ExcelGraphElement> getValueList()
	{
		return valueList;
	}
	public void setValueList(List<ExcelGraphElement> valueList)
	{
		this.valueList = valueList;
	}

	@Override
    public Integer getGraphType()
	{
		return graphType;
	}
	public void setGraphType(Integer graphType)
	{
		this.graphType = graphType;
	}
	@Override
    public List<ExcelTitleCell> getTitleCell()
	{
		return titleCell;
	}
	public void setTitleCell(List<ExcelTitleCell> titleCell)
	{
		this.titleCell = titleCell;
	}
	@Override
    public List<String> getTitle()
	{
		return title;
	}
	public void setTitle(List<String> title)
	{
		this.title = title;
	}
	
	
	
}
