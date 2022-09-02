/**
 * 
 */
package org.jeecgframework.poi.excel.graph.entity;

/**
 * @Description Excel 图形构造服务
 * @author liusq
 * @date  2022年1月4号
 */
public class ExcelTitleCell
{
	private Integer row;
	private Integer col;
	
	public ExcelTitleCell(){
		
	}
	
	public ExcelTitleCell(Integer row,Integer col){
		this.row=row;
		this.col=col;
	}
	
	public Integer getRow()
	{
		return row;
	}
	public void setRow(Integer row)
	{
		this.row = row;
	}
	public Integer getCol()
	{
		return col;
	}
	public void setCol(Integer col)
	{
		this.col = col;
	}
	
	
}
