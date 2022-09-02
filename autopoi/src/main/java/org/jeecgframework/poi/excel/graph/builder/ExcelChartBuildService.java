/**
 * 
 */
package org.jeecgframework.poi.excel.graph.builder;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.charts.AxisCrosses;
import org.apache.poi.ss.usermodel.charts.AxisPosition;
import org.apache.poi.ss.usermodel.charts.ChartAxis;
import org.apache.poi.ss.usermodel.charts.ChartDataSource;
import org.apache.poi.ss.usermodel.charts.ChartLegend;
import org.apache.poi.ss.usermodel.charts.DataSources;
import org.apache.poi.ss.usermodel.charts.LegendPosition;
import org.apache.poi.ss.usermodel.charts.LineChartData;
import org.apache.poi.ss.usermodel.charts.ScatterChartData;
import org.apache.poi.ss.usermodel.charts.ValueAxis;
import org.apache.poi.ss.util.CellRangeAddress;

import org.jeecgframework.poi.excel.graph.constant.ExcelGraphElementType;
import org.jeecgframework.poi.excel.graph.constant.ExcelGraphType;
import org.jeecgframework.poi.excel.graph.entity.ExcelGraph;
import org.jeecgframework.poi.excel.graph.entity.ExcelGraphElement;
import org.jeecgframework.poi.excel.graph.entity.ExcelTitleCell;
import org.jeecgframework.poi.util.PoiCellUtil;
import org.jeecgframework.poi.util.PoiExcelGraphDataUtil;

/**
 * @Description
 * @author liusq
 * @data 2022年1月4号
 */
public class ExcelChartBuildService
{
	/**
	 * 
	 * @param workbook
	 * @param graphList
	 * @param build 通过实时数据行来重新计算图形定义
	 * @param append
	 */
	public static void createExcelChart(Workbook workbook, List<ExcelGraph> graphList, Boolean build, Boolean append)
	{
		if(workbook!=null&&graphList!=null){
			//设定默认第一个sheet为数据项
			Sheet dataSouce=workbook.getSheetAt(0);
			if(dataSouce!=null){
				buildTitle(dataSouce,graphList);
				
				if(build){
					PoiExcelGraphDataUtil.buildGraphData(dataSouce, graphList);
				}
				if(append){
					buildExcelChart(dataSouce, dataSouce, graphList);
				}else{
					Sheet sheet=workbook.createSheet("图形界面");
					buildExcelChart(dataSouce, sheet, graphList);
				}
			}
			
		}
	}
	
	/**
	 * 构建基础图形
	 * @param drawing 
	 * @param anchor
	 * @param dataSourceSheet
	 * @param graph
	 */
	private static void buildExcelChart(Drawing drawing,ClientAnchor anchor,Sheet dataSourceSheet,ExcelGraph graph){
		Chart chart = null;
		// TODO  图表没有成功
		//drawing.createChart(anchor);
		ChartLegend legend = chart.getOrCreateLegend();
		legend.setPosition(LegendPosition.TOP_RIGHT);
		
		ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
        ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        ExcelGraphElement categoryElement=graph.getCategory();
        
        ChartDataSource categoryChart;
        if(categoryElement!=null&& categoryElement.getElementType().equals(ExcelGraphElementType.STRING_TYPE)){
        	categoryChart=DataSources.fromStringCellRange(dataSourceSheet, new CellRangeAddress(categoryElement.getStartRowNum(),categoryElement.getEndRowNum(),categoryElement.getStartColNum(),categoryElement.getEndColNum()));
        }else{
        	categoryChart=DataSources.fromNumericCellRange(dataSourceSheet, new CellRangeAddress(categoryElement.getStartRowNum(),categoryElement.getEndRowNum(),categoryElement.getStartColNum(),categoryElement.getEndColNum()));
        }
        
        List<ExcelGraphElement> valueList=graph.getValueList();
        List<ChartDataSource<Number>> chartValueList= new ArrayList<>();
        if(valueList!=null&&valueList.size()>0){
        	for(ExcelGraphElement ele:valueList){
        		ChartDataSource<Number> source=DataSources.fromNumericCellRange(dataSourceSheet, new CellRangeAddress(ele.getStartRowNum(),ele.getEndRowNum(),ele.getStartColNum(),ele.getEndColNum()));
        		chartValueList.add(source);
        	}
        }
        
		if(graph.getGraphType().equals(ExcelGraphType.LINE_CHART)){
			LineChartData data = chart.getChartDataFactory().createLineChartData();
			buildLineChartData(data, categoryChart, chartValueList, graph.getTitle());
			chart.plot(data, bottomAxis, leftAxis);
		}
		else
		{
			ScatterChartData data=chart.getChartDataFactory().createScatterChartData();
			buildScatterChartData(data, categoryChart, chartValueList,graph.getTitle());
			chart.plot(data, bottomAxis, leftAxis);
		} 
	}
	
	
	
	
	/**
	 * 构建多个图形对象
	 * @param dataSourceSheet
	 * @param tragetSheet
	 * @param graphList
	 */
	private static void buildExcelChart(Sheet dataSourceSheet,Sheet tragetSheet,List<ExcelGraph> graphList){
		int len=graphList.size();
		if(len==1)
		{
			buildExcelChart(dataSourceSheet, tragetSheet, graphList.get(0));
		}
		else
		{
			int drawStart=0;
			int drawEnd=20;
			Drawing drawing = PoiExcelGraphDataUtil.getDrawingPatriarch(tragetSheet);
			for(int i=0;i<len;i++){
				ExcelGraph graph=graphList.get(i);
				ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, drawStart, 15, drawEnd);
				buildExcelChart(drawing, anchor, dataSourceSheet, graph);
				drawStart=drawStart+drawEnd;
				drawEnd=drawEnd+drawEnd;
			}
		}
	}
	
	
	
	
	/**
	 * 构建图形对象
	 * @param dataSourceSheet
	 * @param tragetSheet
	 * @param graph
	 */
	private static void buildExcelChart(Sheet dataSourceSheet,Sheet tragetSheet,ExcelGraph graph){
		Drawing drawing = PoiExcelGraphDataUtil.getDrawingPatriarch(tragetSheet);
		ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 0, 15, 20);
		buildExcelChart(drawing, anchor, dataSourceSheet, graph);
	}
	
	
	
	
	/**
	 * 构建Title
	 * @param sheet
	 * @param graph
	 */
	private static void buildTitle(Sheet sheet,ExcelGraph graph){
		int cellTitleLen=graph.getTitleCell().size();
		int titleLen=graph.getTitle().size();
		if(titleLen>0){
			
		}else{
			for(int i=0;i<cellTitleLen;i++){
				ExcelTitleCell titleCell=graph.getTitleCell().get(i);
				if(titleCell!=null){
					graph.getTitle().add(PoiCellUtil.getCellValue(sheet,titleCell.getRow(),titleCell.getCol()));
				}
			}
		}
	}
	
	/**
	 * 构建Title
	 * @param sheet
	 * @param graphList
	 */
	private static void buildTitle(Sheet sheet,List<ExcelGraph> graphList){
		if(graphList!=null&&graphList.size()>0){
			for(ExcelGraph graph:graphList){
				if(graph!=null)
				{
					buildTitle(sheet, graph);
				}
			}
		}
	}
	
	/**
	 * 
	 * @param data
	 * @param categoryChart
	 * @param chartValueList
	 * @param title
	 */
	private static void buildLineChartData(LineChartData data,ChartDataSource categoryChart,List<ChartDataSource<Number>> chartValueList,List<String> title){
		if(chartValueList.size()==title.size())
		{
			int len=title.size();
			for(int i=0;i<len;i++){
			    //TODO 更新版本
				//data.addSerie(categoryChart, chartValueList.get(i)).setTitle(title.get(i));
			}
		}	
		else
		{
			int i=0;
			for(ChartDataSource<Number> source:chartValueList){
				String temp_title=title.get(i);
				if(StringUtils.isNotBlank(temp_title)){
					//data.addSerie(categoryChart, source).setTitle(_title);
				}else{
					//data.addSerie(categoryChart, source);
				}
			}
		}
	}
	
	/**
	 * 
	 * @param data
	 * @param categoryChart
	 * @param chartValueList
	 * @param title
	 */
	private static void buildScatterChartData(ScatterChartData data,ChartDataSource categoryChart,List<ChartDataSource<Number>> chartValueList,List<String> title){
		if(chartValueList.size()==title.size())
		{
			int len=title.size();
			for(int i=0;i<len;i++){
				data.addSerie(categoryChart, chartValueList.get(i)).setTitle(title.get(i));
			}
		}	
		else
		{
			int i=0;
			for(ChartDataSource<Number> source:chartValueList){
				String temp_title=title.get(i);
				if(StringUtils.isNotBlank(temp_title)){
					data.addSerie(categoryChart, source).setTitle(temp_title);
				}else{
					data.addSerie(categoryChart, source);
				}
			}
		}
	}
	
	
}
