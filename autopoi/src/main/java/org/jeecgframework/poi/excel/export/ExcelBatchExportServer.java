package org.jeecgframework.poi.excel.export;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.jeecgframework.poi.excel.annotation.ExcelTarget;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.jeecgframework.poi.excel.entity.vo.PoiBaseConstants;
import org.jeecgframework.poi.excel.export.styler.IExcelExportStyler;
import org.jeecgframework.poi.exception.excel.ExcelExportException;
import org.jeecgframework.poi.exception.excel.enums.ExcelExportEnum;
import org.jeecgframework.poi.handler.inter.IExcelExportServer;
import org.jeecgframework.poi.handler.inter.IWriter;
import org.jeecgframework.poi.util.PoiExcelGraphDataUtil;
import org.jeecgframework.poi.util.PoiPublicUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.*;

import static  org.jeecgframework.poi.excel.ExcelExportUtil.USE_SXSSF_LIMIT;

/**
 * 提供批次插入服务
 * @author liusq
 * @date 2022年1月4日
 */
public class ExcelBatchExportServer extends ExcelExportServer implements IWriter<Workbook> {

	private final static Logger LOGGER = LoggerFactory.getLogger(ExcelBatchExportServer.class);

	private Workbook                workbook;
	private Sheet                   sheet;
	private List<ExcelExportEntity> excelParams;
	private ExportParams entity;
	private int                     titleHeight;
	private Drawing                 patriarch;
	private short                   rowHeight;
	private int                     index;

	public void init(ExportParams entity, Class<?> pojoClass) {
		List<ExcelExportEntity> excelParams = createExcelExportEntityList(entity, pojoClass);
		init(entity, excelParams);
	}

	/**
	 * 初始化数据
	 * @param entity  导出参数
	 * @param excelParams
	 */
	public void init(ExportParams entity, List<ExcelExportEntity> excelParams) {
		LOGGER.debug("ExcelBatchExportServer only support SXSSFWorkbook");
		entity.setType(ExcelType.XSSF);
		workbook = new SXSSFWorkbook();
		this.entity = entity;
		this.excelParams = excelParams;
		super.type = entity.getType();
		createSheet(workbook, entity, excelParams);
		if (entity.getMaxNum() == 0) {
			entity.setMaxNum(USE_SXSSF_LIMIT);
		}
		insertDataToSheet(workbook, entity, excelParams, null, sheet);
	}

	public List<ExcelExportEntity> createExcelExportEntityList(ExportParams entity, Class<?> pojoClass) {
		try {
			List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
			if (entity.isAddIndex()) {
				excelParams.add(indexExcelEntity(entity));
			}
			// 得到所有字段
			Field[]     fileds   = PoiPublicUtil.getClassFields(pojoClass);
			ExcelTarget etarget  = pojoClass.getAnnotation(ExcelTarget.class);
			String      targetId = etarget == null ? null : etarget.value();
			getAllExcelField(entity.getExclusions(), targetId, fileds, excelParams, pojoClass,
					null);
			sortAllParams(excelParams);

			return excelParams;
		} catch (Exception e) {
			throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e);
		}
	}

	public void createSheet(Workbook workbook, ExportParams entity, List<ExcelExportEntity> excelParams) {
		if (LOGGER.isDebugEnabled()) {
			LOGGER.debug("Excel export start ,List<ExcelExportEntity> is {}", excelParams);
			LOGGER.debug("Excel version is {}",
					entity.getType().equals(ExcelType.HSSF) ? "03" : "07");
		}
		if (workbook == null || entity == null || excelParams == null) {
			throw new ExcelExportException(ExcelExportEnum.PARAMETER_ERROR);
		}
		try {
			try {
				sheet = workbook.createSheet(entity.getSheetName());
			} catch (Exception e) {
				// 重复遍历,出现了重名现象,创建非指定的名称Sheet
				sheet = workbook.createSheet();
			}
		} catch (Exception e) {
			throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e);
		}
	}

	@Override
	protected void insertDataToSheet(Workbook workbook, ExportParams entity,
									 List<ExcelExportEntity> entityList, Collection<? extends Map<?, ?>> dataSet,
									 Sheet sheet) {
		try {
			dataHanlder = entity.getDataHanlder();
			if (dataHanlder != null && dataHanlder.getNeedHandlerFields() != null) {
				needHanlderList = Arrays.asList(dataHanlder.getNeedHandlerFields());
			}
			// 创建表格样式
			setExcelExportStyler((IExcelExportStyler) entity.getStyle()
					.getConstructor(Workbook.class).newInstance(workbook));
			patriarch = PoiExcelGraphDataUtil.getDrawingPatriarch(sheet);
			List<ExcelExportEntity> excelParams = new ArrayList<ExcelExportEntity>();
			if (entity.isAddIndex()) {
				excelParams.add(indexExcelEntity(entity));
			}
			excelParams.addAll(entityList);
			sortAllParams(excelParams);
			this.index = entity.isCreateHeadRows()
					? createHeaderAndTitle(entity, sheet, workbook, excelParams) : 0;
			titleHeight = index;
			setCellWith(excelParams, sheet);
			setColumnHidden(excelParams, sheet);
			rowHeight = getRowHeight(excelParams);
			setCurrentIndex(1);
		} catch (Exception e) {
			LOGGER.error(e.getMessage(), e);
			throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e.getCause());
		}
	}

	public Workbook exportBigExcel(IExcelExportServer server, Object queryParams) {
		int page = 1;
		List<Object> list = server
				.selectListForExcelExport(queryParams, page++);
		while (list != null && list.size() > 0) {
			write(list);
			list = server.selectListForExcelExport(queryParams, page++);
		}
		return close();
	}

	@Override
	public Workbook get() {
		return this.workbook;
	}

	@Override
	public IWriter<Workbook> write(Collection data) {
		if (sheet.getLastRowNum() + data.size() > entity.getMaxNum()) {
			sheet = workbook.createSheet();
			index = 0;
		}
		Iterator<?> its = data.iterator();
		while (its.hasNext()) {
			Object t = its.next();
			try {
				index += createCells(patriarch, index, t, excelParams, sheet, workbook, rowHeight, 0)[0];
			} catch (Exception e) {
				LOGGER.error(e.getMessage(), e);
				throw new ExcelExportException(ExcelExportEnum.EXPORT_ERROR, e);
			}
		}
		return this;
	}

	@Override
	public Workbook close() {
		if (entity.getFreezeCol() != 0) {
			sheet.createFreezePane(entity.getFreezeCol(), titleHeight, entity.getFreezeCol(), titleHeight);
		}
		mergeCells(sheet, excelParams, titleHeight);
		// 创建合计信息
		addStatisticsRow(getExcelExportStyler().getStyles(true, null), sheet);
		return workbook;
	}
	/**
	 * 添加Index列
	 */
	@Override
	public ExcelExportEntity indexExcelEntity(ExportParams entity) {
		ExcelExportEntity exportEntity = new ExcelExportEntity();
		//保证是第一排
		exportEntity.setOrderNum(Integer.MIN_VALUE);
		exportEntity.setNeedMerge(true);
		exportEntity.setName(entity.getIndexName());
		exportEntity.setWidth(10);
		exportEntity.setFormat(PoiBaseConstants.IS_ADD_INDEX);
		return exportEntity;
	}
}