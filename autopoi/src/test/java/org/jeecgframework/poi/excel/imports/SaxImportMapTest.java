package org.jeecgframework.poi.excel.imports;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.excel.ExcelImportUtil;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.jeecgframework.poi.handler.inter.IExcelReadRowHanlder;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

/**
 * SAX 导入以 Map.class 作为目标类型：每一行返回一个 Map，key 是表头，value 是单元格的值。
 * DOM 导入早就支持，SAX 以前每行都返回空 Map（读取内容为空）。
 *
 * @since 2026-10-20 issues/140 importExcelBySax 读取内容为空
 */
public class SaxImportMapTest {

	/** rows 里的元素：null 表示不创建单元格；Number 数值单元格；Date 日期单元格；其它按文本 */
	private static InputStream sheetOf(Object[]... rows) throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		CellStyle dateStyle = wb.createCellStyle();
		dateStyle.setDataFormat(wb.createDataFormat().getFormat("yyyy-mm-dd"));
		Sheet sheet = wb.createSheet("Sheet1");
		for (int r = 0; r < rows.length; r++) {
			if (rows[r] == null) {
				continue;
			}
			Row row = sheet.createRow(r);
			for (int c = 0; c < rows[r].length; c++) {
				Object v = rows[r][c];
				if (v == null) {
					continue;
				}
				if (v instanceof Number) {
					row.createCell(c).setCellValue(((Number) v).doubleValue());
				} else if (v instanceof Date) {
					row.createCell(c).setCellValue((Date) v);
					row.getCell(c).setCellStyle(dateStyle);
				} else {
					row.createCell(c).setCellValue(String.valueOf(v));
				}
			}
		}
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		wb.write(bos);
		wb.close();
		return new ByteArrayInputStream(bos.toByteArray());
	}

	private static ImportParams params() {
		ImportParams params = new ImportParams();
		params.setSheetNum(1);
		return params;
	}

	@SuppressWarnings("rawtypes")
	private static List<Map> importMaps(InputStream is) {
		return ExcelImportUtil.importExcelBySax(is, Map.class, params());
	}

	@Test
	public void eachRowShouldBeAMapKeyedByHeader() throws Exception {
		List<Map> list = importMaps(sheetOf(
				new Object[]{"A", "B"}, new Object[]{"a1", "b1"}, new Object[]{"a2", "b2"}));
		assertEquals(2, list.size());
		assertEquals("a1", list.get(0).get("A"));
		assertEquals("b1", list.get(0).get("B"));
		assertEquals("a2", list.get(1).get("A"));
		assertEquals("b2", list.get(1).get("B"));
	}

	@Test
	public void columnOrderShouldBeKept() throws Exception {
		List<Map> list = importMaps(sheetOf(
				new Object[]{"名称", "数量", "备注"}, new Object[]{"苹果", 3, "好"}));
		assertTrue("应是保持列顺序的 Map", list.get(0) instanceof LinkedHashMap);
		assertEquals(Arrays.asList("名称", "数量", "备注"), new ArrayList<Object>(list.get(0).keySet()));
	}

	@Test
	public void numbersShouldKeepTheirValue() throws Exception {
		List<Map> list = importMaps(sheetOf(
				new Object[]{"名称", "数量", "单价"}, new Object[]{"苹果", 3, 2.5}));
		assertEquals(0, new BigDecimal("3").compareTo((BigDecimal) list.get(0).get("数量")));
		assertEquals(0, new BigDecimal("2.5").compareTo((BigDecimal) list.get(0).get("单价")));
	}

	@Test
	public void dateCellShouldBeDate() throws Exception {
		Date d = new Date(1700000000000L);
		List<Map> list = importMaps(sheetOf(
				new Object[]{"名称", "日期"}, new Object[]{"甲", d}));
		Object v = list.get(0).get("日期");
		assertNotNull(v);
		assertTrue("日期单元格应是 Date，实际: " + v.getClass(), v instanceof Date);
	}

	/** 与 DOM 一致：每个表头都有对应的 key，空单元格的值是 null */
	@Test
	public void emptyCellsShouldBeNullButKeysShouldExist() throws Exception {
		List<Map> list = importMaps(sheetOf(
				new Object[]{"A", "B", "C"}, new Object[]{"a1", null, "c1"}, new Object[]{null, "b2", null}));
		assertEquals(2, list.size());
		assertEquals(3, list.get(0).size());
		assertTrue(list.get(0).containsKey("B"));
		assertNull(list.get(0).get("B"));
		assertEquals("c1", list.get(0).get("C"));
		assertNull(list.get(1).get("A"));
		assertEquals("b2", list.get(1).get("B"));
		assertNull(list.get(1).get("C"));
	}

	/** 第一列(主键列)为空的行也是一条新数据，不能被当成上一条的子行而丢弃 */
	@Test
	public void rowWithEmptyFirstColumnShouldNotBeDropped() throws Exception {
		List<Map> list = importMaps(sheetOf(
				new Object[]{"A", "B"}, new Object[]{"a1", "b1"}, new Object[]{null, "b2"}, new Object[]{"a3", "b3"}));
		assertEquals(3, list.size());
		assertEquals("b2", list.get(1).get("B"));
		assertEquals("a3", list.get(2).get("A"));
	}

	@Test
	public void emptyHeaderCellShouldBeSkipped() throws Exception {
		List<Map> list = importMaps(sheetOf(
				new Object[]{"A", null, "C"}, new Object[]{"a", "ignored", "c"}));
		assertEquals(2, list.get(0).size());
		assertEquals("a", list.get(0).get("A"));
		assertEquals("c", list.get(0).get("C"));
	}

	@Test
	public void headerOnlyShouldReturnEmptyList() throws Exception {
		List<Map> list = importMaps(sheetOf(new Object[]{"A", "B"}));
		assertEquals(0, list.size());
	}

	/** 和 issue #140 一致的参数：titleRows=0、headRows=1、startSheetIndex=0、sheetNum=1、startRows=0 */
	@Test
	public void paramsFromTheIssueShouldWork() throws Exception {
		ImportParams p = new ImportParams();
		p.setTitleRows(0);
		p.setHeadRows(1);
		p.setStartSheetIndex(0);
		p.setSheetNum(1);
		p.setNeedSave(true);
		List<Map> list = ExcelImportUtil.importExcelBySax(
				sheetOf(new Object[]{"A", "B"}, new Object[]{"a1", "b1"}), Map.class, p);
		assertEquals(1, list.size());
		assertEquals("a1", list.get(0).get("A"));
	}

	@Test
	public void titleRowsAboveHeaderShouldBeSkipped() throws Exception {
		ImportParams p = params();
		p.setTitleRows(1);
		p.setHeadRows(1);
		List<Map> list = ExcelImportUtil.importExcelBySax(
				sheetOf(new Object[]{"这是标题"}, new Object[]{"A", "B"}, new Object[]{"a1", "b1"}), Map.class, p);
		assertEquals(1, list.size());
		assertEquals("a1", list.get(0).get("A"));
	}

	/** 回调模式（传 IExcelReadRowHanlder）下 Map 同样可用 */
	@Test
	@SuppressWarnings({"rawtypes", "unchecked"})
	public void handlerModeShouldDeliverMaps() throws Exception {
		final List<Map> got = new ArrayList<Map>();
		ExcelImportUtil.importExcelBySax(sheetOf(
				new Object[]{"A"}, new Object[]{"a1"}, new Object[]{"a2"}, new Object[]{"a3"}),
				Map.class, params(), new IExcelReadRowHanlder<Map>() {
					@Override
					public void hanlder(Map map) {
						got.add(map);
					}
				});
		// 回调模式下一条数据要等读到下一条才输出，这里至少前两条一定会被回调（最后一条见 SAX 回调漏最后一行的单独问题）
		assertTrue(got.size() >= 2);
		assertEquals("a1", got.get(0).get("A"));
		assertEquals("a2", got.get(1).get("A"));
	}
}
