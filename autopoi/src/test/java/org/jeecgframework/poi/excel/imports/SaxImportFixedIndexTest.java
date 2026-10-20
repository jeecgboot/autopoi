package org.jeecgframework.poi.excel.imports;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.excel.ExcelImportUtil;
import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.List;

import static org.junit.Assert.assertEquals;

/**
 * @Excel(fixedIndex = n) 按固定列号映射字段，DOM 和 SAX 导入都应生效，不依赖表头文字。
 *
 * @since 2026-10-20 issues/189 SAX 导入 fixedIndex 不生效
 */
public class SaxImportFixedIndexTest {

	public static class FixedEntity {
		@Excel(name = "任意表头1", fixedIndex = 0)
		private String first;
		@Excel(name = "任意表头2", fixedIndex = 1)
		private String second;

		public String getFirst() {
			return first;
		}

		public void setFirst(String first) {
			this.first = first;
		}

		public String getSecond() {
			return second;
		}

		public void setSecond(String second) {
			this.second = second;
		}
	}

	/** 一个字段按表头名称，一个字段按固定列号 */
	public static class MixedEntity {
		@Excel(name = "名称")
		private String name;
		@Excel(name = "编码", fixedIndex = 2)
		private String code;

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}

		public String getCode() {
			return code;
		}

		public void setCode(String code) {
			this.code = code;
		}
	}

	/** 第一行是表头（headers，可以为 null 表示没有表头），其余是数据 */
	private static InputStream sheetOf(String[] headers, String[]... rows) throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Sheet1");
		int r = 0;
		if (headers != null) {
			Row head = sheet.createRow(r++);
			for (int i = 0; i < headers.length; i++) {
				head.createCell(i).setCellValue(headers[i]);
			}
		}
		for (String[] values : rows) {
			Row row = sheet.createRow(r++);
			for (int i = 0; i < values.length; i++) {
				row.createCell(i).setCellValue(values[i]);
			}
		}
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		wb.write(bos);
		wb.close();
		return new ByteArrayInputStream(bos.toByteArray());
	}

	private static ImportParams saxParams() {
		ImportParams params = new ImportParams();
		params.setSheetNum(1);
		return params;
	}

	@Test
	public void saxShouldMapByFixedIndexIgnoringHeaderText() throws Exception {
		List<FixedEntity> list = ExcelImportUtil.importExcelBySax(
				sheetOf(new String[]{"表头甲", "表头乙"}, new String[]{"v1", "v2"}, new String[]{"v3", "v4"}), FixedEntity.class, saxParams());
		assertEquals(2, list.size());
		assertEquals("v1", list.get(0).getFirst());
		assertEquals("v2", list.get(0).getSecond());
		assertEquals("v3", list.get(1).getFirst());
		assertEquals("v4", list.get(1).getSecond());
	}

	@Test
	public void domShouldMapByFixedIndexIgnoringHeaderText() throws Exception {
		List<FixedEntity> list = ExcelImportUtil.importExcel(
				sheetOf(new String[]{"表头甲", "表头乙"}, new String[]{"v1", "v2"}), FixedEntity.class, new ImportParams());
		assertEquals(1, list.size());
		assertEquals("v1", list.get(0).getFirst());
		assertEquals("v2", list.get(0).getSecond());
	}

	@Test
	public void saxShouldMixNameAndFixedIndex() throws Exception {
		List<MixedEntity> list = ExcelImportUtil.importExcelBySax(
				sheetOf(new String[]{"名称", "无关列", "随便写"}, new String[]{"n1", "x", "c1"}), MixedEntity.class, saxParams());
		assertEquals(1, list.size());
		assertEquals("n1", list.get(0).getName());
		assertEquals("c1", list.get(0).getCode());
	}

	/** 固定列的表头文字恰好是另一个字段的名称，固定列号优先 */
	@Test
	public void fixedIndexShouldWinOverHeaderText() throws Exception {
		List<MixedEntity> list = ExcelImportUtil.importExcelBySax(
				sheetOf(new String[]{"名称", "x", "名称"}, new String[]{"n1", "x1", "c1"}), MixedEntity.class, saxParams());
		assertEquals(1, list.size());
		assertEquals("c1", list.get(0).getCode());
	}

	/** 没有表头行（headRows=0）时，固定列号仍然能映射 */
	@Test
	public void saxShouldMapByFixedIndexWithoutHeaderRow() throws Exception {
		ImportParams params = saxParams();
		params.setHeadRows(0);
		List<FixedEntity> list = ExcelImportUtil.importExcelBySax(
				sheetOf(null, new String[]{"v1", "v2"}, new String[]{"v3", "v4"}), FixedEntity.class, params);
		assertEquals(2, list.size());
		assertEquals("v1", list.get(0).getFirst());
		assertEquals("v4", list.get(1).getSecond());
	}

	/** 同一个实体里按表头名称映射的字段，行为不变 */
	@Test
	public void headerNameMappingShouldKeepWorking() throws Exception {
		List<MixedEntity> list = ExcelImportUtil.importExcelBySax(
				sheetOf(new String[]{"名称", "x", "y"}, new String[]{"n1", "x1", "y1"}), MixedEntity.class, saxParams());
		assertEquals("n1", list.get(0).getName());
	}
}
