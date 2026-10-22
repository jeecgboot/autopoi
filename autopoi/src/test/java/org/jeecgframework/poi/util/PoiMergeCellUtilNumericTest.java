package org.jeecgframework.poi.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;
import org.junit.Test;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.junit.Assert.assertEquals;

/**
 * 纵向合并相同内容的单元格时，被合并的列（或依赖的列）是数值类型也应能正常合并，
 * 不能因为 getStringCellValue 读取数值单元格而抛异常。
 *
 * @since 2026-10-20 issues/29 合并单元格，@Excel 的 type 设置非 1 类型时读取单元格内容报错
 */
public class PoiMergeCellUtilNumericTest {

	public static class MergeNumber {
		@Excel(name = "数量", orderNum = "1", type = 4, mergeVertical = true)
		private Integer qty;
		@Excel(name = "名称", orderNum = "2")
		private String name;

		public MergeNumber(Integer qty, String name) {
			this.qty = qty;
			this.name = name;
		}

		public Integer getQty() {
			return qty;
		}

		public void setQty(Integer qty) {
			this.qty = qty;
		}

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}
	}

	public static class MergeDecimal {
		@Excel(name = "金额", orderNum = "1", type = 4, mergeVertical = true)
		private BigDecimal amount;
		@Excel(name = "名称", orderNum = "2")
		private String name;

		public MergeDecimal(BigDecimal amount, String name) {
			this.amount = amount;
			this.name = name;
		}

		public BigDecimal getAmount() {
			return amount;
		}

		public void setAmount(BigDecimal amount) {
			this.amount = amount;
		}

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}
	}

	/** 名称列纵向合并，依赖数值列"组"：组不同的行不能被合并到一起 */
	public static class MergeRelyNumber {
		@Excel(name = "组", orderNum = "1", type = 4)
		private Integer group;
		@Excel(name = "名称", orderNum = "2", mergeVertical = true, mergeRely = {0})
		private String name;

		public MergeRelyNumber(Integer group, String name) {
			this.group = group;
			this.name = name;
		}

		public Integer getGroup() {
			return group;
		}

		public void setGroup(Integer group) {
			this.group = group;
		}

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}
	}

	public static class MergeText {
		@Excel(name = "类别", orderNum = "1", mergeVertical = true)
		private String category;
		@Excel(name = "名称", orderNum = "2")
		private String name;

		public MergeText(String category, String name) {
			this.category = category;
			this.name = name;
		}

		public String getCategory() {
			return category;
		}

		public void setCategory(String category) {
			this.category = category;
		}

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}
	}

	/** 返回纵向合并区域（只统计合并多行的、指定列上的区域） */
	private static List<CellRangeAddress> verticalRegions(Sheet sheet, int col) {
		List<CellRangeAddress> list = new ArrayList<CellRangeAddress>();
		for (CellRangeAddress r : sheet.getMergedRegions()) {
			if (r.getFirstColumn() == col && r.getLastColumn() == col && r.getLastRow() > r.getFirstRow()) {
				list.add(r);
			}
		}
		return list;
	}

	private static ExportParams params(ExcelType type) {
		ExportParams params = new ExportParams();
		params.setType(type);
		return params;
	}

	@Test
	public void numericColumnShouldMergeInHssf() {
		List<MergeNumber> list = new ArrayList<MergeNumber>(Arrays.asList(
				new MergeNumber(1, "a"), new MergeNumber(1, "b"), new MergeNumber(2, "c")));
		Workbook wb = ExcelExportUtil.exportExcel(params(ExcelType.HSSF), MergeNumber.class, list);
		List<CellRangeAddress> regions = verticalRegions(wb.getSheetAt(0), 0);
		assertEquals(1, regions.size());
		assertEquals(1, regions.get(0).getFirstRow());
		assertEquals(2, regions.get(0).getLastRow());
	}

	@Test
	public void numericColumnShouldMergeInXssf() {
		List<MergeNumber> list = new ArrayList<MergeNumber>(Arrays.asList(
				new MergeNumber(7, "a"), new MergeNumber(7, "b"), new MergeNumber(7, "c"), new MergeNumber(8, "d")));
		Workbook wb = ExcelExportUtil.exportExcel(params(ExcelType.XSSF), MergeNumber.class, list);
		List<CellRangeAddress> regions = verticalRegions(wb.getSheetAt(0), 0);
		assertEquals(1, regions.size());
		assertEquals(1, regions.get(0).getFirstRow());
		assertEquals(3, regions.get(0).getLastRow());
	}

	@Test
	public void zeroValuesShouldMergeToo() {
		List<MergeNumber> list = new ArrayList<MergeNumber>(Arrays.asList(
				new MergeNumber(0, "a"), new MergeNumber(0, "b")));
		Workbook wb = ExcelExportUtil.exportExcel(params(ExcelType.XSSF), MergeNumber.class, list);
		assertEquals(1, verticalRegions(wb.getSheetAt(0), 0).size());
	}

	@Test
	public void decimalColumnShouldMerge() {
		List<MergeDecimal> list = new ArrayList<MergeDecimal>(Arrays.asList(
				new MergeDecimal(new BigDecimal("1.5"), "a"), new MergeDecimal(new BigDecimal("1.5"), "b"),
				new MergeDecimal(new BigDecimal("2.5"), "c")));
		Workbook wb = ExcelExportUtil.exportExcel(params(ExcelType.XSSF), MergeDecimal.class, list);
		assertEquals(1, verticalRegions(wb.getSheetAt(0), 0).size());
	}

	@Test
	public void numericRelyColumnShouldNotBreakMerge() {
		// 名称都是 a，但第 3 行的"组"不同，不能和前两行合并
		List<MergeRelyNumber> list = new ArrayList<MergeRelyNumber>(Arrays.asList(
				new MergeRelyNumber(1, "a"), new MergeRelyNumber(1, "a"), new MergeRelyNumber(2, "a")));
		Workbook wb = ExcelExportUtil.exportExcel(params(ExcelType.XSSF), MergeRelyNumber.class, list);
		List<CellRangeAddress> regions = verticalRegions(wb.getSheetAt(0), 1);
		assertEquals(1, regions.size());
		assertEquals(1, regions.get(0).getFirstRow());
		assertEquals(2, regions.get(0).getLastRow());
	}

	@Test
	public void textColumnShouldStillMerge() {
		List<MergeText> list = new ArrayList<MergeText>(Arrays.asList(
				new MergeText("x", "a"), new MergeText("x", "b"), new MergeText("y", "c")));
		Workbook wb = ExcelExportUtil.exportExcel(params(ExcelType.HSSF), MergeText.class, list);
		assertEquals(1, verticalRegions(wb.getSheetAt(0), 0).size());
	}

	@Test
	public void directCallOnNumericSheetShouldMerge() {
		for (Workbook wb : new Workbook[]{new HSSFWorkbook(), new XSSFWorkbook()}) {
			Sheet sheet = wb.createSheet("s");
			double[] values = {5, 5, 5, 6};
			for (int i = 0; i < values.length; i++) {
				Row row = sheet.createRow(i);
				row.createCell(0).setCellValue(values[i]);
			}
			PoiMergeCellUtil.mergeCells(sheet, 0, 0);
			List<CellRangeAddress> regions = verticalRegions(sheet, 0);
			assertEquals(wb.getClass().getSimpleName(), 1, regions.size());
			assertEquals(0, regions.get(0).getFirstRow());
			assertEquals(2, regions.get(0).getLastRow());
		}
	}
}
