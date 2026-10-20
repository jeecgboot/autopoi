package org.jeecgframework.poi.excel.imports;

import org.apache.poi.ss.usermodel.CellStyle;
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
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

/**
 * SAX 导入时，单元格的类型必须由该单元格自己的 t 属性决定，不能沿用上一个单元格的类型。
 * 否则数字单元格紧跟在字符串单元格后面时，会被当成共享字符串的下标，取到别的单元格(如表头)的文本。
 *
 * Excel / WPS 保存数字单元格时写成 &lt;c r="B2"&gt;&lt;v&gt;3&lt;/v&gt;&lt;/c&gt;（没有 t、没有 .0），
 * POI 生成的文件是 &lt;c t="n" s="0"&gt;&lt;v&gt;3.0&lt;/v&gt;&lt;/c&gt;，"3.0" 恰好不是整数而没有触发问题，
 * 所以测试里按 Excel 的写法改写 sheet XML。
 *
 * @since 2026-10-20 issues/189 SAX 数字单元格取到表头的文本
 */
public class SaxImportCellTypeTest {

	public static class Row4 {
		@Excel(name = "名称")
		private String name;
		@Excel(name = "数量")
		private Integer qty;
		@Excel(name = "备注")
		private String remark;
		@Excel(name = "单价")
		private BigDecimal price;

		public String getName() { return name; }
		public void setName(String name) { this.name = name; }
		public Integer getQty() { return qty; }
		public void setQty(Integer qty) { this.qty = qty; }
		public String getRemark() { return remark; }
		public void setRemark(String remark) { this.remark = remark; }
		public BigDecimal getPrice() { return price; }
		public void setPrice(BigDecimal price) { this.price = price; }
	}

	public static class NumFirst {
		@Excel(name = "编号")
		private Integer no;
		@Excel(name = "名称")
		private String name;

		public Integer getNo() { return no; }
		public void setNo(Integer no) { this.no = no; }
		public String getName() { return name; }
		public void setName(String name) { this.name = name; }
	}

	public static class WithFormula {
		@Excel(name = "名称")
		private String name;
		@Excel(name = "公式文本")
		private String formulaText;

		public String getName() { return name; }
		public void setName(String name) { this.name = name; }
		public String getFormulaText() { return formulaText; }
		public void setFormulaText(String formulaText) { this.formulaText = formulaText; }
	}

	public static class WithDate {
		@Excel(name = "名称")
		private String name;
		@Excel(name = "日期", format = "yyyy-MM-dd")
		private Date day;
		@Excel(name = "数量")
		private Integer qty;

		public String getName() { return name; }
		public void setName(String name) { this.name = name; }
		public Date getDay() { return day; }
		public void setDay(Date day) { this.day = day; }
		public Integer getQty() { return qty; }
		public void setQty(Integer qty) { this.qty = qty; }
	}

	/** 把 POI 生成的 xlsx 的 sheet1.xml 改写成 Excel 保存数字单元格的方式：去掉 t="n"、s="0" 和 ".0" */
	private static InputStream asExcelWrites(XSSFWorkbook wb) throws Exception {
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		wb.write(bos);
		wb.close();
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		ZipInputStream zin = new ZipInputStream(new ByteArrayInputStream(bos.toByteArray()));
		ZipOutputStream zout = new ZipOutputStream(out);
		ZipEntry entry;
		while ((entry = zin.getNextEntry()) != null) {
			ByteArrayOutputStream content = new ByteArrayOutputStream();
			byte[] buf = new byte[4096];
			int n;
			while ((n = zin.read(buf)) > 0) {
				content.write(buf, 0, n);
			}
			byte[] data = content.toByteArray();
			if (entry.getName().startsWith("xl/worksheets/sheet")) {
				String xml = new String(data, StandardCharsets.UTF_8);
				xml = xml.replace(" t=\"n\"", "").replace(" s=\"0\"", "").replaceAll("<v>(-?\\d+)\\.0</v>", "<v>$1</v>");
				data = xml.getBytes(StandardCharsets.UTF_8);
			}
			zout.putNextEntry(new ZipEntry(entry.getName()));
			zout.write(data);
			zout.closeEntry();
		}
		zout.close();
		return new ByteArrayInputStream(out.toByteArray());
	}

	private static ImportParams params() {
		ImportParams params = new ImportParams();
		params.setSheetNum(1);
		return params;
	}

	private static XSSFWorkbook workbookOf(Object[]... rows) {
		XSSFWorkbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Sheet1");
		for (int r = 0; r < rows.length; r++) {
			Row row = sheet.createRow(r);
			for (int c = 0; c < rows[r].length; c++) {
				Object v = rows[r][c];
				if (v == null) {
					continue;
				}
				if (v instanceof Number) {
					row.createCell(c).setCellValue(((Number) v).doubleValue());
				} else {
					row.createCell(c).setCellValue(String.valueOf(v));
				}
			}
		}
		return wb;
	}

	@Test
	public void integerAfterStringShouldKeepItsValueWhenExcelWritesNoDotZero() throws Exception {
		XSSFWorkbook wb = workbookOf(
				new Object[]{"名称", "数量", "备注", "单价"},
				new Object[]{"苹果", 3, "好", 2.5},
				new Object[]{"香蕉", 12, "新鲜", 10.25});
		List<Row4> list = ExcelImportUtil.importExcelBySax(asExcelWrites(wb), Row4.class, params());
		assertEquals(2, list.size());
		assertEquals("苹果", list.get(0).getName());
		assertEquals(Integer.valueOf(3), list.get(0).getQty());
		assertEquals("好", list.get(0).getRemark());
		assertEquals(0, new BigDecimal("2.5").compareTo(list.get(0).getPrice()));
		assertEquals(Integer.valueOf(12), list.get(1).getQty());
		assertEquals(0, new BigDecimal("10.25").compareTo(list.get(1).getPrice()));
	}

	@Test
	public void numberAfterStringShouldNotBeTakenFromSharedStrings() throws Exception {
		// 数量 1、2 恰好是共享字符串表里前两项（表头 "名称"、"数量"）的下标
		XSSFWorkbook wb = workbookOf(
				new Object[]{"名称", "数量", "备注", "单价"},
				new Object[]{"a", 1, "b", 2});
		List<Row4> list = ExcelImportUtil.importExcelBySax(asExcelWrites(wb), Row4.class, params());
		assertEquals(Integer.valueOf(1), list.get(0).getQty());
		assertEquals(0, BigDecimal.valueOf(2).compareTo(list.get(0).getPrice()));
	}

	@Test
	public void numberAsFirstCellOfRowShouldBeImported() throws Exception {
		XSSFWorkbook wb = workbookOf(
				new Object[]{"编号", "名称"},
				new Object[]{7, "甲"},
				new Object[]{8, "乙"});
		List<NumFirst> list = ExcelImportUtil.importExcelBySax(asExcelWrites(wb), NumFirst.class, params());
		assertEquals(2, list.size());
		assertEquals(Integer.valueOf(7), list.get(0).getNo());
		assertEquals("甲", list.get(0).getName());
		assertEquals(Integer.valueOf(8), list.get(1).getNo());
	}

	@Test
	public void poiStyleNumbersShouldStillWork() throws Exception {
		XSSFWorkbook wb = workbookOf(
				new Object[]{"名称", "数量", "备注", "单价"},
				new Object[]{"苹果", 3, "好", 2.5});
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		wb.write(bos);
		List<Row4> list = ExcelImportUtil.importExcelBySax(new ByteArrayInputStream(bos.toByteArray()), Row4.class, params());
		assertEquals(Integer.valueOf(3), list.get(0).getQty());
	}

	/** 公式的字符串结果（t="str"）里 v 是文本本身，不是共享字符串的下标 */
	@Test
	public void formulaStringResultShouldNotBeTakenFromSharedStrings() throws Exception {
		XSSFWorkbook wb = workbookOf(
				new Object[]{"名称", "公式文本"},
				new Object[]{"甲", null});
		wb.getSheetAt(0).getRow(1).createCell(1).setCellFormula("\"1\"");
		wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
		List<WithFormula> list = ExcelImportUtil.importExcelBySax(asExcelWrites(wb), WithFormula.class, params());
		assertEquals(1, list.size());
		assertEquals("甲", list.get(0).getName());
		assertEquals("1", list.get(0).getFormulaText());
	}

	/** 日期单元格（沿用现有的 style 识别）和其后的数字都应正确 */
	@Test
	public void dateFollowedByNumberShouldBeImported() throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		CellStyle dateStyle = wb.createCellStyle();
		dateStyle.setDataFormat(wb.createDataFormat().getFormat("yyyy-mm-dd"));
		Sheet sheet = wb.createSheet("Sheet1");
		Row head = sheet.createRow(0);
		head.createCell(0).setCellValue("名称");
		head.createCell(1).setCellValue("日期");
		head.createCell(2).setCellValue("数量");
		Row row = sheet.createRow(1);
		row.createCell(0).setCellValue("甲");
		row.createCell(1).setCellValue(new Date(0));
		row.getCell(1).setCellStyle(dateStyle);
		row.createCell(2).setCellValue(2);
		List<WithDate> list = ExcelImportUtil.importExcelBySax(asExcelWrites(wb), WithDate.class, params());
		assertEquals(1, list.size());
		assertNotNull(list.get(0).getDay());
		assertEquals(Integer.valueOf(2), list.get(0).getQty());
	}
}
