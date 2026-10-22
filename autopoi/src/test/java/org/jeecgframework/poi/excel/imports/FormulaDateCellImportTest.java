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
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.Date;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

/**
 * [issues/40] 公式单元格的结果是日期（如 DATE(2026,1,2)、A1+1）时，导入到 Date 字段抛"Excel 值获取失败"：
 * 代码只判断 CellType.NUMERIC，公式单元格的类型是 FORMULA，于是走到"按文本解析"的分支，
 * 把日期序列号当成文本再按 format 解析，解析失败。
 */
public class FormulaDateCellImportTest {

	public static class Row1 {
		@Excel(name = "名称")
		private String name;
		@Excel(name = "日期", format = "yyyy-MM-dd")
		private Date day;
		@Excel(name = "本地日期", format = "yyyy-MM-dd")
		private LocalDate localDay;
		public String getName() { return name; }
		public void setName(String name) { this.name = name; }
		public Date getDay() { return day; }
		public void setDay(Date day) { this.day = day; }
		public LocalDate getLocalDay() { return localDay; }
		public void setLocalDay(LocalDate localDay) { this.localDay = localDay; }
	}

	private static InputStream toStream(XSSFWorkbook wb) throws Exception {
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		wb.write(bos);
		wb.close();
		return new ByteArrayInputStream(bos.toByteArray());
	}

	private static XSSFWorkbook workbook() {
		XSSFWorkbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("s");
		Row head = sheet.createRow(0);
		head.createCell(0).setCellValue("名称");
		head.createCell(1).setCellValue("日期");
		head.createCell(2).setCellValue("本地日期");
		return wb;
	}

	private static List<Row1> read(XSSFWorkbook wb) throws Exception {
		wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
		return ExcelImportUtil.importExcel(toStream(wb), Row1.class, new ImportParams());
	}

	private static String format(Date date) {
		return new SimpleDateFormat("yyyy-MM-dd").format(date);
	}

	/** 公式 DATE(2026,1,2) 的结果是日期：以前抛异常 */
	@Test
	public void dateFormulaShouldImportIntoDateField() throws Exception {
		XSSFWorkbook wb = workbook();
		Row row = wb.getSheetAt(0).createRow(1);
		row.createCell(0).setCellValue("a");
		row.createCell(1).setCellFormula("DATE(2026,1,2)");
		List<Row1> list = read(wb);
		assertEquals(1, list.size());
		assertNotNull(list.get(0).getDay());
		assertEquals("2026-01-02", format(list.get(0).getDay()));
	}

	/** 引用其它日期单元格再加天数：A=普通日期单元格，B=A+1 */
	@Test
	public void formulaReferencingDateCellShouldImport() throws Exception {
		XSSFWorkbook wb = workbook();
		Row row = wb.getSheetAt(0).createRow(1);
		row.createCell(0).setCellValue("a");
		row.createCell(3).setCellFormula("DATE(2026,3,1)");
		row.createCell(1).setCellFormula("D2+10");
		List<Row1> list = read(wb);
		assertEquals("2026-03-11", format(list.get(0).getDay()));
	}

	/** 其它日期类型（LocalDate）同样要支持 */
	@Test
	public void dateFormulaShouldImportIntoLocalDateField() throws Exception {
		XSSFWorkbook wb = workbook();
		Row row = wb.getSheetAt(0).createRow(1);
		row.createCell(0).setCellValue("a");
		row.createCell(2).setCellFormula("DATE(2026,5,6)");
		List<Row1> list = read(wb);
		assertEquals(LocalDate.of(2026, 5, 6), list.get(0).getLocalDay());
	}

	/** 公式的结果是文本日期（TEXT(...)）：仍按 format 解析，行为不变 */
	@Test
	public void textFormulaResultShouldStillBeParsedByFormat() throws Exception {
		XSSFWorkbook wb = workbook();
		Row row = wb.getSheetAt(0).createRow(1);
		row.createCell(0).setCellValue("a");
		row.createCell(1).setCellFormula("TEXT(DATE(2026,7,8),\"yyyy-mm-dd\")");
		List<Row1> list = read(wb);
		assertEquals("2026-07-08", format(list.get(0).getDay()));
	}

	/** 原来的用法不变：普通日期数值单元格、文本日期单元格 */
	@Test
	public void plainDateCellsShouldStillImport() throws Exception {
		XSSFWorkbook wb = workbook();
		Row r1 = wb.getSheetAt(0).createRow(1);
		r1.createCell(0).setCellValue("a");
		r1.createCell(1).setCellValue(java.sql.Date.valueOf("2026-09-10"));
		Row r2 = wb.getSheetAt(0).createRow(2);
		r2.createCell(0).setCellValue("b");
		r2.createCell(1).setCellValue("2026-11-12");
		List<Row1> list = read(wb);
		assertEquals("2026-09-10", format(list.get(0).getDay()));
		assertEquals("2026-11-12", format(list.get(1).getDay()));
	}
}
