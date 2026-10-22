package org.jeecgframework.poi.excel.imports;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.excel.ExcelImportUtil;
import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.jeecgframework.poi.exception.excel.ExcelImportException;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.fail;

/**
 * 日期以字符串形式保存在单元格里时（如 "2026-01-31 19:28:40"），SAX 导入应按 @Excel(format) 转成 Date，
 * 与 DOM 导入一致。
 *
 * @since 2026-10-20 issues/190 SAX 下字符串日期没有按 format 转 Date，反射赋值类型不匹配导致导入中断
 */
public class SaxImportStringDateTest {

	public static class DateEntity {
		@Excel(name = "名称")
		private String name;
		@Excel(name = "日期", format = "yyyy-MM-dd HH:mm:ss")
		private Date day;

		public String getName() { return name; }
		public void setName(String name) { this.name = name; }
		public Date getDay() { return day; }
		public void setDay(Date day) { this.day = day; }
	}

	public static class DayEntity {
		@Excel(name = "日期", format = "yyyy/MM/dd")
		private Date day;

		public Date getDay() { return day; }
		public void setDay(Date day) { this.day = day; }
	}

	private static Date parse(String pattern, String text) throws Exception {
		return new SimpleDateFormat(pattern).parse(text);
	}

	private static InputStream sheetOf(String header, String... values) throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Sheet1");
		Row head = sheet.createRow(0);
		head.createCell(0).setCellValue("名称");
		head.createCell(1).setCellValue(header);
		for (int i = 0; i < values.length; i++) {
			Row row = sheet.createRow(i + 1);
			row.createCell(0).setCellValue("n" + (i + 1));
			if (values[i] != null) {
				row.createCell(1).setCellValue(values[i]);
			}
		}
		return toStream(wb);
	}

	private static InputStream toStream(XSSFWorkbook wb) throws Exception {
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
	public void stringDateShouldBeParsedByFormatInSax() throws Exception {
		List<DateEntity> list = ExcelImportUtil.importExcelBySax(
				sheetOf("日期", "2026-01-31 19:28:40", "2025-12-01 00:00:01"), DateEntity.class, saxParams());
		assertEquals(2, list.size());
		assertEquals(parse("yyyy-MM-dd HH:mm:ss", "2026-01-31 19:28:40"), list.get(0).getDay());
		assertEquals(parse("yyyy-MM-dd HH:mm:ss", "2025-12-01 00:00:01"), list.get(1).getDay());
	}

	@Test
	public void stringDateShouldBeParsedByDifferentFormat() throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Sheet1");
		sheet.createRow(0).createCell(0).setCellValue("日期");
		sheet.createRow(1).createCell(0).setCellValue("2026/03/09");
		List<DayEntity> list = ExcelImportUtil.importExcelBySax(toStream(wb), DayEntity.class, saxParams());
		assertEquals(parse("yyyy/MM/dd", "2026/03/09"), list.get(0).getDay());
	}

	@Test
	public void saxResultShouldEqualDomResult() throws Exception {
		List<DateEntity> sax = ExcelImportUtil.importExcelBySax(
				sheetOf("日期", "2026-01-31 19:28:40"), DateEntity.class, saxParams());
		List<DateEntity> dom = ExcelImportUtil.importExcel(
				sheetOf("日期", "2026-01-31 19:28:40"), DateEntity.class, new ImportParams());
		assertEquals(dom.get(0).getDay(), sax.get(0).getDay());
	}

	@Test
	public void invalidStringDateShouldFailWithImportException() throws Exception {
		try {
			ExcelImportUtil.importExcelBySax(sheetOf("日期", "不是日期"), DateEntity.class, saxParams());
			fail("无法按 format 解析的日期应抛出 ExcelImportException");
		} catch (ExcelImportException expected) {
			// ok，与 DOM 导入一致
		}
	}

	@Test
	public void emptyDateCellShouldBeNull() throws Exception {
		List<DateEntity> list = ExcelImportUtil.importExcelBySax(
				sheetOf("日期", "2026-01-31 19:28:40", null), DateEntity.class, saxParams());
		assertEquals(2, list.size());
		assertNotNull(list.get(0).getDay());
		assertNull(list.get(1).getDay());
	}

	/** 文本里的前后空格不影响解析 */
	@Test
	public void stringDateWithBlanksShouldBeParsed() throws Exception {
		List<DateEntity> list = ExcelImportUtil.importExcelBySax(
				sheetOf("日期", "  2026-01-31 19:28:40  "), DateEntity.class, saxParams());
		assertEquals(parse("yyyy-MM-dd HH:mm:ss", "2026-01-31 19:28:40"), list.get(0).getDay());
	}

	/** 用日期样式保存的真正的日期单元格（数值）仍然得到 Date，不受影响 */
	@Test
	public void realDateCellShouldStillWork() throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		CellStyle dateStyle = wb.createCellStyle();
		dateStyle.setDataFormat(wb.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
		Sheet sheet = wb.createSheet("Sheet1");
		Row head = sheet.createRow(0);
		head.createCell(0).setCellValue("名称");
		head.createCell(1).setCellValue("日期");
		Row row = sheet.createRow(1);
		row.createCell(0).setCellValue("n1");
		row.createCell(1).setCellValue(parse("yyyy-MM-dd HH:mm:ss", "2026-01-31 19:28:40"));
		row.getCell(1).setCellStyle(dateStyle);
		List<DateEntity> list = ExcelImportUtil.importExcelBySax(toStream(wb), DateEntity.class, saxParams());
		assertEquals(1, list.size());
		assertNotNull(list.get(0).getDay());
	}
}
