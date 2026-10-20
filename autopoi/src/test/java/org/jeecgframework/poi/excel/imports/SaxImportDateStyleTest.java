package org.jeecgframework.poi.excel.imports;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
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
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

/**
 * SAX 导入时，单元格是不是日期必须由样式的数字格式决定，不能只认 style 下标为 1 的样式。
 * 真实的 Excel 文件里日期样式的下标取决于文件定义了多少种样式（常常是 s="3"、s="5"…），
 * 并且下标为 1 的样式也可能根本不是日期。
 *
 * @since 2026-10-20 issues/190 SAX 日期只认 style=1，其它日期样式被当成数字
 */
public class SaxImportDateStyleTest {

	public static class DateRow {
		@Excel(name = "名称")
		private String name;
		@Excel(name = "日期")
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

	/** 日期列是 String 字段，指定 format */
	public static class StringDateWithFormat {
		@Excel(name = "日期", format = "yyyy/MM/dd")
		private String day;

		public String getDay() { return day; }
		public void setDay(String day) { this.day = day; }
	}

	/** 日期列是 String 字段，没有指定 format */
	public static class StringDateNoFormat {
		@Excel(name = "日期")
		private String day;

		public String getDay() { return day; }
		public void setDay(String day) { this.day = day; }
	}

	public static class NumRow {
		@Excel(name = "数量")
		private Integer qty;
		@Excel(name = "金额")
		private BigDecimal amount;

		public Integer getQty() { return qty; }
		public void setQty(Integer qty) { this.qty = qty; }
		public BigDecimal getAmount() { return amount; }
		public void setAmount(BigDecimal amount) { this.amount = amount; }
	}

	private static Date date(String text) throws Exception {
		return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").parse(text);
	}

	/** 先建几个和日期无关的样式，让日期样式的下标不是 1（和真实 Excel 文件一致） */
	private static void addUnrelatedStyles(XSSFWorkbook wb, int count) {
		for (int i = 0; i < count; i++) {
			CellStyle s = wb.createCellStyle();
			s.setWrapText(i % 2 == 0);
			s.setIndention((short) (i + 1));
		}
	}

	private static CellStyle styleOf(XSSFWorkbook wb, String format) {
		CellStyle s = wb.createCellStyle();
		DataFormat df = wb.createDataFormat();
		s.setDataFormat(df.getFormat(format));
		return s;
	}

	/** 把 POI 生成的 xlsx 的 sheet XML 改写成 Excel 保存数字单元格的方式：去掉 t="n"、s="0" 和 ".0" */
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
				xml = xml.replace(" t=\"n\"", "").replace(" s=\"0\"", "");
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

	/** 名称 | 日期 | 数量 三列，日期列使用 dateFormat 样式 */
	private static XSSFWorkbook dateWorkbook(int unrelatedStyles, String dateFormat, Date value) {
		XSSFWorkbook wb = new XSSFWorkbook();
		addUnrelatedStyles(wb, unrelatedStyles);
		CellStyle dateStyle = styleOf(wb, dateFormat);
		Sheet sheet = wb.createSheet("Sheet1");
		Row head = sheet.createRow(0);
		head.createCell(0).setCellValue("名称");
		head.createCell(1).setCellValue("日期");
		head.createCell(2).setCellValue("数量");
		Row row = sheet.createRow(1);
		row.createCell(0).setCellValue("甲");
		row.createCell(1).setCellValue(value);
		row.getCell(1).setCellStyle(dateStyle);
		row.createCell(2).setCellValue(7);
		return wb;
	}

	@Test
	public void dateStyleWithIndexNotOneShouldBeDate() throws Exception {
		Date expected = date("2026-01-31 19:28:40");
		XSSFWorkbook wb = dateWorkbook(3, "yyyy-mm-dd hh:mm:ss", expected);
		List<DateRow> list = ExcelImportUtil.importExcelBySax(asExcelWrites(wb), DateRow.class, params());
		assertEquals(1, list.size());
		assertNotNull("日期样式下标不是 1 时也应识别为日期", list.get(0).getDay());
		assertEquals(expected, list.get(0).getDay());
		assertEquals("日期后面的数字不受影响", Integer.valueOf(7), list.get(0).getQty());
	}

	@Test
	public void builtInDateFormatsShouldBeDate() throws Exception {
		// 内置日期格式：14 = m/d/yy，22 = m/d/yy h:mm
		for (String format : new String[]{"m/d/yy", "m/d/yy h:mm", "d-mmm-yy"}) {
			Date expected = date("2025-12-01 08:30:00");
			XSSFWorkbook wb = dateWorkbook(2, format, expected);
			List<DateRow> list = ExcelImportUtil.importExcelBySax(asExcelWrites(wb), DateRow.class, params());
			assertNotNull(format + " 应识别为日期", list.get(0).getDay());
		}
	}

	@Test
	public void customChineseDateFormatShouldBeDate() throws Exception {
		Date expected = date("2026-03-09 00:00:00");
		XSSFWorkbook wb = dateWorkbook(4, "yyyy\"年\"m\"月\"d\"日\"", expected);
		List<DateRow> list = ExcelImportUtil.importExcelBySax(asExcelWrites(wb), DateRow.class, params());
		assertEquals(expected, list.get(0).getDay());
	}

	@Test
	public void dateStyleWithIndexOneShouldStillBeDate() throws Exception {
		Date expected = date("2026-01-31 19:28:40");
		XSSFWorkbook wb = dateWorkbook(0, "yyyy-mm-dd hh:mm:ss", expected);
		List<DateRow> list = ExcelImportUtil.importExcelBySax(asExcelWrites(wb), DateRow.class, params());
		assertEquals(expected, list.get(0).getDay());
	}

	/** 下标为 1 的样式不是日期（比如只设置了自动换行）时，数值不能被当成日期 */
	@Test
	public void styleIndexOneThatIsNotDateShouldStayNumber() throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		CellStyle wrap = wb.createCellStyle();
		wrap.setWrapText(true);
		Sheet sheet = wb.createSheet("Sheet1");
		sheet.createRow(0).createCell(0).setCellValue("数量");
		sheet.getRow(0).createCell(1).setCellValue("金额");
		Row row = sheet.createRow(1);
		row.createCell(0).setCellValue(5);
		row.getCell(0).setCellStyle(wrap);
		row.createCell(1).setCellValue(2.5);
		List<NumRow> list = ExcelImportUtil.importExcelBySax(asExcelWrites(wb), NumRow.class, params());
		assertEquals(1, list.size());
		assertEquals(Integer.valueOf(5), list.get(0).getQty());
		assertEquals(0, new BigDecimal("2.5").compareTo(list.get(0).getAmount()));
	}

	/** 数字、百分比、千分位、文本等非日期的数字格式都保持为数字 */
	@Test
	public void nonDateNumberFormatsShouldStayNumber() throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		addUnrelatedStyles(wb, 2);
		Sheet sheet = wb.createSheet("Sheet1");
		sheet.createRow(0).createCell(0).setCellValue("数量");
		sheet.getRow(0).createCell(1).setCellValue("金额");
		Row row = sheet.createRow(1);
		row.createCell(0).setCellValue(12);
		row.getCell(0).setCellStyle(styleOf(wb, "#,##0"));
		row.createCell(1).setCellValue(3.25);
		row.getCell(1).setCellStyle(styleOf(wb, "0.00%"));
		List<NumRow> list = ExcelImportUtil.importExcelBySax(asExcelWrites(wb), NumRow.class, params());
		assertEquals(Integer.valueOf(12), list.get(0).getQty());
		assertEquals(0, new BigDecimal("3.25").compareTo(list.get(0).getAmount()));
	}

	@Test
	public void stringFieldShouldGetDateFormattedByFormat() throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		addUnrelatedStyles(wb, 3);
		Sheet sheet = wb.createSheet("Sheet1");
		sheet.createRow(0).createCell(0).setCellValue("日期");
		Row row = sheet.createRow(1);
		row.createCell(0).setCellValue(date("2026-03-09 10:20:30"));
		row.getCell(0).setCellStyle(styleOf(wb, "yyyy-mm-dd hh:mm:ss"));
		List<StringDateWithFormat> list = ExcelImportUtil.importExcelBySax(asExcelWrites(wb), StringDateWithFormat.class, params());
		assertEquals("2026/03/09", list.get(0).getDay());
	}

	@Test
	public void stringFieldWithoutFormatShouldGetDefaultDateTimeText() throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		addUnrelatedStyles(wb, 3);
		Sheet sheet = wb.createSheet("Sheet1");
		sheet.createRow(0).createCell(0).setCellValue("日期");
		Row row = sheet.createRow(1);
		row.createCell(0).setCellValue(date("2026-03-09 10:20:30"));
		row.getCell(0).setCellStyle(styleOf(wb, "yyyy-mm-dd hh:mm:ss"));
		List<StringDateNoFormat> list = ExcelImportUtil.importExcelBySax(asExcelWrites(wb), StringDateNoFormat.class, params());
		assertEquals("2026-03-09 10:20:30", list.get(0).getDay());
	}
}
