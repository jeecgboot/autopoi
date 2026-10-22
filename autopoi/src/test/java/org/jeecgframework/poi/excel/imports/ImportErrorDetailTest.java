package org.jeecgframework.poi.excel.imports;

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
import java.text.ParseException;
import java.util.Date;

import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

/**
 * [issues/160] 值转换失败时，往外抛的异常要带上列名和原始值：
 * 日期按 format 解析失败时，以前只抛“Excel 值获取失败”，既没有列名也没有值，也没有原始异常。
 */
public class ImportErrorDetailTest {

	public static class DateRow {
		@Excel(name = "开票日期", format = "yyyy-MM-dd")
		private Date day;
		public Date getDay() { return day; }
		public void setDay(Date day) { this.day = day; }
	}

	private static InputStream workbook(String head, String value) throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("s");
		sheet.createRow(0).createCell(0).setCellValue(head);
		Row row = sheet.createRow(1);
		row.createCell(0).setCellValue(value);
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		wb.write(bos);
		wb.close();
		return new ByteArrayInputStream(bos.toByteArray());
	}

	/** 日期文本无法按 format 解析：异常信息要有列名、原始值和 format，并保留 ParseException */
	@Test
	public void dateParseFailureShouldReportColumnValueAndFormat() throws Exception {
		try {
			ExcelImportUtil.importExcel(workbook("开票日期", "2026/01/02"), DateRow.class, new ImportParams());
			fail("无法按 yyyy-MM-dd 解析，应该抛出异常");
		} catch (ExcelImportException e) {
			String message = e.getMessage();
			assertTrue(message, message.contains("开票日期"));
			assertTrue(message, message.contains("2026/01/02"));
			assertTrue(message, message.contains("yyyy-MM-dd"));
			Throwable cause = e.getCause();
			while (cause != null && !(cause instanceof ParseException)) {
				cause = cause.getCause();
			}
			assertNotNull("应保留原始的 ParseException", cause);
		}
	}

	/** 日期文本里有不可见字符（如不间断空格）：信息里要能看出来 */
	@Test
	public void dateParseFailureShouldShowInvisibleCharacters() throws Exception {
		try {
			ExcelImportUtil.importExcel(workbook("开票日期", "2026\u00A0-01-02"), DateRow.class, new ImportParams());
			fail("应该抛出异常");
		} catch (ExcelImportException e) {
			assertTrue(e.getMessage(), e.getMessage().contains("\\u00A0"));
		}
	}
}
