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
 * 导入时字典(replace)翻译，单元格值前后带空白（普通空格、不间断空格、全角空格）也应能匹配。
 *
 * @since 2026-10-20 issues/124 Integer 字段导入字典翻译，值带空格不会自动过滤
 */
public class ImportDictTrimTest {

	public static class SexInt {
		@Excel(name = "性别", replace = {"男_1", "女_2"})
		private Integer sex;

		public Integer getSex() {
			return sex;
		}

		public void setSex(Integer sex) {
			this.sex = sex;
		}
	}

	public static class SexStr {
		@Excel(name = "性别", replace = {"男_1", "女_2"})
		private String sex;

		public String getSex() {
			return sex;
		}

		public void setSex(String sex) {
			this.sex = sex;
		}
	}

	private static InputStream sheetOf(String... values) throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Sheet1");
		sheet.createRow(0).createCell(0).setCellValue("性别");
		for (int i = 0; i < values.length; i++) {
			Row row = sheet.createRow(i + 1);
			row.createCell(0).setCellValue(values[i]);
		}
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		wb.write(bos);
		wb.close();
		return new ByteArrayInputStream(bos.toByteArray());
	}

	@Test
	public void integerFieldShouldTranslateValueWithTrailingSpace() throws Exception {
		List<SexInt> list = ExcelImportUtil.importExcel(sheetOf("女 "), SexInt.class, new ImportParams());
		assertEquals(Integer.valueOf(2), list.get(0).getSex());
	}

	@Test
	public void integerFieldShouldTranslateValueWithVariousBlanks() throws Exception {
		List<SexInt> list = ExcelImportUtil.importExcel(
				sheetOf(" 男", "男 ", "\u3000女", "女\u3000 ", "\t男\t", "\u00A0男\u00A0"), SexInt.class, new ImportParams());
		assertEquals(6, list.size());
		assertEquals(Integer.valueOf(1), list.get(0).getSex());
		assertEquals(Integer.valueOf(1), list.get(1).getSex());
		assertEquals(Integer.valueOf(2), list.get(2).getSex());
		assertEquals(Integer.valueOf(2), list.get(3).getSex());
		assertEquals(Integer.valueOf(1), list.get(4).getSex());
		assertEquals(Integer.valueOf(1), list.get(5).getSex());
	}

	@Test
	public void stringFieldShouldTranslateValueWithBlanks() throws Exception {
		List<SexStr> list = ExcelImportUtil.importExcel(sheetOf("男 ", " 女"), SexStr.class, new ImportParams());
		assertEquals("1", list.get(0).getSex());
		assertEquals("2", list.get(1).getSex());
	}

	@Test
	public void multiValueWithBlanksShouldTranslateEachValue() throws Exception {
		List<SexStr> list = ExcelImportUtil.importExcel(sheetOf("男, 女"), SexStr.class, new ImportParams());
		assertEquals("1,2", list.get(0).getSex());
	}

	@Test
	public void valueWithoutBlanksShouldStillTranslate() throws Exception {
		List<SexInt> list = ExcelImportUtil.importExcel(sheetOf("男", "女"), SexInt.class, new ImportParams());
		assertEquals(Integer.valueOf(1), list.get(0).getSex());
		assertEquals(Integer.valueOf(2), list.get(1).getSex());
	}

	@Test
	public void valueNotInDictShouldBeKeptAsIs() throws Exception {
		List<SexStr> list = ExcelImportUtil.importExcel(sheetOf("未知"), SexStr.class, new ImportParams());
		assertEquals("未知", list.get(0).getSex());
	}
}
