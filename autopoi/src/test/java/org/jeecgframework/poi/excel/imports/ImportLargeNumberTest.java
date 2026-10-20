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
 * 数值单元格导入到 Integer / Long 字段的回归测试：
 * 8 位及以上的整数会被 POI 取成 Double，toString 变成科学计数法（如 1.2345678E7），不能直接解析为整数。
 *
 * @since 2026-10-20 issues/120 导入Excel功能中8位数的整形数据报错
 */
public class ImportLargeNumberTest {

	public static class NumberEntity {
		@Excel(name = "整数")
		private Integer intValue;
		@Excel(name = "长整数")
		private Long longValue;

		public Integer getIntValue() {
			return intValue;
		}

		public void setIntValue(Integer intValue) {
			this.intValue = intValue;
		}

		public Long getLongValue() {
			return longValue;
		}

		public void setLongValue(Long longValue) {
			this.longValue = longValue;
		}
	}

	private static InputStream sheetOf(double intCell, double longCell) throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Sheet1");
		Row head = sheet.createRow(0);
		head.createCell(0).setCellValue("整数");
		head.createCell(1).setCellValue("长整数");
		Row row = sheet.createRow(1);
		row.createCell(0).setCellValue(intCell);
		row.createCell(1).setCellValue(longCell);
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		wb.write(bos);
		wb.close();
		return new ByteArrayInputStream(bos.toByteArray());
	}

	private static NumberEntity importOne(double intCell, double longCell) throws Exception {
		List<NumberEntity> list = ExcelImportUtil.importExcel(sheetOf(intCell, longCell), NumberEntity.class, new ImportParams());
		assertEquals(1, list.size());
		return list.get(0);
	}

	@Test
	public void eightDigitIntegerShouldImport() throws Exception {
		NumberEntity e = importOne(12345678, 12345678);
		assertEquals(Integer.valueOf(12345678), e.getIntValue());
		assertEquals(Long.valueOf(12345678L), e.getLongValue());
	}

	@Test
	public void negativeEightDigitIntegerShouldImport() throws Exception {
		NumberEntity e = importOne(-12345678, -12345678);
		assertEquals(Integer.valueOf(-12345678), e.getIntValue());
		assertEquals(Long.valueOf(-12345678L), e.getLongValue());
	}

	@Test
	public void longValueBeyondIntRangeShouldImport() throws Exception {
		NumberEntity e = importOne(1, 123456789012L);
		assertEquals(Long.valueOf(123456789012L), e.getLongValue());
	}

	@Test
	public void smallNumbersShouldStillImport() throws Exception {
		NumberEntity e = importOne(5, 0);
		assertEquals(Integer.valueOf(5), e.getIntValue());
		assertEquals(Long.valueOf(0L), e.getLongValue());
	}
}
