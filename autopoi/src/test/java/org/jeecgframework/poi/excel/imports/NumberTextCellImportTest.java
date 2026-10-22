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
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

/**
 * [issues/190] 数字类型字段遇到文本值：前后空格去掉、全空格按空值、报错信息带列名/原始值且不可见字符可见。
 */
public class NumberTextCellImportTest {

	public static class NumEntity {
		@Excel(name = "金额")
		private BigDecimal amount;
		@Excel(name = "数量")
		private Integer qty;
		@Excel(name = "总数")
		private Long total;
		@Excel(name = "比例")
		private Double rate;
		public BigDecimal getAmount() { return amount; }
		public void setAmount(BigDecimal amount) { this.amount = amount; }
		public Integer getQty() { return qty; }
		public void setQty(Integer qty) { this.qty = qty; }
		public Long getTotal() { return total; }
		public void setTotal(Long total) { this.total = total; }
		public Double getRate() { return rate; }
		public void setRate(Double rate) { this.rate = rate; }
	}

	private static final String[] HEADERS = {"金额", "数量", "总数", "比例"};

	/** 第一行是表头，其余是数据；String→字符串单元格，Number→数值单元格 */
	private static XSSFWorkbook sheetOf(String[] headers, Object[]... rows) {
		XSSFWorkbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Sheet1");
		Row head = sheet.createRow(0);
		for (int i = 0; i < headers.length; i++) {
			head.createCell(i).setCellValue(headers[i]);
		}
		for (int r = 0; r < rows.length; r++) {
			Row row = sheet.createRow(r + 1);
			for (int c = 0; c < rows[r].length; c++) {
				Object v = rows[r][c];
				if (v instanceof Number) {
					row.createCell(c).setCellValue(((Number) v).doubleValue());
				} else {
					row.createCell(c).setCellValue(String.valueOf(v));
				}
			}
		}
		return wb;
	}

	private static InputStream toStream(XSSFWorkbook wb) throws IOException {
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

	private static List<NumEntity> importDom(Object[]... rows) throws Exception {
		XSSFWorkbook wb = sheetOf(HEADERS, rows);
		return ExcelImportUtil.importExcel(toStream(wb), NumEntity.class, params());
	}

	private static List<NumEntity> importSax(Object[]... rows) throws Exception {
		XSSFWorkbook wb = sheetOf(HEADERS, rows);
		return ExcelImportUtil.importExcelBySax(toStream(wb), NumEntity.class, params());
	}

	@Test
	public void dom_textWithSurroundingSpacesShouldBeTrimmed() throws Exception {
		List<NumEntity> list = importDom(new Object[]{" 12.5 ", "  7", "8  ", " 1.5 "});
		assertEquals(1, list.size());
		assertEquals(new BigDecimal("12.5"), list.get(0).getAmount());
		assertEquals(Integer.valueOf(7), list.get(0).getQty());
		assertEquals(Long.valueOf(8), list.get(0).getTotal());
		assertEquals(Double.valueOf(1.5), list.get(0).getRate());
	}

	@Test
	public void dom_blankTextShouldBeNull() throws Exception {
		List<NumEntity> list = importDom(new Object[]{"   ", "  ", " ", "  "}, new Object[]{"  ", 2, 3, 4});
		// 整行都是空白会被当成空行跳过，所以第一行不会出现在结果里
		assertEquals(1, list.size());
		assertNull(list.get(0).getAmount());
		assertEquals(Integer.valueOf(2), list.get(0).getQty());
	}

	@Test
	public void dom_errorMessageShowsColumnValueAndInvisibleChar() throws Exception {
		//   不间断空格：String.trim() 去不掉，肉眼也看不出来
		try {
			importDom(new Object[]{"5\u00A0", 1, 1, 1});
			fail("不间断空格无法转成 BigDecimal，应该抛出异常");
		} catch (ExcelImportException e) {
			assertTrue(e.getMessage(), e.getMessage().contains("金额"));
			assertTrue(e.getMessage(), e.getMessage().contains("5\\u00A0"));
			assertTrue(e.getMessage(), e.getMessage().contains("java.math.BigDecimal"));
		}
	}

	@Test
	public void sax_errorMessageShowsColumnValueAndInvisibleChar() throws Exception {
		try {
			importSax(new Object[]{1, "3\u200B", 1, 1});
			fail("零宽空格无法转成 Integer，应该抛出异常");
		} catch (ExcelImportException e) {
			assertTrue(e.getMessage(), e.getMessage().contains("数量"));
			assertTrue(e.getMessage(), e.getMessage().contains("3\\u200B"));
			assertTrue(e.getMessage(), e.getMessage().contains("java.lang.Integer"));
		}
	}

	@Test
	public void sax_normalNumbersStillImport() throws Exception {
		List<NumEntity> list = importSax(new Object[]{12.5, 7, 8, 1.5});
		assertEquals(1, list.size());
		assertEquals(0, new BigDecimal("12.5").compareTo(list.get(0).getAmount()));
		assertEquals(Integer.valueOf(7), list.get(0).getQty());
		assertEquals(Long.valueOf(8), list.get(0).getTotal());
		assertEquals(Double.valueOf(1.5), list.get(0).getRate());
	}
}
