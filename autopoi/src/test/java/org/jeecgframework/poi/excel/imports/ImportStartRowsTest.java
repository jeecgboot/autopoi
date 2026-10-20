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
 * ImportParams.startRows（表头后需要跳过的行数）回归测试，DOM 和 SAX 两种导入模式。
 *
 * @since 2026-10-20 issues/137 ImportParams.startRows 不生效
 */
public class ImportStartRowsTest {

	public static class NameEntity {
		@Excel(name = "名称")
		private String name;

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}
	}

	/** 表头 + 3 行数据：skip1、skip2、keep */
	private static InputStream threeRows() throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Sheet1");
		sheet.createRow(0).createCell(0).setCellValue("名称");
		String[] values = {"skip1", "skip2", "keep"};
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
	public void domImportShouldSkipStartRows() throws Exception {
		ImportParams params = new ImportParams();
		params.setStartRows(2);
		List<NameEntity> list = ExcelImportUtil.importExcel(threeRows(), NameEntity.class, params);
		assertEquals(1, list.size());
		assertEquals("keep", list.get(0).getName());
	}

	@Test
	public void domImportWithoutStartRowsShouldReadAllRows() throws Exception {
		List<NameEntity> list = ExcelImportUtil.importExcel(threeRows(), NameEntity.class, new ImportParams());
		assertEquals(3, list.size());
	}

	@Test
	public void saxImportShouldSkipStartRows() throws Exception {
		ImportParams params = new ImportParams();
		params.setSheetNum(1);
		params.setStartRows(2);
		List<NameEntity> list = ExcelImportUtil.importExcelBySax(threeRows(), NameEntity.class, params);
		assertEquals(1, list.size());
		assertEquals("keep", list.get(0).getName());
	}
}
