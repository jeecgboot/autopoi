package org.jeecgframework.poi.excel.imports;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.excel.ExcelImportUtil;
import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.jeecgframework.poi.excel.entity.sax.SaxReadCellEntity;
import org.jeecgframework.poi.excel.imports.sax.parse.ISaxRowRead;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.assertEquals;

/**
 * SAX 导入对 ImportParams 的处理：
 * 1. ImportParams 为 null 不应报空指针；
 * 2. 未设置 sheetNum（默认 0）时，与 DOM 导入一致，读取全部 sheet。
 *
 * @since 2026-10-20 issues/91
 */
public class SaxImportParamsTest {

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

	/** 两个 sheet，每个 sheet：表头 + 1 行数据（s1-a / s2-a） */
	private static InputStream twoSheets() throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		for (int s = 1; s <= 2; s++) {
			Sheet sheet = wb.createSheet("Sheet" + s);
			sheet.createRow(0).createCell(0).setCellValue("名称");
			Row row = sheet.createRow(1);
			row.createCell(0).setCellValue("s" + s + "-a");
		}
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		wb.write(bos);
		wb.close();
		return new ByteArrayInputStream(bos.toByteArray());
	}

	@Test
	public void defaultSheetNumShouldReadAllSheets() throws Exception {
		List<NameEntity> list = ExcelImportUtil.importExcelBySax(twoSheets(), NameEntity.class, new ImportParams());
		assertEquals(2, list.size());
		assertEquals("s1-a", list.get(0).getName());
		assertEquals("s2-a", list.get(1).getName());
	}

	@Test
	public void sheetNumOneShouldOnlyReadFirstSheet() throws Exception {
		ImportParams params = new ImportParams();
		params.setSheetNum(1);
		List<NameEntity> list = ExcelImportUtil.importExcelBySax(twoSheets(), NameEntity.class, params);
		assertEquals(1, list.size());
		assertEquals("s1-a", list.get(0).getName());
	}

	@Test
	public void sheetNumTwoShouldReadTwoSheets() throws Exception {
		ImportParams params = new ImportParams();
		params.setSheetNum(2);
		List<NameEntity> list = ExcelImportUtil.importExcelBySax(twoSheets(), NameEntity.class, params);
		assertEquals(2, list.size());
	}

	@Test
	public void sheetNumLargerThanSheetCountShouldNotFail() throws Exception {
		ImportParams params = new ImportParams();
		params.setSheetNum(10);
		List<NameEntity> list = ExcelImportUtil.importExcelBySax(twoSheets(), NameEntity.class, params);
		assertEquals(2, list.size());
	}

	@Test
	public void sheetNumShouldNotBeModifiedByImport() throws Exception {
		ImportParams params = new ImportParams();
		ExcelImportUtil.importExcelBySax(twoSheets(), NameEntity.class, params);
		assertEquals("导入不应修改调用方传入的参数", 0, params.getSheetNum());
	}

	@Test
	public void nullParamsWithEntityShouldReadAllSheets() throws Exception {
		List<NameEntity> list = ExcelImportUtil.importExcelBySax(twoSheets(), NameEntity.class, (ImportParams) null);
		assertEquals(2, list.size());
	}

	@Test
	public void customRowReadWithoutParamsShouldWork() throws Exception {
		final List<Integer> rows = new ArrayList<Integer>();
		ExcelImportUtil.importExcelBySax(twoSheets(), new ISaxRowRead() {
			@Override
			public <T> List<T> getList() {
				return null;
			}

			@Override
			public void parse(int index, List<SaxReadCellEntity> datas) {
				rows.add(index);
			}
		});
		// 两个 sheet，每个 sheet 回调表头行(0)和数据行(1)
		assertEquals(4, rows.size());
	}
}
