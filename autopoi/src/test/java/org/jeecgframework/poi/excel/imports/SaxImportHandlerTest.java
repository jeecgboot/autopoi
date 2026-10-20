package org.jeecgframework.poi.excel.imports;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.excel.ExcelImportUtil;
import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.jeecgframework.poi.excel.entity.sax.SaxReadCellEntity;
import org.jeecgframework.poi.excel.imports.sax.SaxReadExcel;
import org.jeecgframework.poi.excel.imports.sax.parse.ISaxRowRead;
import org.jeecgframework.poi.handler.inter.IExcelReadRowHanlder;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.junit.Assert.assertEquals;

/**
 * SAX 导入使用 IExcelReadRowHanlder 回调时，每一行数据都应被回调，包括最后一行。
 *
 * @since 2026-10-20 issues/189 SAX 回调模式少读取最后一行
 */
public class SaxImportHandlerTest {

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

	private static InputStream sheetOf(String... values) throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Sheet1");
		sheet.createRow(0).createCell(0).setCellValue("名称");
		for (int i = 0; i < values.length; i++) {
			Row row = sheet.createRow(i + 1);
			row.createCell(0).setCellValue(values[i]);
		}
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

	private static List<String> importWithHandler(String... values) throws Exception {
		final List<String> got = new ArrayList<String>();
		ExcelImportUtil.importExcelBySax(sheetOf(values), NameEntity.class, params(), new IExcelReadRowHanlder<NameEntity>() {
			@Override
			public void hanlder(NameEntity obj) {
				got.add(obj.getName());
			}
		});
		return got;
	}

	@Test
	public void allRowsShouldBeDeliveredInOrder() throws Exception {
		assertEquals(Arrays.asList("a1", "a2", "a3"), importWithHandler("a1", "a2", "a3"));
	}

	@Test
	public void singleRowShouldBeDelivered() throws Exception {
		assertEquals(Arrays.asList("only"), importWithHandler("only"));
	}

	@Test
	public void headerOnlyShouldDeliverNothing() throws Exception {
		assertEquals(0, importWithHandler().size());
	}

	@Test
	public void lastRowShouldNotBeDeliveredTwice() throws Exception {
		List<String> got = importWithHandler("a1", "a2");
		assertEquals(2, got.size());
	}

	@Test
	public void listModeShouldStillReturnAllRows() throws Exception {
		List<NameEntity> list = ExcelImportUtil.importExcelBySax(sheetOf("a1", "a2", "a3"), NameEntity.class, params());
		assertEquals(3, list.size());
		assertEquals("a3", list.get(2).getName());
	}

	/** 已有的自定义 ISaxRowRead 实现没有 endSheet 方法，也要能继续使用（接口里是 default 方法） */
	@Test
	public void customRowReadWithoutEndSheetShouldStillWork() throws Exception {
		final List<Integer> rows = new ArrayList<Integer>();
		new SaxReadExcel().readExcel(sheetOf("a1", "a2"), null, params(), new ISaxRowRead() {
			@Override
			public <T> List<T> getList() {
				return null;
			}

			@Override
			public void parse(int index, List<SaxReadCellEntity> datas) {
				rows.add(index);
			}
		}, null);
		// 表头 + 2 行数据
		assertEquals(Arrays.asList(0, 1, 2), rows);
	}
}
