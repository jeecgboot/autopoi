package org.jeecgframework.poi.excel.imports;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.ExcelImportUtil;
import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.annotation.ExcelCollection;
import org.jeecgframework.poi.excel.annotation.ExcelTarget;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.jeecgframework.poi.exception.excel.ExcelImportException;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.fail;

/**
 * 实体字段类型是枚举时的导入导出。
 *
 * @since 2026-10-20 issues/18 导入导出时没有考虑字段类型是枚举的情况
 */
public class ImportEnumFieldTest {

	public enum Status {
		ENABLE, DISABLE
	}

	public static class PlainEnumEntity {
		@Excel(name = "名称", orderNum = "1")
		private String name;
		@Excel(name = "状态", orderNum = "2")
		private Status status;

		public PlainEnumEntity() {
		}

		public PlainEnumEntity(String name, Status status) {
			this.name = name;
			this.status = status;
		}

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}

		public Status getStatus() {
			return status;
		}

		public void setStatus(Status status) {
			this.status = status;
		}
	}

	public static class DictEnumEntity {
		@Excel(name = "状态", replace = {"启用_ENABLE", "禁用_DISABLE"})
		private Status status;

		public Status getStatus() {
			return status;
		}

		public void setStatus(Status status) {
			this.status = status;
		}
	}

	@ExcelTarget("enumParent")
	public static class Parent {
		@Excel(name = "名称", needMerge = true)
		private String name;
		@ExcelCollection(name = "明细")
		private List<Child> children = new ArrayList<Child>();

		public String getName() {
			return name;
		}

		public void setName(String name) {
			this.name = name;
		}

		public List<Child> getChildren() {
			return children;
		}

		public void setChildren(List<Child> children) {
			this.children = children;
		}
	}

	public static class Child {
		@Excel(name = "子状态")
		private Status status;

		public Status getStatus() {
			return status;
		}

		public void setStatus(Status status) {
			this.status = status;
		}
	}

	private static InputStream sheetOf(String header, String... values) throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Sheet1");
		sheet.createRow(0).createCell(0).setCellValue(header);
		for (int i = 0; i < values.length; i++) {
			Row row = sheet.createRow(i + 1);
			if (values[i] != null) {
				row.createCell(0).setCellValue(values[i]);
			}
		}
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		wb.write(bos);
		wb.close();
		return new ByteArrayInputStream(bos.toByteArray());
	}

	@Test
	public void enumNameShouldBeImported() throws Exception {
		List<PlainEnumEntity> list = ExcelImportUtil.importExcel(sheetOf("状态", "ENABLE", "DISABLE"), PlainEnumEntity.class, new ImportParams());
		assertEquals(2, list.size());
		assertEquals(Status.ENABLE, list.get(0).getStatus());
		assertEquals(Status.DISABLE, list.get(1).getStatus());
	}

	@Test
	public void enumNameWithBlanksShouldBeImported() throws Exception {
		List<PlainEnumEntity> list = ExcelImportUtil.importExcel(sheetOf("状态", " ENABLE "), PlainEnumEntity.class, new ImportParams());
		assertEquals(Status.ENABLE, list.get(0).getStatus());
	}

	@Test
	public void emptyEnumCellShouldBeNull() throws Exception {
		// 名称列有值、状态列为空的行不是空行，状态应为 null
		XSSFWorkbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Sheet1");
		Row head = sheet.createRow(0);
		head.createCell(0).setCellValue("名称");
		head.createCell(1).setCellValue("状态");
		String[] names = {"a", "b", "c"};
		String[] statuses = {"ENABLE", null, "DISABLE"};
		for (int i = 0; i < names.length; i++) {
			Row row = sheet.createRow(i + 1);
			row.createCell(0).setCellValue(names[i]);
			if (statuses[i] != null) {
				row.createCell(1).setCellValue(statuses[i]);
			}
		}
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		wb.write(bos);
		wb.close();
		List<PlainEnumEntity> list = ExcelImportUtil.importExcel(new ByteArrayInputStream(bos.toByteArray()), PlainEnumEntity.class, new ImportParams());
		assertEquals(3, list.size());
		assertEquals(Status.ENABLE, list.get(0).getStatus());
		assertNull(list.get(1).getStatus());
		assertEquals(Status.DISABLE, list.get(2).getStatus());
	}

	@Test
	public void enumShouldWorkWithDictReplace() throws Exception {
		List<DictEnumEntity> list = ExcelImportUtil.importExcel(sheetOf("状态", "启用", "禁用"), DictEnumEntity.class, new ImportParams());
		assertEquals(Status.ENABLE, list.get(0).getStatus());
		assertEquals(Status.DISABLE, list.get(1).getStatus());
	}

	@Test
	public void unknownEnumValueShouldFailWithImportException() throws Exception {
		try {
			ExcelImportUtil.importExcel(sheetOf("状态", "NOT_EXIST"), PlainEnumEntity.class, new ImportParams());
			fail("不存在的枚举值应抛出 ExcelImportException");
		} catch (ExcelImportException expected) {
			// ok
		}
	}

	@Test
	public void enumShouldBeExportedAsName() {
		List<PlainEnumEntity> data = new ArrayList<PlainEnumEntity>(Arrays.asList(
				new PlainEnumEntity("a", Status.ENABLE), new PlainEnumEntity("b", Status.DISABLE)));
		Workbook wb = ExcelExportUtil.exportExcel(new ExportParams(), PlainEnumEntity.class, data);
		Sheet sheet = wb.getSheetAt(0);
		assertEquals("ENABLE", sheet.getRow(1).getCell(1).getStringCellValue());
		assertEquals("DISABLE", sheet.getRow(2).getCell(1).getStringCellValue());
	}

	@Test
	public void exportedFileShouldBeImportedBack() throws Exception {
		List<PlainEnumEntity> data = new ArrayList<PlainEnumEntity>(Arrays.asList(
				new PlainEnumEntity("a", Status.ENABLE), new PlainEnumEntity("b", Status.DISABLE)));
		Workbook wb = ExcelExportUtil.exportExcel(new ExportParams(), PlainEnumEntity.class, data);
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		wb.write(bos);
		List<PlainEnumEntity> list = ExcelImportUtil.importExcel(new ByteArrayInputStream(bos.toByteArray()), PlainEnumEntity.class, new ImportParams());
		assertEquals(2, list.size());
		assertEquals("a", list.get(0).getName());
		assertEquals(Status.ENABLE, list.get(0).getStatus());
		assertEquals(Status.DISABLE, list.get(1).getStatus());
	}

	@Test
	public void enumInsideCollectionShouldBeImported() throws Exception {
		Child c1 = new Child();
		c1.setStatus(Status.ENABLE);
		Child c2 = new Child();
		c2.setStatus(Status.DISABLE);
		Parent p = new Parent();
		p.setName("P");
		p.setChildren(new ArrayList<Child>(Arrays.asList(c1, c2)));
		Workbook wb = ExcelExportUtil.exportExcel(new ExportParams("t", "s"), Parent.class, new ArrayList<Parent>(Arrays.asList(p)));
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		wb.write(bos);
		ImportParams params = new ImportParams();
		params.setTitleRows(1);
		params.setHeadRows(2);
		List<Parent> list = ExcelImportUtil.importExcel(new ByteArrayInputStream(bos.toByteArray()), Parent.class, params);
		assertEquals(1, list.size());
		assertEquals(2, list.get(0).getChildren().size());
		assertEquals(Status.ENABLE, list.get(0).getChildren().get(0).getStatus());
		assertEquals(Status.DISABLE, list.get(0).getChildren().get(1).getStatus());
	}

	@Test
	public void enumShouldBeImportedBySax() throws Exception {
		ImportParams params = new ImportParams();
		params.setSheetNum(1);
		List<PlainEnumEntity> list = ExcelImportUtil.importExcelBySax(sheetOf("状态", "ENABLE", "DISABLE"), PlainEnumEntity.class, params);
		assertEquals(2, list.size());
		assertEquals(Status.ENABLE, list.get(0).getStatus());
		assertEquals(Status.DISABLE, list.get(1).getStatus());
	}
}
