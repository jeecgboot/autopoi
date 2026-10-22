package org.jeecgframework.poi.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.word.entity.params.ExcelListEntity;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;

/**
 * [issues/108] Word 模板表格用 ExcelListEntity 循环输出时，占位符所在的行 {{$fe:entity}} 没有被删除，
 * 数据行追加在它后面（List 模式会先删除占位符行，两种模式不一致）。
 */
public class WordTableListEntityTest {

	public static class Person {
		@Excel(name = "名称")
		private String name;
		@Excel(name = "年龄")
		private String age;
		public Person() {
		}
		public Person(String name, String age) {
			this.name = name;
			this.age = age;
		}
		public String getName() { return name; }
		public void setName(String name) { this.name = name; }
		public String getAge() { return age; }
		public void setAge(String age) { this.age = age; }
	}

	/** 表头行 + 占位符行，占位符在第一列 */
	private static File template() throws Exception {
		XWPFDocument doc = new XWPFDocument();
		XWPFTable table = doc.createTable(2, 2);
		table.getRow(0).getCell(0).setText("名称");
		table.getRow(0).getCell(1).setText("年龄");
		table.getRow(1).getCell(0).setText("{{$fe:entity}}");
		File file = File.createTempFile("word108", ".docx");
		FileOutputStream out = new FileOutputStream(file);
		try {
			doc.write(out);
		} finally {
			out.close();
		}
		return file;
	}

	private static List<String> rowTexts(XWPFTable table, int row) {
		List<String> texts = new ArrayList<String>();
		for (XWPFTableCell cell : table.getRow(row).getTableCells()) {
			texts.add(cell.getText());
		}
		return texts;
	}

	@Test
	public void markerRowShouldBeRemovedAndDataRowsFollowHeader() throws Exception {
		File file = template();
		try {
			Map<String, Object> map = new HashMap<String, Object>();
			map.put("entity", new ExcelListEntity(new ArrayList<Person>(Arrays.asList(new Person("甲", "1"), new Person("乙", "2"))), Person.class));
			XWPFDocument result = WordExportUtil.exportWord07(file.getAbsolutePath(), map);
			XWPFTable table = result.getTables().get(0);

			assertEquals("表头 + 2 行数据，占位符行应该被删除", 3, table.getNumberOfRows());
			assertEquals(Arrays.asList("名称", "年龄"), rowTexts(table, 0));
			assertEquals(Arrays.asList("甲", "1"), rowTexts(table, 1));
			assertEquals(Arrays.asList("乙", "2"), rowTexts(table, 2));
			for (int i = 0; i < table.getNumberOfRows(); i++) {
				for (String text : rowTexts(table, i)) {
					assertFalse("不应残留占位符: " + text, text.contains("$fe"));
				}
			}
		} finally {
			file.delete();
		}
	}

	/** 列表为空：占位符行也要删掉，只剩表头 */
	@Test(timeout = 10000)
	public void markerRowShouldBeRemovedWhenListIsEmpty() throws Exception {
		File file = template();
		try {
			Map<String, Object> map = new HashMap<String, Object>();
			map.put("entity", new ExcelListEntity(new ArrayList<Person>(), Person.class));
			XWPFDocument result = WordExportUtil.exportWord07(file.getAbsolutePath(), map);
			XWPFTable table = result.getTables().get(0);
			assertEquals(1, table.getNumberOfRows());
			assertEquals(Arrays.asList("名称", "年龄"), rowTexts(table, 0));
		} finally {
			file.delete();
		}
	}
}
