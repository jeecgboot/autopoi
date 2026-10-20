package org.jeecgframework.poi.excel.imports;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.excel.ExcelImportUtil;
import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.annotation.ExcelCollection;
import org.jeecgframework.poi.excel.annotation.ExcelTarget;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

/**
 * SAX 导入时，空单元格在 XML 中不会生成 &lt;c&gt; 元素，列号必须按单元格坐标（r 属性，如 "C5"）确定，
 * 不能按出现顺序累加，否则空单元格后面的列会整体前移，值被赋给错误的字段。
 *
 * @since 2026-10-20 issues/190 SAX 空单元格导致列错位
 */
public class SaxImportColumnIndexTest {

	public static class Cols {
		@Excel(name = "A")
		private String a;
		@Excel(name = "B")
		private String b;
		@Excel(name = "C")
		private String c;
		@Excel(name = "D")
		private String d;
		@Excel(name = "E")
		private String e;

		public String getA() { return a; }
		public void setA(String a) { this.a = a; }
		public String getB() { return b; }
		public void setB(String b) { this.b = b; }
		public String getC() { return c; }
		public void setC(String c) { this.c = c; }
		public String getD() { return d; }
		public void setD(String d) { this.d = d; }
		public String getE() { return e; }
		public void setE(String e) { this.e = e; }
	}

	public static class IntCols {
		@Excel(name = "名称")
		private String name;
		@Excel(name = "数量")
		private Integer qty;

		public String getName() { return name; }
		public void setName(String name) { this.name = name; }
		public Integer getQty() { return qty; }
		public void setQty(Integer qty) { this.qty = qty; }
	}

	/** 第 27、28 列（AA、AB），列标是两个字母 */
	public static class WideCols {
		@Excel(name = "头1")
		private String first;
		@Excel(name = "头AA")
		private String aa;
		@Excel(name = "头AB")
		private String ab;

		public String getFirst() { return first; }
		public void setFirst(String first) { this.first = first; }
		public String getAa() { return aa; }
		public void setAa(String aa) { this.aa = aa; }
		public String getAb() { return ab; }
		public void setAb(String ab) { this.ab = ab; }
	}

	@ExcelTarget("saxTeacher")
	public static class Teacher {
		@Excel(name = "老师")
		private String name;
		@ExcelCollection(name = "学生")
		private List<Student> students = new ArrayList<Student>();

		public String getName() { return name; }
		public void setName(String name) { this.name = name; }
		public List<Student> getStudents() { return students; }
		public void setStudents(List<Student> students) { this.students = students; }
	}

	public static class Student {
		@Excel(name = "学生名")
		private String name;

		public String getName() { return name; }
		public void setName(String name) { this.name = name; }
	}

	/** rows 里的元素：null 表示不创建该单元格（和 Excel 保存空单元格的方式一致），数字和字符串分别创建对应类型 */
	private static InputStream sheetOf(Object[]... rows) throws Exception {
		XSSFWorkbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet("Sheet1");
		for (int r = 0; r < rows.length; r++) {
			if (rows[r] == null) {
				continue;
			}
			Row row = sheet.createRow(r);
			for (int c = 0; c < rows[r].length; c++) {
				Object v = rows[r][c];
				if (v == null) {
					continue;
				}
				if (v instanceof Number) {
					row.createCell(c).setCellValue(((Number) v).doubleValue());
				} else {
					row.createCell(c).setCellValue(String.valueOf(v));
				}
			}
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

	private static final Object[] HEAD = {"A", "B", "C", "D", "E"};

	private static List<Cols> importCols(Object[]... dataRows) throws Exception {
		Object[][] rows = new Object[dataRows.length + 1][];
		rows[0] = HEAD;
		System.arraycopy(dataRows, 0, rows, 1, dataRows.length);
		return ExcelImportUtil.importExcelBySax(sheetOf(rows), Cols.class, params());
	}

	@Test
	public void rowWithoutEmptyCellsShouldBeUnchanged() throws Exception {
		List<Cols> list = importCols(new Object[]{"a", "b", "c", "d", "e"});
		assertEquals(1, list.size());
		assertEquals("a", list.get(0).getA());
		assertEquals("b", list.get(0).getB());
		assertEquals("c", list.get(0).getC());
		assertEquals("d", list.get(0).getD());
		assertEquals("e", list.get(0).getE());
	}

	@Test
	public void emptyCellInTheMiddleShouldNotShiftFollowingColumns() throws Exception {
		List<Cols> list = importCols(new Object[]{"a", null, "c", "d", "e"});
		assertEquals("a", list.get(0).getA());
		assertNull(list.get(0).getB());
		assertEquals("c", list.get(0).getC());
		assertEquals("d", list.get(0).getD());
		assertEquals("e", list.get(0).getE());
	}

	@Test
	public void emptyFirstColumnShouldNotShiftOthers() throws Exception {
		List<Cols> list = importCols(new Object[]{null, "b", "c", "d", "e"});
		assertNull(list.get(0).getA());
		assertEquals("b", list.get(0).getB());
		assertEquals("e", list.get(0).getE());
	}

	@Test
	public void trailingEmptyColumnsShouldBeNull() throws Exception {
		List<Cols> list = importCols(new Object[]{"a", "b", null, null, null});
		assertEquals("a", list.get(0).getA());
		assertEquals("b", list.get(0).getB());
		assertNull(list.get(0).getC());
		assertNull(list.get(0).getD());
		assertNull(list.get(0).getE());
	}

	@Test
	public void severalConsecutiveEmptyCellsShouldBeHandled() throws Exception {
		List<Cols> list = importCols(new Object[]{"a", null, null, null, "e"});
		assertEquals("a", list.get(0).getA());
		assertNull(list.get(0).getB());
		assertNull(list.get(0).getC());
		assertNull(list.get(0).getD());
		assertEquals("e", list.get(0).getE());
	}

	@Test
	public void eachRowShouldBePositionedIndependently() throws Exception {
		List<Cols> list = importCols(
				new Object[]{"a1", "b1", "c1", "d1", "e1"},
				new Object[]{"a2", null, "c2", null, "e2"},
				new Object[]{null, "b3", null, "d3", null});
		assertEquals(3, list.size());
		assertEquals("c1", list.get(0).getC());
		assertEquals("c2", list.get(1).getC());
		assertNull(list.get(1).getD());
		assertEquals("e2", list.get(1).getE());
		assertNull(list.get(2).getA());
		assertEquals("b3", list.get(2).getB());
		assertEquals("d3", list.get(2).getD());
	}

	@Test
	public void completelyEmptyRowShouldBeSkipped() throws Exception {
		Object[][] rows = {HEAD, {"a1", "b1", "c1", "d1", "e1"}, null, {"a3", "b3", "c3", "d3", "e3"}};
		List<Cols> list = ExcelImportUtil.importExcelBySax(sheetOf(rows), Cols.class, params());
		assertEquals(2, list.size());
		assertEquals("e1", list.get(0).getE());
		assertEquals("a3", list.get(1).getA());
		assertEquals("e3", list.get(1).getE());
	}

	@Test
	public void emptyHeaderCellShouldNotShiftHeaderNames() throws Exception {
		// 表头 A 和 C 之间有一个空表头，数据列应与表头列对齐
		Object[][] rows = {{"A", null, "C"}, {"a", "ignored", "c"}};
		List<Cols> list = ExcelImportUtil.importExcelBySax(sheetOf(rows), Cols.class, params());
		assertEquals("a", list.get(0).getA());
		assertEquals("c", list.get(0).getC());
		assertNull(list.get(0).getB());
	}

	@Test
	public void numberAfterEmptyCellShouldMapToItsOwnColumn() throws Exception {
		Object[][] rows = {{"名称", "数量"}, {null, 5}};
		List<IntCols> list = ExcelImportUtil.importExcelBySax(sheetOf(rows), IntCols.class, params());
		assertEquals(1, list.size());
		assertNull(list.get(0).getName());
		assertEquals(Integer.valueOf(5), list.get(0).getQty());
	}

	@Test
	public void twoLetterColumnsShouldBePositionedCorrectly() throws Exception {
		Object[] head = new Object[28];
		head[0] = "头1";
		head[26] = "头AA";
		head[27] = "头AB";
		Object[] data = new Object[28];
		data[0] = "v1";
		data[27] = "vAB";
		List<WideCols> list = ExcelImportUtil.importExcelBySax(sheetOf(new Object[][]{head, data}), WideCols.class, params());
		assertEquals(1, list.size());
		assertEquals("v1", list.get(0).getFirst());
		assertNull(list.get(0).getAa());
		assertEquals("vAB", list.get(0).getAb());
	}

	/** 一对多：子行的主键列(老师)是空单元格(不生成 <c>)，应继续作为上一个老师的集合元素 */
	@Test
	public void emptyKeyCellShouldContinueTheCollection() throws Exception {
		Object[][] rows = {{"老师", "学生_学生名"}, {"T1", "S1"}, {null, "S2"}, {"T2", "S3"}};
		List<Teacher> list = ExcelImportUtil.importExcelBySax(sheetOf(rows), Teacher.class, params());
		assertEquals(2, list.size());
		assertEquals("T1", list.get(0).getName());
		assertEquals(2, list.get(0).getStudents().size());
		assertEquals("S2", list.get(0).getStudents().get(1).getName());
		assertEquals(1, list.get(1).getStudents().size());
	}
}
