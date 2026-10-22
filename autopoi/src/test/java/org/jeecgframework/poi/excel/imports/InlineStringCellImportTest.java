package org.jeecgframework.poi.excel.imports;

import org.jeecgframework.poi.excel.ExcelImportUtil;
import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

/**
 * [issues/190 Bug1] 内联字符串(t="inlineStr")单元格：DOM 导入时 setCellType(STRING) 会把值清空，
 * 整行读出来全是空被当成空行丢弃，表现为"总共 N 条，实际导入 0 条"。
 * 电子税务局等非 Excel/WPS 生成的 xlsx 常用这种写法（用 WPS 另存后会改成共享字符串，所以能读）。
 */
public class InlineStringCellImportTest {

	public static class Inv {
		@Excel(name = "发票号码")
		private String no;
		@Excel(name = "金额")
		private BigDecimal amount;
		@Excel(name = "开票日期", format = "yyyy-MM-dd HH:mm:ss")
		private Date day;
		public String getNo() { return no; }
		public void setNo(String no) { this.no = no; }
		public BigDecimal getAmount() { return amount; }
		public void setAmount(BigDecimal amount) { this.amount = amount; }
		public Date getDay() { return day; }
		public void setDay(Date day) { this.day = day; }
	}

	private static final String NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

	/** 内联字符串单元格 */
	private static String is(String ref, String v) {
		return "<c r=\"" + ref + "\" t=\"inlineStr\"><is><t>" + v + "</t></is></c>";
	}

	private static String num(String ref, String v) {
		return "<c r=\"" + ref + "\"><v>" + v + "</v></c>";
	}

	private static void put(ZipOutputStream z, String name, String body) throws Exception {
		z.putNextEntry(new ZipEntry(name));
		z.write(body.getBytes("UTF-8"));
		z.closeEntry();
	}

	/** 手工拼一个最小的 xlsx（只有 sheet1，没有共享字符串表），所有字符串都是内联字符串 */
	private static byte[] xlsx(String rowsXml) throws Exception {
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		ZipOutputStream z = new ZipOutputStream(bos);
		String xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
		put(z, "[Content_Types].xml", xml + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
				+ "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
				+ "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
				+ "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
				+ "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>");
		put(z, "_rels/.rels", xml + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
				+ "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
		put(z, "xl/workbook.xml", xml + "<workbook xmlns=\"" + NS + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
				+ "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
		put(z, "xl/_rels/workbook.xml.rels", xml + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
				+ "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
		put(z, "xl/worksheets/sheet1.xml", xml + "<worksheet xmlns=\"" + NS + "\"><sheetData>" + rowsXml + "</sheetData></worksheet>");
		z.close();
		return bos.toByteArray();
	}

	private static final String HEADER = "<row r=\"1\">" + is("A1", "发票号码") + is("B1", "金额") + is("C1", "开票日期") + "</row>";

	private static List<Inv> importDom(String dataRowsXml) throws Exception {
		return ExcelImportUtil.importExcel(new ByteArrayInputStream(xlsx(HEADER + dataRowsXml)), Inv.class, new ImportParams());
	}

	/** 整行都是内联字符串：以前整行被当成空行丢弃，读出 0 条 */
	@Test
	public void allInlineStringRowShouldBeImported() throws Exception {
		List<Inv> list = importDom("<row r=\"2\">" + is("A2", "N001") + is("B2", "12.50") + is("C2", "2026-01-31 19:28:40") + "</row>");
		assertEquals(1, list.size());
		assertEquals("N001", list.get(0).getNo());
		assertEquals(new BigDecimal("12.50"), list.get(0).getAmount());
		assertNotNull(list.get(0).getDay());
		assertEquals("2026-01-31 19:28:40", new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(list.get(0).getDay()));
	}

	/** 内联字符串和数值单元格混合：以前数值能读到，文本字段全是 null */
	@Test
	public void inlineStringShouldNotBeClearedWhenMixedWithNumber() throws Exception {
		List<Inv> list = importDom("<row r=\"2\">" + is("A2", "N002") + num("B2", "7.5") + is("C2", "2026-02-01 00:00:00") + "</row>");
		assertEquals(1, list.size());
		assertEquals("N002", list.get(0).getNo());
		assertEquals(0, new BigDecimal("7.5").compareTo(list.get(0).getAmount()));
	}

	/** 多行都要读到 */
	@Test
	public void allRowsShouldBeImported() throws Exception {
		StringBuilder rows = new StringBuilder();
		for (int i = 2; i <= 6; i++) {
			rows.append("<row r=\"").append(i).append("\">").append(is("A" + i, "N" + i)).append(is("B" + i, i + ".5")).append("</row>");
		}
		List<Inv> list = importDom(rows.toString());
		assertEquals(5, list.size());
		assertEquals("N2", list.get(0).getNo());
		assertEquals("N6", list.get(4).getNo());
	}
}
