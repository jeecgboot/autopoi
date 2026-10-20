package org.jeecgframework.poi.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.word.entity.MyXWPFDocument;
import org.jeecgframework.poi.word.entity.params.ExcelListEntity;
import org.jeecgframework.poi.word.parse.WordTextUtil;
import org.jeecgframework.poi.word.parse.excel.ExcelEntityParse;
import org.jeecgframework.poi.word.parse.excel.ExcelMapParse;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.assertEquals;

/**
 * Word 模板填充时，值里的换行符(\n、\r\n、\r)必须转成 Word 的换行(&lt;w:br/&gt;)，
 * 不能原样放进 &lt;w:t&gt;（Word 会把它当空白，不换行）。
 *
 * @since 2026-10-20 issues/165 Word模板填充时，如文本中存在换行符，无法自动换行
 */
public class WordNewlineTest {

	public static class Remark {
		@Excel(name = "备注")
		private String remark;

		public Remark(String remark) {
			this.remark = remark;
		}

		public String getRemark() {
			return remark;
		}

		public void setRemark(String remark) {
			this.remark = remark;
		}
	}

	/** 每个字符串一个段落（run 里写原文），重新读取成 MyXWPFDocument */
	private static MyXWPFDocument docOf(String... paragraphs) throws Exception {
		XWPFDocument doc = new XWPFDocument();
		for (String p : paragraphs) {
			doc.createParagraph().createRun().setText(p);
		}
		return reload(doc);
	}

	private static MyXWPFDocument reload(XWPFDocument doc) throws Exception {
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		doc.write(bos);
		return new MyXWPFDocument(new ByteArrayInputStream(bos.toByteArray()));
	}

	private static int breaks(XWPFParagraph paragraph) {
		int n = 0;
		for (XWPFRun run : paragraph.getRuns()) {
			n += run.getCTR().sizeOfBrArray();
		}
		return n;
	}

	private static int breaks(XWPFTableCell cell) {
		int n = 0;
		for (XWPFParagraph p : cell.getParagraphs()) {
			n += breaks(p);
		}
		return n;
	}

	private static Map<String, Object> map(String content) {
		Map<String, Object> map = new HashMap<String, Object>();
		map.put("content", content);
		return map;
	}

	// ------------------------------------------------------------- 段落占位符

	@Test
	public void newlinesInParagraphShouldBecomeBreaks() throws Exception {
		MyXWPFDocument doc = docOf("{{content}}");
		WordExportUtil.exportWord07(doc, map("第一行\n第二行\r\n第三行\r第四行"));
		XWPFParagraph p = doc.getParagraphs().get(0);
		assertEquals("3 个换行符(\\n、\\r\\n、\\r)应得到 3 个 w:br", 3, breaks(p));
		assertEquals("第一行\n第二行\n第三行\n第四行", p.getText());
	}

	@Test
	public void textWithoutNewlineShouldBeUnchanged() throws Exception {
		MyXWPFDocument doc = docOf("{{content}}");
		WordExportUtil.exportWord07(doc, map("普通文本"));
		XWPFParagraph p = doc.getParagraphs().get(0);
		assertEquals(0, breaks(p));
		assertEquals("普通文本", p.getText());
	}

	@Test
	public void newlineWithSurroundingTextShouldWork() throws Exception {
		MyXWPFDocument doc = docOf("前缀{{content}}后缀");
		WordExportUtil.exportWord07(doc, map("甲\n乙"));
		XWPFParagraph p = doc.getParagraphs().get(0);
		assertEquals(1, breaks(p));
		assertEquals("前缀甲\n乙后缀", p.getText());
	}

	@Test
	public void consecutiveAndTrailingNewlinesShouldBeKept() throws Exception {
		MyXWPFDocument doc = docOf("{{content}}", "{{content}}");
		WordExportUtil.exportWord07(doc, map("a\n\nb"));
		assertEquals("连续两个换行 = 2 个 w:br", 2, breaks(doc.getParagraphs().get(0)));
		assertEquals("a\n\nb", doc.getParagraphs().get(0).getText());

		MyXWPFDocument doc2 = docOf("{{content}}");
		WordExportUtil.exportWord07(doc2, map("a\n"));
		assertEquals("末尾的换行也保留", 1, breaks(doc2.getParagraphs().get(0)));
	}

	@Test
	public void newlineInTableCellPlaceholderShouldBecomeBreak() throws Exception {
		XWPFDocument doc = new XWPFDocument();
		XWPFTable table = doc.createTable(1, 1);
		table.getRow(0).getCell(0).setText("{{content}}");
		MyXWPFDocument reloaded = reload(doc);
		WordExportUtil.exportWord07(reloaded, map("甲\n乙"));
		XWPFTableCell cell = reloaded.getTables().get(0).getRow(0).getCell(0);
		assertEquals(1, breaks(cell));
		assertEquals("甲\n乙", cell.getText());
	}

	// ------------------------------------------------------------- 表格循环行

	@Test
	public void newlineInMapListRowShouldBecomeBreak() throws Exception {
		XWPFDocument doc = new XWPFDocument();
		XWPFTable table = doc.createTable(1, 2);
		table.getRow(0).getCell(0).setText("name");
		table.getRow(0).getCell(1).setText("remark");
		List<Object> list = new ArrayList<Object>();
		Map<String, Object> row = new HashMap<String, Object>();
		row.put("name", "张三");
		row.put("remark", "第一行\n第二行");
		list.add(row);
		ExcelMapParse.parseNextRowAndAddRow(table, 0, list);
		// 写出再读回，和真实导出的文件一致
		XWPFTable written = reload(doc).getTables().get(0);
		XWPFTableCell remark = written.getRow(0).getCell(1);
		assertEquals(1, breaks(remark));
		assertEquals("第一行\n第二行", remark.getText());
		assertEquals("张三", written.getRow(0).getCell(0).getText());
	}

	@Test
	public void newlineInExcelListEntityRowShouldBecomeBreak() throws Exception {
		XWPFDocument doc = new XWPFDocument();
		XWPFTable table = doc.createTable(1, 1);
		table.getRow(0).getCell(0).setText("备注");
		List<Remark> list = new ArrayList<Remark>();
		list.add(new Remark("第一行\n第二行"));
		new ExcelEntityParse().parseNextRowAndAddRow(table, 1, new ExcelListEntity(list, Remark.class));
		// 写出再读回，和真实导出的文件一致
		XWPFTableCell cell = reload(doc).getTables().get(0).getRow(1).getCell(0);
		assertEquals(1, breaks(cell));
		assertEquals("第一行\n第二行", cell.getText());
	}

	// ------------------------------------------------------------- 工具类

	@Test
	public void setTextShouldReplaceExistingTextOfRun() throws Exception {
		XWPFDocument doc = new XWPFDocument();
		XWPFRun run = doc.createParagraph().createRun();
		run.setText("旧文本");
		WordTextUtil.setText(run, "新\n文本");
		assertEquals("新\n文本", run.getParagraph().getText());
		assertEquals(1, run.getCTR().sizeOfBrArray());
	}

	@Test
	public void setTextWithNullShouldBeEmpty() throws Exception {
		XWPFDocument doc = new XWPFDocument();
		XWPFRun run = doc.createParagraph().createRun();
		WordTextUtil.setText(run, null);
		assertEquals("", run.getParagraph().getText());
		assertEquals(0, run.getCTR().sizeOfBrArray());
	}
}
