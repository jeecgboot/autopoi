package org.jeecgframework.poi.word.parse;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

/**
 * Word 导出时写入文本的工具类。
 * <p>
 * Word 不会把 &lt;w:t&gt; 里的换行符渲染成换行，文本里的换行必须用 &lt;w:br/&gt; 表示，
 * 所以带有换行符(\n、\r\n、\r)的文本要按行拆开，行与行之间用 addBreak() 换行。
 *
 * @since 2026-10-20 issues/165 Word 模板填充时，文本中存在换行符，无法自动换行
 */
public final class WordTextUtil {

	private WordTextUtil() {
	}

	/**
	 * 设置 run 的文本，文本里的换行符转成 Word 的换行
	 *
	 * @param run  目标 run，原来在第 0 个位置的文本会被替换
	 * @param text 文本，null 按空字符串处理
	 */
	public static void setText(XWPFRun run, String text) {
		String[] lines = (text == null ? "" : text).split("\r\n|\r|\n", -1);
		run.setText(lines[0], 0);
		for (int i = 1; i < lines.length; i++) {
			run.addBreak();
			run.setText(lines[i], run.getCTR().sizeOfTArray());
		}
	}

	/**
	 * 设置单元格的文本（新建一个 run），文本里的换行符转成 Word 的换行。
	 * 与 XWPFTableCell.setText 的做法一致：写入单元格的第一个段落。
	 */
	public static void setText(XWPFTableCell cell, String text) {
		CTP ctP = cell.getCTTc().sizeOfPArray() == 0 ? cell.getCTTc().addNewP() : cell.getCTTc().getPArray(0);
		XWPFParagraph paragraph = new XWPFParagraph(ctP, cell);
		setText(paragraph.createRun(), text);
	}
}
