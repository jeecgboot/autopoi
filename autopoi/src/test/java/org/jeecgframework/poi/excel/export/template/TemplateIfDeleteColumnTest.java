package org.jeecgframework.poi.excel.export.template;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;
import org.junit.Test;

import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;

/**
 * 模板导出的 {{!if:(条件)}} 指令：条件为真时删除当前列，为假时清空指令单元格。
 * XSSF 下删除列后再操作已被删除的单元格会抛 XmlValueDisconnectedException（异常被吞，导出返回 null）。
 *
 * @since 2026-10-20 issues/174 使用表达式 !if: 删除当前列发生异常 XmlValueDisconnectedException
 */
public class TemplateIfDeleteColumnTest {

	/** rows 里的元素都按字符串写入，null 表示不创建单元格 */
	private static Workbook template(boolean xssf, String[]... rows) {
		Workbook wb = xssf ? new XSSFWorkbook() : new HSSFWorkbook();
		Sheet sheet = wb.createSheet("t");
		for (int r = 0; r < rows.length; r++) {
			Row row = sheet.createRow(r);
			for (int c = 0; c < rows[r].length; c++) {
				if (rows[r][c] != null) {
					row.createCell(c).setCellValue(rows[r][c]);
				}
			}
		}
		return wb;
	}

	private static Workbook export(Workbook tpl, Map<String, Object> map) {
		TemplateExportParams params = new TemplateExportParams();
		params.setTemplateWb(tpl);
		return ExcelExportUtil.exportExcel(params, map);
	}

	private static Map<String, Object> map(Object... kv) {
		Map<String, Object> m = new HashMap<String, Object>();
		for (int i = 0; i < kv.length; i += 2) {
			m.put((String) kv[i], kv[i + 1]);
		}
		return m;
	}

	private static String text(Workbook wb, int row, int col) {
		Cell cell = wb.getSheetAt(0).getRow(row).getCell(col);
		return cell == null ? null : cell.getStringCellValue();
	}

	private void deleteMiddleColumn(boolean xssf) {
		Workbook wb = export(
				template(xssf, new String[]{"A", "{{!if:(hide)}}", "C"}, new String[]{"a", "b", "c"}),
				map("hide", true));
		assertNotNull("!if 删除列时导出不应失败（返回 null）", wb);
		// 第 2 列被删除，C 左移到第 2 列
		assertEquals("A", text(wb, 0, 0));
		assertEquals("C", text(wb, 0, 1));
		assertNull("被删除后最后一列应为空", text(wb, 0, 2));
		assertEquals("a", text(wb, 1, 0));
		assertEquals("c", text(wb, 1, 1));
		assertNull(text(wb, 1, 2));
	}

	@Test
	public void deleteMiddleColumnInXlsx() {
		deleteMiddleColumn(true);
	}

	@Test
	public void deleteMiddleColumnInXls() {
		deleteMiddleColumn(false);
	}

	private void deleteFirstColumn(boolean xssf) {
		Workbook wb = export(
				template(xssf, new String[]{"{{!if:(hide)}}", "keep", "{{x}}"}),
				map("hide", true, "x", "X"));
		assertNotNull(wb);
		assertEquals("keep", text(wb, 0, 0));
		assertEquals("X", text(wb, 0, 1));
	}

	@Test
	public void deleteFirstColumnInXlsx() {
		deleteFirstColumn(true);
	}

	@Test
	public void deleteFirstColumnInXls() {
		deleteFirstColumn(false);
	}

	private void keepColumnWhenConditionFalse(boolean xssf) {
		Workbook wb = export(
				template(xssf, new String[]{"A", "{{!if:(hide)}}", "C"}, new String[]{"a", "b", "c"}),
				map("hide", false));
		assertNotNull(wb);
		assertEquals("A", text(wb, 0, 0));
		assertEquals("指令单元格应被清空", "", text(wb, 0, 1));
		assertEquals("C", text(wb, 0, 2));
		assertEquals("b", text(wb, 1, 1));
	}

	@Test
	public void keepColumnWhenConditionFalseInXlsx() {
		keepColumnWhenConditionFalse(true);
	}

	@Test
	public void keepColumnWhenConditionFalseInXls() {
		keepColumnWhenConditionFalse(false);
	}

	/** 相邻两列都要删除：删除第一列后，第二列会左移到当前位置，也要被处理 */
	private void deleteAdjacentColumns(boolean xssf) {
		Workbook wb = export(
				template(xssf, new String[]{"{{!if:(a)}}", "{{!if:(b)}}", "keep"}, new String[]{"1", "2", "3"}),
				map("a", true, "b", true));
		assertNotNull(wb);
		assertEquals("keep", text(wb, 0, 0));
		assertNull(text(wb, 0, 1));
		assertEquals("3", text(wb, 1, 0));
	}

	@Test
	public void deleteAdjacentColumnsInXlsx() {
		deleteAdjacentColumns(true);
	}

	@Test
	public void deleteAdjacentColumnsInXls() {
		deleteAdjacentColumns(false);
	}

	/** 一列删除、一列保留：互不影响 */
	@Test
	public void deleteOneAndKeepAnotherInXlsx() {
		Workbook wb = export(
				template(true, new String[]{"{{!if:(a)}}", "{{!if:(b)}}", "keep"}, new String[]{"1", "2", "3"}),
				map("a", true, "b", false));
		assertNotNull(wb);
		assertEquals("", text(wb, 0, 0));
		assertEquals("keep", text(wb, 0, 1));
		assertEquals("2", text(wb, 1, 0));
		assertEquals("3", text(wb, 1, 1));
	}

	@Test
	public void templateWithoutIfShouldBeUnaffected() {
		Workbook wb = export(template(true, new String[]{"{{x}}", "B"}), map("x", "X"));
		assertNotNull(wb);
		assertEquals("X", text(wb, 0, 0));
		assertEquals("B", text(wb, 0, 1));
	}
}
