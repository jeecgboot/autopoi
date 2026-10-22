package org.jeecgframework.poi.excel.export.template;

import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;
import org.junit.Test;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;

/**
 * [issues/54] 模板在 classpath 里（打成 jar 后在 BOOT-INF/classes 或 jar 里）时，
 * 导出用 FileInputStream 读，只认文件系统路径，jar 里的模板读不到，导出返回 null。
 * 这里用 classpath 上的测试模板模拟：它在 target/test-classes 里，不在当前工作目录下，跟 jar 里的模板一样不能用文件路径打开。
 */
public class TemplateClasspathLoadTest {

	private static final String CLASSPATH_TEMPLATE = "templates/issues9741/template_image.xlsx";

	private static Workbook export(String templateUrl) {
		Map<String, Object> map = new HashMap<String, Object>();
		map.put("name", "张三");
		return ExcelExportUtil.exportExcel(new TemplateExportParams(templateUrl), map);
	}

	/** 前提：这个相对路径在当前工作目录下确实不存在，只能从 classpath 找 */
	@Test
	public void templateIsNotReachableAsFilePath() {
		assertEquals(false, new File(CLASSPATH_TEMPLATE).exists());
	}

	@Test
	public void classpathRelativePathShouldLoad() throws Exception {
		Workbook wb = export(CLASSPATH_TEMPLATE);
		assertNotNull("classpath 下的模板应该能读到", wb);
		wb.close();
	}

	@Test
	public void classpathPrefixShouldLoad() throws Exception {
		Workbook wb = export("classpath:" + CLASSPATH_TEMPLATE);
		assertNotNull("classpath: 前缀的模板应该能读到", wb);
		wb.close();
	}

	@Test
	public void leadingSlashShouldLoad() throws Exception {
		Workbook wb = export("/" + CLASSPATH_TEMPLATE);
		assertNotNull("以 / 开头的 classpath 路径应该能读到", wb);
		wb.close();
	}

	/** 原来的用法不能变：文件系统绝对路径照样能读 */
	@Test
	public void absoluteFilePathShouldStillLoad() throws Exception {
		File file = new File("src/test/resources/" + CLASSPATH_TEMPLATE);
		Workbook wb = export(file.getAbsolutePath());
		assertNotNull("文件系统路径应该仍然能读到", wb);
		wb.close();
	}

	/** 文件系统和 classpath 都找不到：仍然返回 null，不抛异常（保持原行为） */
	@Test
	public void missingTemplateShouldStillReturnNull() {
		assertNull(export("templates/not-exists.xlsx"));
	}
}
