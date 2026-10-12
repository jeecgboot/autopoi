package org.jeecgframework.poi.excel.export.template;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.consts.ImageScaleMode;
import org.jeecgframework.poi.entity.ImageEntity;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;
import org.junit.rules.TestName;

import javax.imageio.ImageIO;
import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

/**
 * 模板导出单值占位符图片测试。
 * <p>
 * 模板：src/test/resources/templates/issues9741/template_image.xlsx(.xls)
 * 导出结果另存到 target/test-output/issues9741/ 便于人工打开核对。
 *
 * @author scott
 * @since 2026-10-12 issues/9741 excel模板导出图片失败
 */
public class ExcelTemplateImageExportTest {

	private static final String TEMPLATE_DIR = "/templates/issues9741/";
	private static final String IMAGE_PATH = "/templates/dakytot.jpeg";
	private static final File OUTPUT_DIR = new File("target/test-output/issues9741");

	/** 模板中各占位符所在行（0 基），列均为 B 列 */
	private static final int ROW_NAME = 0;
	private static final int ROW_SIGN = 1;
	private static final int ROW_FIT = 2;
	private static final int ROW_ORIGINAL = 3;
	private static final int ROW_BYTES = 4;
	private static final int ROW_SPAN = 5;
	private static final int ROW_MERGED = 7;
	private static final int ROW_EMPTY = 9;
	private static final int ROW_AMOUNT = 10;
	private static final int ROW_REMARK = 11;
	private static final int COL_VALUE = 1;

	@Rule
	public TemporaryFolder temporaryFolder = new TemporaryFolder();

	@Rule
	public TestName testName = new TestName();

	@Test
	public void shouldExportSingleValueImagesForXlsx() throws Exception {
		assertTemplateExport("template_image.xlsx");
	}

	@Test
	public void shouldExportSingleValueImagesForXls() throws Exception {
		assertTemplateExport("template_image.xls");
	}

	/**
	 * issue 原始场景：仅 {{name}}、{{jyy}} 两个占位符
	 */
	@Test
	public void shouldExportImageForIssueScenario() throws Exception {
		File template = temporaryFolder.newFile("test.xlsx");
		try (XSSFWorkbook workbook = new XSSFWorkbook(); FileOutputStream out = new FileOutputStream(template)) {
			Sheet sheet = workbook.createSheet("sheet1");
			Row row0 = sheet.createRow(0);
			row0.createCell(0).setCellValue("姓名");
			row0.createCell(1).setCellValue("{{name}}");
			Row row1 = sheet.createRow(1);
			row1.createCell(0).setCellValue("签名");
			row1.createCell(1).setCellValue("{{jyy}}");
			workbook.write(out);
		}

		ImageEntity image = new ImageEntity();
		image.setHeight(200);
		image.setWidth(500);
		image.setUrl(resourcePath(IMAGE_PATH));
		Map<String, Object> map = new HashMap<>();
		map.put("jyy", image);
		map.put("name", "图片导出");

		TemplateExportParams params = new TemplateExportParams(template.getAbsolutePath());
		params.setSheetName("测试导出图片");
		Workbook result = ExcelExportUtil.exportExcel(params, map);
		assertNotNull(result);
		Sheet sheet = result.getSheetAt(0);
		assertEquals("测试导出图片", sheet.getSheetName());
		assertEquals("图片导出", sheet.getRow(0).getCell(1).getStringCellValue());
		assertEquals("", sheet.getRow(1).getCell(1).getStringCellValue());
		List<Picture> pictures = pictures(sheet);
		assertEquals(1, pictures.size());
		assertAnchor(findPicture(pictures, 1, 1), 1, 1, 1, 1);
		assertTrue("行高应按 ImageEntity.height 撑开", sheet.getRow(1).getHeight() >= 200);
		result.close();
	}

	/**
	 * ImageEntity 既没有 url 也没有 data 时，单元格置空且不插图、不报错
	 */
	@Test
	public void shouldClearCellWhenImageHasNoData() throws Exception {
		Map<String, Object> map = fullData();
		map.put("sign", new ImageEntity());
		Workbook result = export("template_image.xlsx", map);
		Sheet sheet = result.getSheetAt(0);
		assertEquals("", cellText(sheet, ROW_SIGN));
		assertNull(findPicture(pictures(sheet), ROW_SIGN, COL_VALUE));
		assertEquals("张三", cellText(sheet, ROW_NAME));
		assertEquals(5, pictures(sheet).size());
		result.close();
	}

	/**
	 * 未传入图片占位符时保持原有行为：替换为空字符串
	 */
	@Test
	public void shouldKeepTextBehaviorWhenNoImageProvided() throws Exception {
		Map<String, Object> map = new HashMap<>();
		map.put("name", "李四");
		map.put("amount", "88");
		Workbook result = export("template_image.xlsx", map);
		Sheet sheet = result.getSheetAt(0);
		assertEquals("李四", cellText(sheet, ROW_NAME));
		assertEquals("", cellText(sheet, ROW_SIGN));
		assertEquals("签名人：李四", cellText(sheet, ROW_REMARK));
		assertEquals(88D, sheet.getRow(ROW_AMOUNT).getCell(COL_VALUE).getNumericCellValue(), 0.0001);
		assertTrue(pictures(sheet).isEmpty());
		result.close();
	}

	private void assertTemplateExport(String templateName) throws Exception {
		Workbook result = export(templateName, fullData());
		Sheet sheet = result.getSheetAt(0);

		// 文本 / 数值 / 混合文本 / 空值
		assertEquals("张三", cellText(sheet, ROW_NAME));
		assertEquals("签名人：张三", cellText(sheet, ROW_REMARK));
		assertEquals("", cellText(sheet, ROW_EMPTY));
		assertEquals(CellType.NUMERIC, sheet.getRow(ROW_AMOUNT).getCell(COL_VALUE).getCellType());
		assertEquals(1234.5D, sheet.getRow(ROW_AMOUNT).getCell(COL_VALUE).getNumericCellValue(), 0.0001);

		// 图片单元格不再输出 ImageEntity@xxx
		for (int row : new int[]{ROW_SIGN, ROW_FIT, ROW_ORIGINAL, ROW_BYTES, ROW_SPAN, ROW_MERGED}) {
			assertEquals("第" + (row + 1) + "行图片单元格应为空", "", cellText(sheet, row));
		}

		List<Picture> pictures = pictures(sheet);
		assertEquals(6, pictures.size());
		assertAnchor(findPicture(pictures, ROW_SIGN, COL_VALUE), ROW_SIGN, COL_VALUE, 1, 1);
		assertAnchor(findPicture(pictures, ROW_BYTES, COL_VALUE), ROW_BYTES, COL_VALUE, 1, 1);
		assertAnchor(findPicture(pictures, ROW_SPAN, COL_VALUE), ROW_SPAN, COL_VALUE, 2, 2);
		assertAnchor(findPicture(pictures, ROW_MERGED, COL_VALUE), ROW_MERGED, COL_VALUE, 2, 2);
		assertNotNull(findPicture(pictures, ROW_FIT, COL_VALUE));
		assertNotNull(findPicture(pictures, ROW_ORIGINAL, COL_VALUE));

		// 代码指定 rowspan/colspan 时自动合并 B6:C7；模板已合并的 B8:C9 不重复合并
		assertEquals(1, countMergedRegion(sheet, ROW_SPAN, ROW_SPAN + 1, COL_VALUE, COL_VALUE + 1));
		assertEquals(1, countMergedRegion(sheet, ROW_MERGED, ROW_MERGED + 1, COL_VALUE, COL_VALUE + 1));
		result.close();
	}

	private Workbook export(String templateName, Map<String, Object> map) throws Exception {
		TemplateExportParams params = new TemplateExportParams(resourcePath(TEMPLATE_DIR + templateName));
		Workbook result = ExcelExportUtil.exportExcel(params, map);
		assertNotNull("模板导出失败，返回的 Workbook 为空", result);
		saveOutput(result, templateName);
		return result;
	}

	private Map<String, Object> baseData() {
		Map<String, Object> map = new HashMap<>();
		map.put("name", "张三");
		map.put("amount", "1234.5");
		return map;
	}

	private Map<String, Object> fullData() throws Exception {
		String imagePath = resourcePath(IMAGE_PATH);
		Map<String, Object> map = baseData();
		map.put("sign", image(imagePath, ImageScaleMode.STRETCH));
		map.put("fitImg", image(imagePath, ImageScaleMode.FIT));
		// 原始大小用小尺寸签名图，避免遮挡下方单元格
		ImageEntity originalImg = new ImageEntity();
		originalImg.setData(signatureBytes());
		originalImg.setScaleModeEnum(ImageScaleMode.ORIGINAL);
		map.put("originalImg", originalImg);

		ImageEntity bytesImg = new ImageEntity();
		bytesImg.setData(signatureBytes());
		bytesImg.setHeight(600);
		map.put("bytesImg", bytesImg);

		ImageEntity spanImg = image(imagePath, ImageScaleMode.STRETCH);
		spanImg.setRowspan(2);
		spanImg.setColspan(2);
		map.put("spanImg", spanImg);

		ImageEntity mergedImg = image(imagePath, ImageScaleMode.STRETCH);
		mergedImg.setRowspan(2);
		mergedImg.setColspan(2);
		map.put("mergedImg", mergedImg);

		map.put("emptyImg", null);
		return map;
	}

	private ImageEntity image(String url, ImageScaleMode scaleMode) {
		ImageEntity image = new ImageEntity();
		image.setUrl(url);
		image.setHeight(600);
		image.setScaleModeEnum(scaleMode);
		return image;
	}

	/**
	 * 生成一张模拟手写签名的 PNG
	 */
	private byte[] signatureBytes() throws Exception {
		BufferedImage image = new BufferedImage(240, 80, BufferedImage.TYPE_INT_RGB);
		Graphics2D graphics = image.createGraphics();
		graphics.setColor(Color.WHITE);
		graphics.fillRect(0, 0, 240, 80);
		graphics.setColor(new Color(20, 40, 160));
		graphics.setStroke(new BasicStroke(4));
		graphics.drawPolyline(new int[]{10, 40, 60, 90, 120, 150, 180, 230}, new int[]{60, 20, 55, 25, 60, 15, 50, 30}, 8);
		graphics.dispose();
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		ImageIO.write(image, "png", out);
		return out.toByteArray();
	}

	private String resourcePath(String path) throws Exception {
		return new File(getClass().getResource(path).toURI()).getAbsolutePath();
	}

	private String cellText(Sheet sheet, int row) {
		return sheet.getRow(row).getCell(COL_VALUE).getStringCellValue();
	}

	private List<Picture> pictures(Sheet sheet) {
		List<Picture> pictures = new ArrayList<>();
		Drawing<?> drawing = sheet.getDrawingPatriarch();
		if (drawing == null) {
			return pictures;
		}
		for (Shape shape : drawing) {
			if (shape instanceof Picture) {
				pictures.add((Picture) shape);
			}
		}
		return pictures;
	}

	private Picture findPicture(List<Picture> pictures, int row, int col) {
		for (Picture picture : pictures) {
			ClientAnchor anchor = picture.getClientAnchor();
			if (anchor.getRow1() == row && anchor.getCol1() == col) {
				return picture;
			}
		}
		return null;
	}

	private void assertAnchor(Picture picture, int row, int col, int rowspan, int colspan) {
		assertNotNull("第" + (row + 1) + "行未插入图片", picture);
		ClientAnchor anchor = picture.getClientAnchor();
		assertEquals(row, anchor.getRow1());
		assertEquals(col, anchor.getCol1());
		assertEquals(row + rowspan, anchor.getRow2());
		assertEquals(col + colspan, anchor.getCol2());
	}

	private int countMergedRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
		int count = 0;
		for (CellRangeAddress region : sheet.getMergedRegions()) {
			if (region.getFirstRow() == firstRow && region.getLastRow() == lastRow
					&& region.getFirstColumn() == firstCol && region.getLastColumn() == lastCol) {
				count++;
			}
		}
		return count;
	}

	private void saveOutput(Workbook workbook, String templateName) throws Exception {
		OUTPUT_DIR.mkdirs();
		try (OutputStream out = new FileOutputStream(new File(OUTPUT_DIR, testName.getMethodName() + "_" + templateName))) {
			workbook.write(out);
		}
	}
}
