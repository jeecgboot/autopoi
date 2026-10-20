package org.jeecgframework.poi.util;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.excel.ExcelImportUtil;
import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.junit.Test;

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

/**
 * 读取 sheet 中的图片时，sheet 里含有非图片形状（文本框、箭头、组合图形等）不应报错，图片仍应被读取。
 *
 * @since 2026-10-20 issues/83 issues/73 issues/107 XSSFSimpleShape / XSSFShapeGroup 强转成 XSSFPicture 报错
 */
public class SheetPicturesWithShapesTest {

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

	private static byte[] pngBytes() throws Exception {
		BufferedImage image = new BufferedImage(20, 20, BufferedImage.TYPE_INT_RGB);
		Graphics2D g = image.createGraphics();
		g.setColor(Color.RED);
		g.fillRect(0, 0, 20, 20);
		g.dispose();
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		ImageIO.write(image, "png", bos);
		return bos.toByteArray();
	}

	private static InputStream toStream(Workbook wb) throws Exception {
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		wb.write(bos);
		wb.close();
		return new ByteArrayInputStream(bos.toByteArray());
	}

	private static XSSFWorkbook xlsxWithHeader() {
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet("Sheet1");
		sheet.createRow(0).createCell(0).setCellValue("名称");
		sheet.createRow(1).createCell(0).setCellValue("a");
		return wb;
	}

	private static HSSFWorkbook xlsWithHeader() {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet("Sheet1");
		sheet.createRow(0).createCell(0).setCellValue("名称");
		sheet.createRow(1).createCell(0).setCellValue("a");
		return wb;
	}

	// ------------------------------------------------------------- xlsx

	@Test
	public void xlsxSimpleShapeShouldBeIgnored() throws Exception {
		XSSFWorkbook wb = xlsxWithHeader();
		XSSFDrawing drawing = wb.getSheetAt(0).createDrawingPatriarch();
		drawing.createSimpleShape(drawing.createAnchor(0, 0, 0, 0, 3, 3, 5, 5));
		Map<String, PictureData> pictures = PoiPublicUtil.getSheetPictrues07(wb.getSheetAt(0), wb);
		assertTrue(pictures.isEmpty());
	}

	@Test
	public void xlsxShapeGroupShouldBeIgnored() throws Exception {
		XSSFWorkbook wb = xlsxWithHeader();
		XSSFDrawing drawing = wb.getSheetAt(0).createDrawingPatriarch();
		drawing.createGroup(drawing.createAnchor(0, 0, 0, 0, 3, 3, 5, 5));
		Map<String, PictureData> pictures = PoiPublicUtil.getSheetPictrues07(wb.getSheetAt(0), wb);
		assertTrue(pictures.isEmpty());
	}

	@Test
	public void xlsxPictureShouldStillBeReadNextToShapes() throws Exception {
		XSSFWorkbook wb = xlsxWithHeader();
		XSSFDrawing drawing = wb.getSheetAt(0).createDrawingPatriarch();
		drawing.createSimpleShape(drawing.createAnchor(0, 0, 0, 0, 3, 3, 5, 5));
		int pictureIndex = wb.addPicture(pngBytes(), Workbook.PICTURE_TYPE_PNG);
		drawing.createPicture(drawing.createAnchor(0, 0, 0, 0, 1, 1, 2, 2), pictureIndex);
		Map<String, PictureData> pictures = PoiPublicUtil.getSheetPictrues07(wb.getSheetAt(0), wb);
		assertEquals(1, pictures.size());
		assertNotNull(pictures.get("1_1"));
	}

	@Test
	public void xlsxImportWithShapesShouldWork() throws Exception {
		XSSFWorkbook wb = xlsxWithHeader();
		XSSFDrawing drawing = wb.getSheetAt(0).createDrawingPatriarch();
		drawing.createSimpleShape(drawing.createAnchor(0, 0, 0, 0, 3, 3, 5, 5));
		drawing.createGroup(drawing.createAnchor(0, 0, 0, 0, 6, 6, 8, 8));
		List<NameEntity> list = ExcelImportUtil.importExcel(toStream(wb), NameEntity.class, new ImportParams());
		assertEquals(1, list.size());
		assertEquals("a", list.get(0).getName());
	}

	// ------------------------------------------------------------- xls

	@Test
	public void xlsSimpleShapeNextToPictureShouldBeIgnored() throws Exception {
		HSSFWorkbook wb = xlsWithHeader();
		HSSFSheet sheet = wb.getSheetAt(0);
		HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
		patriarch.createSimpleShape(new HSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 5));
		int pictureIndex = wb.addPicture(pngBytes(), Workbook.PICTURE_TYPE_PNG);
		patriarch.createPicture(new HSSFClientAnchor(0, 0, 0, 0, (short) 1, 1, (short) 2, 2), pictureIndex);
		Map<String, PictureData> pictures = PoiPublicUtil.getSheetPictrues03(sheet, wb);
		assertEquals(1, pictures.size());
		assertNotNull(pictures.get("1_1"));
	}

	@Test
	public void xlsSheetWithoutDrawingShouldNotFailWhenWorkbookHasPictures() throws Exception {
		HSSFWorkbook wb = xlsWithHeader();
		HSSFSheet withPicture = wb.getSheetAt(0);
		int pictureIndex = wb.addPicture(pngBytes(), Workbook.PICTURE_TYPE_PNG);
		withPicture.createDrawingPatriarch().createPicture(new HSSFClientAnchor(0, 0, 0, 0, (short) 1, 1, (short) 2, 2), pictureIndex);
		HSSFSheet withoutDrawing = wb.createSheet("Sheet2");
		withoutDrawing.createRow(0).createCell(0).setCellValue("名称");
		Map<String, PictureData> pictures = PoiPublicUtil.getSheetPictrues03(withoutDrawing, wb);
		assertTrue(pictures == null || pictures.isEmpty());
	}
}
