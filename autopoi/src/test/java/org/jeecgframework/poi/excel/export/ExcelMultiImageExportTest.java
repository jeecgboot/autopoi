package org.jeecgframework.poi.excel.export;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Shape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.ExcelImportUtil;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.jeecgframework.poi.handler.impl.ExcelDataHandlerDefaultImpl;
import org.jeecgframework.poi.util.PoiPublicUtil;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

/**
 * 单元格多图片导出测试。
 *
 * @author scott
 * @since 2026-08-13 issues/9605 Excel 多图片导入导出回归测试
 */
public class ExcelMultiImageExportTest {

	@Rule
	public TemporaryFolder temporaryFolder = new TemporaryFolder();

	@Test
	public void shouldExportAllImagesInOneCell() throws Exception {
		File firstImage = createImage("first.png", Color.RED);
		File secondImage = createImage("second.png", Color.BLUE);
		String imagePaths = firstImage.getName() + "," + secondImage.getName();

		assertImageCount(ExcelType.XSSF, imagePaths, 2);
		assertImageCount(ExcelType.HSSF, imagePaths, 2);
	}

	@Test
	public void shouldKeepSingleImageExport() throws Exception {
		File image = createImage("single.png", Color.GREEN);

		assertImageCount(ExcelType.XSSF, image.getName(), 1);
		assertImageCount(ExcelType.HSSF, image.getName(), 1);
	}

	@Test
	public void shouldKeepSingleEmbeddedImageImport() throws Exception {
		File image = createImage("single-import.png", Color.GREEN);

		assertImportedImageCount(ExcelType.XSSF, image.getName(), 1);
		assertImportedImageCount(ExcelType.HSSF, image.getName(), 1);
	}

	@Test
	public void shouldExportAndImportThreeImagesInOneCell() throws Exception {
		File firstImage = createImage("three-first.png", Color.RED);
		File secondImage = createImage("three-second.png", Color.GREEN);
		File thirdImage = createImage("three-third.png", Color.BLUE);
		String imagePaths = String.join(",", firstImage.getName(), secondImage.getName(), thirdImage.getName());

		assertImageCount(ExcelType.XSSF, imagePaths, 3);
		assertImageCount(ExcelType.HSSF, imagePaths, 3);
		assertImportedImageCount(ExcelType.XSSF, imagePaths, 3);
		assertImportedImageCount(ExcelType.HSSF, imagePaths, 3);
	}

	@Test
	public void shouldIgnoreBlankImagePathSegments() throws Exception {
		File firstImage = createImage("blank-first.png", Color.RED);
		File secondImage = createImage("blank-second.png", Color.BLUE);
		String imagePaths = " , " + firstImage.getName() + ",, " + secondImage.getName() + " , ";

		assertImageCount(ExcelType.XSSF, imagePaths, 2);
		assertImageCount(ExcelType.HSSF, imagePaths, 2);
	}

	@Test
	public void shouldSkipMissingImageAndExportValidImage() throws Exception {
		File validImage = createImage("valid.png", Color.GREEN);
		String imagePaths = "missing.png," + validImage.getName();

		assertImageCount(ExcelType.XSSF, imagePaths, 1);
		assertImageCount(ExcelType.HSSF, imagePaths, 1);
	}

	@Test
	public void shouldImportAllEmbeddedImagesInOneCell() throws Exception {
		File firstImage = createImage("import-first.png", Color.RED);
		File secondImage = createImage("import-second.png", Color.BLUE);
		String imagePaths = firstImage.getName() + "," + secondImage.getName();

		assertImportedImageCount(ExcelType.XSSF, imagePaths, 2);
		assertImportedImageCount(ExcelType.HSSF, imagePaths, 2);
	}

	@Test
	public void shouldFallbackToImagePathWhenCellHasNoPicture() throws Exception {
		String imagePaths = "upload/first.png,upload/second.png";

		assertImportedImagePath(ExcelType.XSSF, imagePaths);
		assertImportedImagePath(ExcelType.HSSF, imagePaths);
	}

	@Test
	public void shouldKeepPlainTextImportWhenColumnIsNotMarkedAsImage() throws Exception {
		String imagePaths = "upload/first.png,upload/second.png";

		assertImportedPlainText(ExcelType.XSSF, imagePaths);
		assertImportedPlainText(ExcelType.HSSF, imagePaths);
	}

	@Test
	public void shouldKeepLegacyPictureKeyAndExposeAllCellPictures() throws Exception {
		File firstImage = createImage("legacy-first.png", Color.RED);
		File secondImage = createImage("legacy-second.png", Color.BLUE);
		String imagePaths = firstImage.getName() + "," + secondImage.getName();

		assertLegacyPictureLookup(ExcelType.XSSF, imagePaths);
		assertLegacyPictureLookup(ExcelType.HSSF, imagePaths);
	}

	private void assertImageCount(ExcelType excelType, String imagePaths, int expectedCount) throws Exception {
		try (Workbook workbook = exportWorkbook(excelType, imagePaths, true)) {
			assertEquals(expectedCount, workbook.getAllPictures().size());
			List<ClientAnchor> anchors = getPictureAnchors(workbook);
			assertEquals(expectedCount, anchors.size());
			if (expectedCount == 1) {
				ClientAnchor anchor = anchors.get(0);
				assertEquals(anchor.getCol1() + 1, anchor.getCol2());
				assertEquals(anchor.getRow1() + 1, anchor.getRow2());
			}
			for (int index = 1; index < anchors.size(); index++) {
				ClientAnchor previous = anchors.get(index - 1);
				ClientAnchor current = anchors.get(index);
				assertEquals(previous.getCol1(), current.getCol1());
				assertEquals(previous.getRow1(), current.getRow1());
				assertEquals(previous.getDx2(), current.getDx1());
			}
		}
	}

	private void assertImportedImageCount(ExcelType excelType, String imagePaths, int expectedCount) throws Exception {
		CountingImageHandler dataHandler = new CountingImageHandler();
		List<Map<String, Object>> rows = importWorkbook(exportWorkbook(excelType, imagePaths, true), dataHandler);

		assertEquals(1, rows.size());
		assertEquals(expectedCount, dataHandler.getImageCount());
	}

	private void assertImportedImagePath(ExcelType excelType, String imagePaths) throws Exception {
		CountingImageHandler dataHandler = new CountingImageHandler();
		List<Map<String, Object>> rows = importWorkbook(exportWorkbook(excelType, imagePaths, false), dataHandler);

		assertEquals(1, rows.size());
		assertEquals(0, dataHandler.getImageCount());
		assertEquals(imagePaths, rows.get(0).get("图片"));
	}

	private void assertImportedPlainText(ExcelType excelType, String imagePaths) throws Exception {
		CountingImageHandler dataHandler = new CountingImageHandler();
		List<Map<String, Object>> rows = importWorkbook(exportWorkbook(excelType, imagePaths, false), dataHandler, false);

		assertEquals(1, rows.size());
		assertEquals(0, dataHandler.getImageCount());
		assertEquals(imagePaths, rows.get(0).get("图片"));
	}

	private void assertLegacyPictureLookup(ExcelType excelType, String imagePaths) throws Exception {
		try (Workbook workbook = exportWorkbook(excelType, imagePaths, true)) {
			ClientAnchor firstAnchor = getPictureAnchors(workbook).get(0);
			String legacyKey = firstAnchor.getRow1() + "_" + firstAnchor.getCol1();
			Map<String, PictureData> pictures;
			if (workbook instanceof XSSFWorkbook) {
				pictures = PoiPublicUtil.getSheetPictrues07((XSSFSheet) workbook.getSheetAt(0), (XSSFWorkbook) workbook);
			} else {
				pictures = PoiPublicUtil.getSheetPictrues03((HSSFSheet) workbook.getSheetAt(0), (HSSFWorkbook) workbook);
			}

			assertNotNull(pictures.get(legacyKey));
			assertEquals(2, PoiPublicUtil.getPictureDataList(pictures, legacyKey).size());
		}
	}

	private List<ClientAnchor> getPictureAnchors(Workbook workbook) {
		Drawing<?> drawing = workbook.getSheetAt(0).getDrawingPatriarch();
		List<ClientAnchor> anchors = new ArrayList<>();
		for (Shape shape : drawing) {
			if (shape instanceof Picture) {
				anchors.add(((Picture) shape).getClientAnchor());
			}
		}
		return anchors;
	}

	private List<Map<String, Object>> importWorkbook(Workbook workbook, CountingImageHandler dataHandler) throws Exception {
		return importWorkbook(workbook, dataHandler, true);
	}

	private List<Map<String, Object>> importWorkbook(Workbook workbook, CountingImageHandler dataHandler, boolean imageColumn) throws Exception {
		try (Workbook closeableWorkbook = workbook; ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
			closeableWorkbook.write(outputStream);
			ImportParams importParams = new ImportParams();
			if (imageColumn) {
				importParams.setImageList(Collections.singletonList("图片"));
			}
			importParams.setDataHanlder(dataHandler);
			return ExcelImportUtil.importExcel(new ByteArrayInputStream(outputStream.toByteArray()), Map.class, importParams);
		}
	}

	private Workbook exportWorkbook(ExcelType excelType, String imagePaths, boolean embed) {
		ExcelExportEntity imageEntity = new ExcelExportEntity("图片", "images");
		imageEntity.setType(embed ? 2 : 1);
		if (embed) {
			imageEntity.setExportImageType(3);
			imageEntity.setImageBasePath(temporaryFolder.getRoot().getAbsolutePath());
		}
		imageEntity.setWidth(60);
		imageEntity.setHeight(50);

		Map<String, Object> row = new HashMap<>();
		row.put("images", imagePaths);
		ExportParams exportParams = new ExportParams(null, "多图片", excelType);
		List<ExcelExportEntity> entityList = new ArrayList<>();
		entityList.add(imageEntity);
		List<Map<String, Object>> dataList = new ArrayList<>();
		dataList.add(row);
		return ExcelExportUtil.exportExcel(exportParams, entityList, dataList);
	}

	private File createImage(String fileName, Color color) throws Exception {
		File imageFile = temporaryFolder.newFile(fileName);
		BufferedImage image = new BufferedImage(20, 20, BufferedImage.TYPE_INT_RGB);
		Graphics2D graphics = image.createGraphics();
		try {
			graphics.setColor(color);
			graphics.fillRect(0, 0, image.getWidth(), image.getHeight());
		} finally {
			graphics.dispose();
		}
		ImageIO.write(image, "png", imageFile);
		return imageFile;
	}

	private static class CountingImageHandler extends ExcelDataHandlerDefaultImpl {

		private int imageCount;

		@Override
		public void setMapValue(Map<String, Object> map, String originKey, Object value) {
			if (value instanceof byte[]) {
				imageCount++;
				map.put(originKey, "image-" + imageCount);
				return;
			}
			super.setMapValue(map, originKey, value);
		}

		private int getImageCount() {
			return imageCount;
		}
	}
}
