/**
 * Copyright 2013-2015 JEECG (jeecgos@163.com)
 *   
 *  Licensed under the Apache License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 *  Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.jeecgframework.poi.util;

import java.awt.image.BufferedImage;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.net.URI;
import java.net.URISyntaxException;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import javax.imageio.ImageIO;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jeecgframework.poi.excel.annotation.Excel;
import org.jeecgframework.poi.excel.annotation.ExcelCollection;
import org.jeecgframework.poi.excel.annotation.ExcelEntity;
import org.jeecgframework.poi.excel.annotation.ExcelIgnore;
import org.jeecgframework.poi.excel.entity.vo.PoiBaseConstants;
import org.jeecgframework.poi.word.entity.WordImageEntity;
import org.jeecgframework.poi.word.entity.params.ExcelListEntity;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.ClassUtils;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

/**
 * AutoPoi 的公共基础类
 * 
 * @author JEECG
 * @date 2015年4月5日 上午12:59:22
 */
public final class PoiPublicUtil {

	private static final Logger LOGGER = LoggerFactory.getLogger(PoiPublicUtil.class);

	private PoiPublicUtil() {

	}

	@SuppressWarnings({ "unchecked" })
	public static <K, V> Map<K, V> mapFor(Object... mapping) {
		Map<K, V> map = new HashMap<K, V>();
		for (int i = 0; i < mapping.length; i += 2) {
			map.put((K) mapping[i], (V) mapping[i + 1]);
		}
		return map;
	}

	/**
	 * 彻底创建一个对象
	 * 
	 * @param clazz
	 * @return
	 */
	public static Object createObject(Class<?> clazz, String targetId) {
		Object obj = null;
		Method setMethod;
		try {
			if (clazz.equals(Map.class)) {
				return new LinkedHashMap<String, Object>();
			}
			obj = clazz.newInstance();
			Field[] fields = getClassFields(clazz);
			for (Field field : fields) {
				if (isNotUserExcelUserThis(null, field, targetId)) {
					continue;
				}
				if (isCollection(field.getType())) {
					ExcelCollection collection = field.getAnnotation(ExcelCollection.class);
					setMethod = getMethod(field.getName(), clazz, field.getType());
					setMethod.invoke(obj, collection.type().newInstance());
				} else if (!isJavaClass(field)) {
					setMethod = getMethod(field.getName(), clazz, field.getType());
					setMethod.invoke(obj, createObject(field.getType(), targetId));
				}
			}

		} catch (Exception e) {
			LOGGER.error(e.getMessage(), e);
			throw new RuntimeException("创建对象异常");
		}
		return obj;

	}

	/**
	 * 获取class的 包括父类的
	 * 
	 * @param clazz
	 * @return
	 */
	public static Field[] getClassFields(Class<?> clazz) {
		List<Field> list = new ArrayList<Field>();
		Field[] fields;
		do {
			fields = clazz.getDeclaredFields();
			for (int i = 0; i < fields.length; i++) {
				list.add(fields[i]);
			}
			clazz = clazz.getSuperclass();
		} while (clazz != Object.class && clazz != null);
		return list.toArray(fields);
	}

	/**
	 * @param photoByte
	 * @return
	 */
	public static String getFileExtendName(byte[] photoByte) {
		String strFileExtendName = "JPG";
		if ((photoByte[0] == 71) && (photoByte[1] == 73) && (photoByte[2] == 70) && (photoByte[3] == 56) && ((photoByte[4] == 55) || (photoByte[4] == 57)) && (photoByte[5] == 97)) {
			strFileExtendName = "GIF";
		} else if ((photoByte[6] == 74) && (photoByte[7] == 70) && (photoByte[8] == 73) && (photoByte[9] == 70)) {
			strFileExtendName = "JPG";
		} else if ((photoByte[0] == 66) && (photoByte[1] == 77)) {
			strFileExtendName = "BMP";
		} else if ((photoByte[1] == 80) && (photoByte[2] == 78) && (photoByte[3] == 71)) {
			strFileExtendName = "PNG";
		}
		return strFileExtendName;
	}

	/**
	 * 获取GET方法
	 * 
	 * @param name
	 * @param pojoClass
	 * @return
	 * @throws Exception
	 */
	public static Method getMethod(String name, Class<?> pojoClass) throws Exception {
		StringBuffer getMethodName = new StringBuffer(PoiBaseConstants.GET);
		getMethodName.append(name.substring(0, 1).toUpperCase());
		getMethodName.append(name.substring(1));
		Method method = null;
		try {
			method = pojoClass.getMethod(getMethodName.toString(), new Class[] {});
		} catch (Exception e) {
			method = pojoClass.getMethod(getMethodName.toString().replace(PoiBaseConstants.GET, PoiBaseConstants.IS), new Class[] {});
		}
		return method;
	}

	/**
	 * 获取SET方法
	 * 
	 * @param name
	 * @param pojoClass
	 * @param type
	 * @return
	 * @throws Exception
	 */
	public static Method getMethod(String name, Class<?> pojoClass, Class<?> type) throws Exception {
		StringBuffer getMethodName = new StringBuffer(PoiBaseConstants.SET);
		getMethodName.append(name.substring(0, 1).toUpperCase());
		getMethodName.append(name.substring(1));
		return pojoClass.getMethod(getMethodName.toString(), new Class[] { type });
	}
	
	//update-begin-author:taoyan date:20180615 for:TASK #2798 导入扩展方法，支持自定义导入字段转换规则
	/**
	 * 获取get方法 通过EXCEL注解exportConvert判断是否支持值的转换
	 * @param name
	 * @param pojoClass
	 * @param convert
	 * @return
	 * @throws Exception
	 */
	public static Method getMethod(String name, Class<?> pojoClass,boolean convert) throws Exception {
		StringBuffer getMethodName = new StringBuffer();
		if(convert){
			getMethodName.append(PoiBaseConstants.CONVERT);
		}
		getMethodName.append(PoiBaseConstants.GET);
		getMethodName.append(name.substring(0, 1).toUpperCase());
		getMethodName.append(name.substring(1));
		Method method = null;
		try {
			method = pojoClass.getMethod(getMethodName.toString(), new Class[] {});
		} catch (Exception e) {
			method = pojoClass.getMethod(getMethodName.toString().replace(PoiBaseConstants.GET, PoiBaseConstants.IS), new Class[] {});
		}
		return method;
	}
	
	/**
	 * 获取set方法  通过EXCEL注解importConvert判断是否支持值的转换
	 * @param name
	 * @param pojoClass
	 * @param type
	 * @param convert
	 * @return
	 * @throws Exception
	 */
	public static Method getMethod(String name, Class<?> pojoClass, Class<?> type,boolean convert) throws Exception {
		StringBuffer setMethodName = new StringBuffer();
		if(convert){
			setMethodName.append(PoiBaseConstants.CONVERT);
		}
		setMethodName.append(PoiBaseConstants.SET);
		setMethodName.append(name.substring(0, 1).toUpperCase());
		setMethodName.append(name.substring(1));
		return pojoClass.getMethod(setMethodName.toString(), new Class[] { type });
	}
	//update-end-author:taoyan date:20180615 for:TASK #2798 导入扩展方法，支持自定义导入字段转换规则
	
	/**
	 * 获取Excel2003图片
	 * 
	 * @param sheet
	 *            当前sheet对象
	 * @param workbook
	 *            工作簿对象
	 * @return Map key:图片单元格索引（1_1）String，value:图片流PictureData
	 */
	public static Map<String, PictureData> getSheetPictrues03(HSSFSheet sheet, HSSFWorkbook workbook) {
		Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
		List<HSSFPictureData> pictures = workbook.getAllPictures();
		if (!pictures.isEmpty()) {
			for (HSSFShape shape : sheet.getDrawingPatriarch().getChildren()) {
				HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
				if (shape instanceof HSSFPicture) {
					HSSFPicture pic = (HSSFPicture) shape;
					int pictureIndex = pic.getPictureIndex() - 1;
					HSSFPictureData picData = pictures.get(pictureIndex);
					String picIndex = String.valueOf(anchor.getRow1()) + "_" + String.valueOf(anchor.getCol1());
					sheetIndexPicMap.put(picIndex, picData);
				}
			}
			return sheetIndexPicMap;
		} else {
			return null;
		}
	}

	/**
	 * 获取Excel2007图片
	 * 
	 * @param sheet
	 *            当前sheet对象
	 * @param workbook
	 *            工作簿对象
	 * @return Map key:图片单元格索引（1_1）String，value:图片流PictureData
	 */
	public static Map<String, PictureData> getSheetPictrues07(XSSFSheet sheet, XSSFWorkbook workbook) {
		Map<String, PictureData> sheetIndexPicMap = new HashMap<String, PictureData>();
		for (POIXMLDocumentPart dr : sheet.getRelations()) {
			if (dr instanceof XSSFDrawing) {
				XSSFDrawing drawing = (XSSFDrawing) dr;
				List<XSSFShape> shapes = drawing.getShapes();
				for (XSSFShape shape : shapes) {
					XSSFPicture pic = (XSSFPicture) shape;
					XSSFClientAnchor anchor = pic.getPreferredSize();
					CTMarker ctMarker = anchor.getFrom();
					String picIndex = ctMarker.getRow() + "_" + ctMarker.getCol();
					sheetIndexPicMap.put(picIndex, pic.getPictureData());
				}
			}
		}
		return sheetIndexPicMap;
	}

	/**
	 *
	 * 获取嵌入图片 <br/>
	 * 支持excel2007+版本.
	 * @param sheet
	 * @param isCopy
	 * @param book
	 * @return
	 * @author chenrui
	 * @date 2024/4/2 20:25
	 */
    public static Map<String, PictureData> getCellImages(Sheet sheet, ByteArrayOutputStream isCopy,Workbook book) {
        // 获取所有嵌入图片的单元格的内容 date:2024/4/2
        Map<String, CellImage> cellImageMap = new HashMap<>();
        Iterator<Row> rows = sheet.rowIterator();
        while (rows.hasNext()) {
            Row row = rows.next();
            Iterator<Cell> cells = row.cellIterator();
            while (cells.hasNext()) {
                Cell cell = cells.next();
				CellType cellType = cell.getCellType();
				if(cellType.equals(CellType.FORMULA)) {
					CellType resultType = cell.getCachedFormulaResultType();
					if(!resultType.equals(CellType.STRING)){
						continue;
					}
					String cellVal = cell.getStringCellValue();
					if (null != cellVal && cellVal.startsWith("=DISPIMG")) {
						int start = cellVal.indexOf("\"");
						int end = cellVal.lastIndexOf("\"");
						if (start != -1 && end != -1) {
							String imgId = cellVal.substring(start + 1, end);
							CellImage cellImage = new CellImage();
							cellImage.setImgId(imgId);
							cellImage.setCellStr(cellVal);
							cellImageMap.put(row.getRowNum() + "_" + cell.getColumnIndex(), cellImage);
						}
					}
				}
            }
        }

        try (ZipInputStream zis = new ZipInputStream(new ByteArrayInputStream(isCopy.toByteArray()));
			 ZipInputStream fzis = new ZipInputStream(new ByteArrayInputStream(isCopy.toByteArray()))) {
			//update-begin---author:chenrui ---date:20240407  for：[QQYUN-8898]不依赖hutool,xml解析改为dom------------
			DocumentBuilder documentBuilder = DocumentBuilderFactory.newInstance().newDocumentBuilder();
			ZipEntry entry;
			// 获取嵌入单元格图片的rid
            while ((entry = zis.getNextEntry()) != null) {
				try {
                    final String fileName = entry.getName();
                    if (Objects.equals(fileName, "xl/cellimages.xml")) {
                        String content = IOUtils.toString(zis, StandardCharsets.UTF_8);
						Document document = documentBuilder.parse(new InputSource(new ByteArrayInputStream(content.getBytes(StandardCharsets.UTF_8))));
						NodeList cellImages = document.getElementsByTagName("etc:cellImage");
						if (Objects.isNull(cellImages)) {
							continue;
						}
						for (int i = 0; i < cellImages.getLength(); i++) {
							Node cellImageNode = cellImages.item(i);
							NodeList cNvPr = ((Element) cellImageNode).getElementsByTagName("xdr:cNvPr");
							if(cNvPr.getLength()<1){
								continue;
							}
							Node cNvPrNode = cNvPr.item(0);
							String name = ((Element) cNvPrNode).getAttribute("name");
							if (StringUtils.isNotEmpty(name)) {
								CellImage tempCellimage = cellImageMap.values().stream().filter(item -> Objects.equals(item.getImgId(), name)).findFirst().orElse(null);
								if (Objects.nonNull(tempCellimage)) {
									NodeList blips = ((Element) cellImageNode).getElementsByTagName("a:blip");
									if(blips.getLength()<1){
										continue;
									}
									Node blip = blips.item(0);
									String embed =((Element) blip).getAttribute("r:embed");
									if(embed.isEmpty()){
										continue;
									}
									tempCellimage.setRId(embed);
								}
							}
						}
                    }
                } catch (SAXException e) {
                    throw new RuntimeException(e);
                } finally {
                    zis.closeEntry();
                }
            }
			// 获取嵌入单元格图片的存放位置
            while ((entry = fzis.getNextEntry()) != null) {
                try {
                    final String fileName = entry.getName();
                    if (Objects.equals(fileName, "xl/_rels/cellimages.xml.rels")) {
                        String content = IOUtils.toString(fzis, StandardCharsets.UTF_8);

						Document document = documentBuilder.parse(new InputSource(new ByteArrayInputStream(content.getBytes(StandardCharsets.UTF_8))));
						NodeList relationships = document.getElementsByTagName("Relationship");
                        if (Objects.isNull(relationships)) {
                            continue;
                        }
						for (int i = 0; i < relationships.getLength(); i++) {
							Node relationshipNode = relationships.item(i);
							if(relationshipNode instanceof Element){
								Element relationshipEl = (Element) relationshipNode;
								String id = relationshipEl.getAttribute("Id");
								String target = "/xl/" + relationshipEl.getAttribute("Target");
								if (StringUtils.isNotEmpty(id)) {
									List<CellImage> cellImages = cellImageMap.values().stream().filter(item -> Objects.equals(item.getRId(), id)).collect(Collectors.toList());
									cellImages.stream().filter(Objects::nonNull).forEach(cellImage -> cellImage.setImgName(target));
								}
							}
						}
                    }
                } catch (SAXException e) {
                    throw new RuntimeException(e);
                } finally {
                    fzis.closeEntry();
                }
            }
			// 获取嵌入单元格图片的图片数据
            List<XSSFPictureData> allPictures = (List<XSSFPictureData>) book.getAllPictures();
            for (XSSFPictureData pictureData : allPictures) {
                PackagePartName partName = pictureData.getPackagePart().getPartName();
                URI uri = partName.getURI();
                List<CellImage> cellImages = cellImageMap.values().stream().filter(i -> Objects.equals(i.getImgName(), uri.toString())).collect(Collectors.toList());
				cellImages.stream().filter(Objects::nonNull).forEach(cellImage -> cellImage.setPictureData(pictureData));
            }
			//update-end---author:chenrui ---date:20240407  for：[QQYUN-8898]不依赖hutool,xml解析改为dom------------
		} catch (IOException | ParserConfigurationException e) {
			throw new RuntimeException(e);
		}

        Map<String, PictureData> resp = new HashMap<>();
        if (!cellImageMap.isEmpty()) {
            cellImageMap.forEach((key, cellImage) -> {
                resp.put(key, cellImage.getPictureData());
            });
        }
        return resp;
    }


	/**
	 * 嵌入单元格图片对象
	 *
	 * @author chenrui
	 * @date 2024/4/3 18:27
	 */
    static class CellImage {
		/**
		 * 图片id
		 */
        private String imgId;
		/**
		 * 单元格内容
		 */
        private String cellStr;
		/**
		 * RId
		 */
        private String rId;
		/**
		 * 图片名称
		 */
        private String imgName;
		/**
		 * 图片对象
		 */
        private XSSFPictureData pictureData;

        public String getImgId() {
            return imgId;
        }

        public void setImgId(String imgId) {
            this.imgId = imgId;
        }

        public String getCellStr() {
            return cellStr;
        }

        public void setCellStr(String cellStr) {
            this.cellStr = cellStr;
        }

        public String getRId() {
            return rId;
        }

        public void setRId(String rId) {
            this.rId = rId;
        }

        public String getImgName() {
            return imgName;
        }

        public void setImgName(String imgName) {
            this.imgName = imgName;
        }

        public XSSFPictureData getPictureData() {
            return pictureData;
        }

        public void setPictureData(XSSFPictureData pictureData) {
            this.pictureData = pictureData;
        }

    }

	public static String getWebRootPath(String filePath) {
		try {
			String path = null;
			try {
				path = PoiPublicUtil.class.getClassLoader().getResource("").toURI().getPath();
			} catch (URISyntaxException e) {
				//e.printStackTrace();
			//update-begin-author:taoyan date:20211116 for: JAR包分离 发布出空指针 https://gitee.com/jeecg/jeecg-boot/issues/I4CMHK
			}catch (NullPointerException e) {
				path =  PoiPublicUtil.class.getProtectionDomain().getCodeSource().getLocation().getPath();
			}
			//update-end-author:taoyan date:20211116 for: JAR包分离 发布出空指针 https://gitee.com/jeecg/jeecg-boot/issues/I4CMHK
			//update-begin--Author:zhangdaihao  Date:20190424 for：解决springboot 启动模式，上传路径获取为空问题---------------------
			if (path == null || path == "") {
				//解决springboot 启动模式，上传路径获取为空问题
				path = ClassUtils.getDefaultClassLoader().getResource("").getPath();
			}
			//update-end--Author:zhangdaihao  Date:20190424 for：解决springboot 启动模式，上传路径获取为空问题----------------------
			LOGGER.debug("--- getWebRootPath ----filePath--- " + path);
			path = path.replace("WEB-INF/classes/", "");
			path = path.replace("file:/", "");
			LOGGER.debug("--- path---  " + path);
			LOGGER.debug("--- filePath---  " + filePath);
			return path + filePath;
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 判断是不是集合的实现类
	 * 
	 * @param clazz
	 * @return
	 */
	public static boolean isCollection(Class<?> clazz) {
		return Collection.class.isAssignableFrom(clazz);
	}

	/**
	 * 是不是java基础类
	 * 
	 * @param field
	 * @return
	 */
	public static boolean isJavaClass(Field field) {
		Class<?> fieldType = field.getType();
		boolean isBaseClass = false;
		if (fieldType.isArray()) {
			isBaseClass = false;
		} else if (fieldType.isPrimitive()
				|| fieldType.getPackage() == null
				|| fieldType.getPackage().getName().equals("java.lang")
				|| fieldType.getPackage().getName().equals("java.math")
				|| fieldType.getPackage().getName().equals("java.sql")
				|| fieldType.getPackage().getName().equals("java.util")
				|| fieldType.getPackage().getName().equals("java.time")) {
			isBaseClass = true;
		}
		return isBaseClass;
	}

	/**
	 * 判断是否不要在这个excel操作中
	 * 
	 * @param
	 * @param field
	 * @param targetId
	 * @return
	 */
	public static boolean isNotUserExcelUserThis(List<String> exclusionsList, Field field, String targetId) {
		boolean boo = true;
		if (field.getAnnotation(ExcelIgnore.class) != null) {
			boo = true;
		} else if (boo && field.getAnnotation(ExcelCollection.class) != null && isUseInThis(field.getAnnotation(ExcelCollection.class).name(), targetId) && (exclusionsList == null || !exclusionsList.contains(field.getAnnotation(ExcelCollection.class).name()))) {
			boo = false;
		} else if (boo && field.getAnnotation(Excel.class) != null && isUseInThis(field.getAnnotation(Excel.class).name(), targetId) && (exclusionsList == null || !exclusionsList.contains(field.getAnnotation(Excel.class).name()))) {
			boo = false;
		} else if (boo && field.getAnnotation(ExcelEntity.class) != null && isUseInThis(field.getAnnotation(ExcelEntity.class).name(), targetId) && (exclusionsList == null || !exclusionsList.contains(field.getAnnotation(ExcelEntity.class).name()))) {
			boo = false;
		}
		return boo;
	}

	/**
	 * 判断是不是使用
	 * 
	 * @param exportName
	 * @param targetId
	 * @return
	 */
	private static boolean isUseInThis(String exportName, String targetId) {
		return targetId == null || exportName.equals("") || exportName.indexOf("_") < 0 || exportName.indexOf(targetId) != -1;
	}

	private static Integer getImageType(String type) {
		if (type.equalsIgnoreCase("JPG") || type.equalsIgnoreCase("JPEG")) {
			return XWPFDocument.PICTURE_TYPE_JPEG;
		}
		if (type.equalsIgnoreCase("GIF")) {
			return XWPFDocument.PICTURE_TYPE_GIF;
		}
		if (type.equalsIgnoreCase("BMP")) {
			return XWPFDocument.PICTURE_TYPE_GIF;
		}
		if (type.equalsIgnoreCase("PNG")) {
			return XWPFDocument.PICTURE_TYPE_PNG;
		}
		return XWPFDocument.PICTURE_TYPE_JPEG;
	}

	/**
	 * 返回流和图片类型
	 * 
	 * @Author JEECG
	 * @date 2013-11-20
	 * @param entity
	 * @return (byte[]) isAndType[0],(Integer)isAndType[1]
	 * @throws Exception
	 */
	public static Object[] getIsAndType(WordImageEntity entity) throws Exception {
		Object[] result = new Object[2];
		String type;
		if (entity.getType().equals(WordImageEntity.URL)) {
			ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
			BufferedImage bufferImg;
			String path = Thread.currentThread().getContextClassLoader().getResource("").toURI().getPath() + entity.getUrl();
			path = path.replace("WEB-INF/classes/", "");
			path = path.replace("file:/", "");
			bufferImg = ImageIO.read(new File(path));
			//update-begin-author:taoYan date:20211203 for: Excel 导出图片的文件带小数点符号 导出报错 https://gitee.com/jeecg/jeecg-boot/issues/I4JNHR
			ImageIO.write(bufferImg, entity.getUrl().substring(entity.getUrl().lastIndexOf(".") + 1, entity.getUrl().length()), byteArrayOut);
			//update-end-author:taoYan date:20211203 for: Excel 导出图片的文件带小数点符号 导出报错 https://gitee.com/jeecg/jeecg-boot/issues/I4JNHR
			result[0] = byteArrayOut.toByteArray();
			type = entity.getUrl().split("/.")[entity.getUrl().split("/.").length - 1];
		} else {
			result[0] = entity.getData();
			type = PoiPublicUtil.getFileExtendName(entity.getData());
		}
		result[1] = getImageType(type);
		return result;
	}

	/**
	 * 获取参数值
	 * 
	 * @param params
	 * @param map
	 * @return
	 */
	@SuppressWarnings("rawtypes")
	public static Object getParamsValue(String params, Object object) throws Exception {
		if (params.indexOf(".") != -1) {
			String[] paramsArr = params.split("\\.");
			return getValueDoWhile(object, paramsArr, 0);
		}
		if (object instanceof Map) {
			return ((Map) object).get(params);
		}
		return getMethod(params, object.getClass()).invoke(object, new Object[] {});
	}

	/**
	 * 解析数据
	 * 
	 * @Author JEECG
	 * @date 2013-11-16
	 * @return
	 */
	public static Object getRealValue(String currentText, Map<String, Object> map) throws Exception {
		String params = "";
		while (currentText.indexOf("{{") != -1) {
			params = currentText.substring(currentText.indexOf("{{") + 2, currentText.indexOf("}}"));
			//update-begin-author:liusq---date:2024-08-07--for: [issues/6925]autopoi通过word模板生成word时：三目、求长、常量、日期转换没起效果
			Object obj = PoiElUtil.eval(params.trim(), map);
			//update-end-author:liusq---date:2024-08-07--for: [issues/6925]autopoi通过word模板生成word时：三目、求长、常量、日期转换没起效果
			// 判断图片或者是集合
			// update-begin-author:taoyan date:20210914 for:autopoi模板导出，赋值的方法建议增加空判断或抛出异常说明。 /issues/3005
			if(obj==null){
				obj = "";
			}
			// update-end-author:taoyan date:20210914 for:autopoi模板导出，赋值的方法建议增加空判断或抛出异常说明。/issues/3005
			if (obj instanceof WordImageEntity || obj instanceof List || obj instanceof ExcelListEntity) {
				return obj;
			} else {
				currentText = currentText.replace("{{" + params + "}}", obj.toString());
			}
		}
		return currentText;
	}

	/**
	 * 通过遍历过去对象值
	 * 
	 * @param object
	 * @param paramsArr
	 * @param index
	 * @return
	 * @throws Exception
	 */
	@SuppressWarnings("rawtypes")
	public static Object getValueDoWhile(Object object, String[] paramsArr, int index) throws Exception {
		if (object == null) {
			return "";
		}
		if (object instanceof WordImageEntity) {
			return object;
		}
		if (object instanceof Map) {
			object = ((Map) object).get(paramsArr[index]);
		} else {
			object = getMethod(paramsArr[index], object.getClass()).invoke(object, new Object[] {});
		}
		return (index == paramsArr.length - 1) ? (object == null ? "" : object) : getValueDoWhile(object, paramsArr, ++index);
	}

	/**
	 * double to String 防止科学计数法
	 * 
	 * @param value
	 * @return
	 */
	public static String doubleToString(Double value) {
		String temp = value.toString();
		if (temp.contains("E")) {
			BigDecimal bigDecimal = new BigDecimal(temp);
			temp = bigDecimal.toPlainString();
		}
		//---update-begin-----autor:scott------date:20191016-------for:excel导入数字类型，去掉后缀.0------
		return ExcelUtil.remove0Suffix(temp);
		//---update-end-----autor:scott------date:20191016-------for:excel导入数字类型，去掉后缀.0------
	}

	/**
	 * 判断是否是数值类型
	 * @param xclass
	 * @return
	 */
	public static boolean isNumber(String xclass){
		if(xclass==null){
			return false;
		}
		String temp = xclass.toLowerCase();
		if(temp.indexOf("int")>=0 || temp.indexOf("double")>=0 || temp.indexOf("decimal")>=0){
			return true;
		}
		return false;
	}
	//update-begin---author:liusq  Date:20211217  for：[LOWCOD-2521]【autopoi】大数据导出方法【全局】----
	/**
	 * 统一 key的获取规则
	 * @param key
	 * @param targetId
	 * @date  2022年1月4号
	 * @return
	 */
	public static String getValueByTargetId(String key, String targetId, String defalut) {
		if (StringUtils.isEmpty(targetId) || key.indexOf("_") < 0) {
			return key;
		}
		String[] arr = key.split(",");
		String[] tempArr;
		for (String str : arr) {
			tempArr = str.split("_");
			if (tempArr == null || tempArr.length < 2) {
				return defalut;
			}
			if (targetId.equals(tempArr[1])) {
				return tempArr[0];
			}
		}
		return defalut;
	}
	//update-end---author:liusq  Date:20211217  for：[LOWCOD-2521]【autopoi】大数据导出方法【全局】----

}
