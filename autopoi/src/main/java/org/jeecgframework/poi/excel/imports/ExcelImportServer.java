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
package org.jeecgframework.poi.excel.imports;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jeecgframework.core.util.ApplicationContextUtil;
import org.jeecgframework.poi.excel.annotation.ExcelTarget;
import org.jeecgframework.poi.excel.entity.ImportParams;
import org.jeecgframework.poi.excel.entity.params.ExcelCollectionParams;
import org.jeecgframework.poi.excel.entity.params.ExcelImportEntity;
import org.jeecgframework.poi.excel.entity.result.ExcelImportResult;
import org.jeecgframework.poi.excel.entity.result.ExcelVerifyHanlderResult;
import org.jeecgframework.poi.excel.imports.base.ImportBaseService;
import org.jeecgframework.poi.excel.imports.base.ImportFileServiceI;
import org.jeecgframework.poi.excel.imports.verifys.VerifyHandlerServer;
import org.jeecgframework.poi.exception.excel.ExcelImportException;
import org.jeecgframework.poi.exception.excel.enums.ExcelImportEnum;
import org.jeecgframework.poi.util.ExcelUtil;
import org.jeecgframework.poi.util.PoiPublicUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Excel 导入服务
 * 
 * @author JEECG
 * @date 2014年6月26日 下午9:20:51
 */
@SuppressWarnings({ "rawtypes", "unchecked", "hiding" })
public class ExcelImportServer extends ImportBaseService {

	private final static Logger LOGGER = LoggerFactory.getLogger(ExcelImportServer.class);

	private CellValueServer cellValueServer;

	private VerifyHandlerServer verifyHandlerServer;

	private boolean verfiyFail = false;
	//仅允许字母数字字符的正则表达式
	private static final Pattern lettersAndNumbersPattern = Pattern.compile("^[a-zA-Z0-9]+$") ;
	/**
	 * 异常数据styler
	 */
	private CellStyle errorCellStyle;

	public ExcelImportServer() {
		this.cellValueServer = new CellValueServer();
		this.verifyHandlerServer = new VerifyHandlerServer();
	}

	/***
	 * 向List里面继续添加元素
	 * 
	 * @param object
	 * @param param
	 * @param row
	 * @param titlemap
	 * @param targetId
	 * @param pictures
	 * @param params
	 */
	private void addListContinue(Object object, ExcelCollectionParams param, Row row, Map<Integer, String> titlemap, String targetId, Map<String, PictureData> pictures, ImportParams params) throws Exception {
		Collection collection = (Collection) PoiPublicUtil.getMethod(param.getName(), object.getClass()).invoke(object, new Object[] {});
		Object entity = PoiPublicUtil.createObject(param.getType(), targetId);
		String picId;
		boolean isUsed = false;// 是否需要加上这个对象
		for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
			Cell cell = row.getCell(i);
			String titleString = (String) titlemap.get(i);
			if (param.getExcelParams().containsKey(titleString)) {
				if (param.getExcelParams().get(titleString).getType() == 2) {
					picId = row.getRowNum() + "_" + i;
                    //update-begin---author:chenrui ---date:20240402  for：[issue/#6025/#6040]子表图片导入报错------------
					saveImage(entity, picId, param.getExcelParams(), titleString, pictures, params);
                    //update-end---author:chenrui ---date:20240402  for：[issue/#6025/#6040]子表图片导入报错------------
				} else {
					saveFieldValue(params, entity, cell, param.getExcelParams(), titleString, row);
				}
				isUsed = true;
			}
		}
		if (isUsed) {
			collection.add(entity);
		}
	}

	/**
	 * 获取key的值,针对不同类型获取不同的值
	 * 
	 * @Author JEECG
	 * @date 2013-11-21
	 * @param cell
	 * @return
	 */
	private String getKeyValue(Cell cell) {
		if(cell==null){
			return null;
		}
		Object obj = null;
		switch (cell.getCellTypeEnum()) {
		case STRING:
			obj = cell.getStringCellValue();
			break;
		case BOOLEAN:
			obj = cell.getBooleanCellValue();
			break;
		case NUMERIC:
			obj = cell.getNumericCellValue();
			break;
		case FORMULA:
			obj = cell.getCellFormula();
			break;
		}
		return obj == null ? null : obj.toString().trim();
	}

	/**
	 * 获取保存的真实路径
	 * 
	 * @param excelImportEntity
	 * @param object
	 * @return
	 * @throws Exception
	 */
	private String getSaveUrl(ExcelImportEntity excelImportEntity, Object object) throws Exception {
		String url = "";
		if (excelImportEntity.getSaveUrl().equals("upload")) {
			if (excelImportEntity.getMethods() != null && excelImportEntity.getMethods().size() > 0) {
				object = getFieldBySomeMethod(excelImportEntity.getMethods(), object);
			}
			url = object.getClass().getName().split("\\.")[object.getClass().getName().split("\\.").length - 1];
			return excelImportEntity.getSaveUrl() + "/" + url.substring(0, url.lastIndexOf("Entity"));
		}
		return excelImportEntity.getSaveUrl();
	}
	//update-begin--Author:xuelin  Date:20171205 for：TASK #2098 【excel问题】 Online 一对多导入失败--------------------
	private <T> List<T> importExcel(Collection<T> result, Sheet sheet, Class<?> pojoClass, ImportParams params, Map<String, PictureData> pictures) throws Exception {
		List collection = new ArrayList();
		Map<String, ExcelImportEntity> excelParams = new HashMap<String, ExcelImportEntity>();
		List<ExcelCollectionParams> excelCollection = new ArrayList<ExcelCollectionParams>();
		String targetId = null;
		if (!Map.class.equals(pojoClass)) {
			Field fileds[] = PoiPublicUtil.getClassFields(pojoClass);
			ExcelTarget etarget = pojoClass.getAnnotation(ExcelTarget.class);
			if (etarget != null) {
				targetId = etarget.value();
			}
			getAllExcelField(targetId, fileds, excelParams, excelCollection, pojoClass, null);
		}
		ignoreHeaderHandler(excelParams, params);
		Iterator<Row> rows = sheet.rowIterator();
		Map<Integer, String> titlemap = getTitleMap(sheet, rows, params, excelCollection);
		//update-begin-author:liusq date:20220310 for:[issues/I4PU45]@excel里面新增属性fixedIndex
		Set<String> keys = excelParams.keySet();
		for (String key : keys) {
			if (key.startsWith("FIXED_")) {
				String[] arr = key.split("_");
				titlemap.put(Integer.parseInt(arr[1]), key);
			}
		}
		//update-end-author:liusq date:20220310 for:[issues/I4PU45]@excel里面新增属性fixedIndex
		Set<Integer> columnIndexSet = titlemap.keySet();
        Integer maxColumnIndex = Collections.max(columnIndexSet);
        Integer minColumnIndex = Collections.min(columnIndexSet);
		Row row = null;
		//跳过表头和标题行
		for (int j = 0; j < params.getTitleRows() + params.getHeadRows(); j++) {
			row = rows.next();
		}
		Object object = null;
		String picId;
		while (rows.hasNext() && (row == null || sheet.getLastRowNum() - row.getRowNum() > params.getLastOfInvalidRow())) {
			row = rows.next();
			//update-begin--Author:xuelin  Date:20171017 for：TASK #2373 【bug】表改造问题，导致 3.7.1批量导入用户bug-导入不成功--------------------
			// 判断是集合元素还是不是集合元素,如果是就继续加入这个集合,不是就创建新的对象
			//update-begin--Author:xuelin  Date:20171206 for：TASK #2451 【excel导出bug】online 一对多导入成功， 但是现在代码生成后的一对多online导入有问题了
			Cell keyIndexCell = row.getCell(params.getKeyIndex());
			if (excelCollection.size()>0 && StringUtils.isEmpty(getKeyValue(keyIndexCell)) && object != null && !Map.class.equals(pojoClass)) {
				//update-end--Author:xuelin  Date:20171206 for：TASK #2451 【excel导出bug】online 一对多导入成功， 但是现在代码生成后的一对多online导入有问题了
				for (ExcelCollectionParams param : excelCollection) {
					addListContinue(object, param, row, titlemap, targetId, pictures, params);
				}
				
			} else {			
				object = PoiPublicUtil.createObject(pojoClass, targetId);
				try {
                    //update-begin-author:taoyan date:20200303 for:导入图片
				    int firstCellNum = row.getFirstCellNum();
				    if(firstCellNum>minColumnIndex){
                        firstCellNum = minColumnIndex;
                    }
                    int lastCellNum = row.getLastCellNum();
                    if(lastCellNum<maxColumnIndex+1){
                        lastCellNum = maxColumnIndex+1;
                    }
					//update-begin---author:chenrui ---date:20240306  for：[QQYUN-8394]Excel导入时空行校验问题------------
					int noneCellNum = 0;
					for (int i = firstCellNum, le = lastCellNum; i < le; i++) {
						Cell cell = row.getCell(i);
						String titleString = (String) titlemap.get(i);
						if (excelParams.containsKey(titleString) || Map.class.equals(pojoClass)) {
							if (excelParams.get(titleString) != null && excelParams.get(titleString).getType() == 2) {
								picId = row.getRowNum() + "_" + i;
								saveImage(object, picId, excelParams, titleString, pictures, params);
							} else {
								if(params.getImageList()!=null && params.getImageList().contains(titleString)){
									if (pictures != null) {
										picId = row.getRowNum() + "_" + i;
										PictureData image = pictures.get(picId);
										if(image!=null){
											byte[] data = image.getData();
											params.getDataHanlder().setMapValue((Map) object, titleString, data);
										}
									}
								}else{
									Object value = saveFieldValue(params, object, cell, excelParams, titleString, row);
									if(null == value){
										noneCellNum++;
									}
								}
                        //update-end-author:taoyan date:20200303 for:导入图片
							}
						}
					}

					for (ExcelCollectionParams param : excelCollection) {
						addListContinue(object, param, row, titlemap, targetId, pictures, params);
					}
					//update-begin-author:taoyan date:20210526 for:autopoi导入excel 如果单元格被设置边框，即使没有内容也会被当做是一条数据导入 #2484
                    if (isNotNullObject(pojoClass, object) && noneCellNum < (lastCellNum - firstCellNum)) {
					//update-end---author:chenrui ---date:20240306  for：[QQYUN-8394]Excel导入时空行校验问题------------
						collection.add(object);
					}
					//update-end-author:taoyan date:20210526 for:autopoi导入excel 如果单元格被设置边框，即使没有内容也会被当做是一条数据导入 #2484
				} catch (ExcelImportException e) {
					if (!e.getType().equals(ExcelImportEnum.VERIFY_ERROR)) {
						throw new ExcelImportException(e.getType(), e);
					}
				}
			}
			//update-end--Author:xuelin  Date:20171017 for：TASK #2373 【bug】表改造问题，导致 3.7.1批量导入用户bug-导入不成功--------------------
		}
		return collection;
	}

	/**
	 * 判断当前对象不是空
	 * @param pojoClass
	 * @param object
	 * @return
	 */
	private boolean isNotNullObject(Class pojoClass, Object object){
		try {
			Method method = pojoClass.getMethod("isNullObject");
			if(method!=null){
				Object flag = method.invoke(object);
				if(flag!=null && true == Boolean.parseBoolean(flag.toString())){
					return false;
				}
			}
		} catch (NoSuchMethodException e) {
			LOGGER.debug("未定义方法 isNullObject");
		} catch (IllegalAccessException e) {
			LOGGER.warn("没有权限访问该方法 isNullObject");
		} catch (InvocationTargetException e) {
			LOGGER.warn("方法调用失败 isNullObject");
		}
		return true;
	}

	/**
	 * 获取忽略的表头信息
	 * @param excelParams
	 * @param params
	 */
	private void ignoreHeaderHandler(Map<String, ExcelImportEntity> excelParams,ImportParams params){
		List<String> ignoreList = new ArrayList<>();
		for(String key:excelParams.keySet()){
			String temp = excelParams.get(key).getGroupName();
			if(temp!=null && temp.length()>0){
				ignoreList.add(temp);
			}
		}
		params.setIgnoreHeaderList(ignoreList);
	}

	/**
	 * 获取表格字段列名对应信息
	 * 
	 * @param rows
	 * @param params
	 * @param excelCollection
	 * @return
	 */
	private Map<Integer, String> getTitleMap(Sheet sheet, Iterator<Row> rows, ImportParams params, List<ExcelCollectionParams> excelCollection) throws Exception {
		Map<Integer, String> titlemap = new HashMap<Integer, String>();
		Iterator<Cell> cellTitle = null;
		String collectionName = null;
		ExcelCollectionParams collectionParams = null;
		Row headRow = null;
		int headBegin = params.getTitleRows();
		//update_begin-author:taoyan date:2020622 for：当文件行数小于代码里设置的TitleRows时headRow一直为空就会出现死循环
		int allRowNum = sheet.getPhysicalNumberOfRows();
		//找到首行表头，每个sheet都必须至少有一行表头
		while(headRow == null && headBegin < allRowNum){
			headRow = sheet.getRow(headBegin++);
		}
		if(headRow==null){
			throw new Exception("不识别该文件");
		}
		//update-end-author:taoyan date:2020622 for：当文件行数小于代码里设置的TitleRows时headRow一直为空就会出现死循环

		//设置表头行数
		if (ExcelUtil.isMergedRegion(sheet, headRow.getRowNum(), 0)) {
			params.setHeadRows(2);
		}else{
			params.setHeadRows(1);
		}
		cellTitle = headRow.cellIterator();
		while (cellTitle.hasNext()) {
			Cell cell = cellTitle.next();
			String value = getKeyValue(cell);
			if (StringUtils.isNotEmpty(value)) {
				titlemap.put(cell.getColumnIndex(), value);//加入表头列表
			}
		}
		
		//多行表头
		for (int j = headBegin; j < headBegin + params.getHeadRows()-1; j++) {
			headRow = sheet.getRow(j);
			cellTitle = headRow.cellIterator();
			while (cellTitle.hasNext()) {
				Cell cell = cellTitle.next();
				String value = getKeyValue(cell);
				if (StringUtils.isNotEmpty(value)) {
					int columnIndex = cell.getColumnIndex();
					//当前cell的上一行是否为合并单元格
					if(ExcelUtil.isMergedRegion(sheet, cell.getRowIndex()-1, columnIndex)){
						collectionName = ExcelUtil.getMergedRegionValue(sheet, cell.getRowIndex()-1, columnIndex);
						if(params.isIgnoreHeader(collectionName)){
							titlemap.put(cell.getColumnIndex(), value);
						}else{
							titlemap.put(cell.getColumnIndex(), collectionName + "_" + value);
						}
					}else{
						//update-begin-author:taoyan date:20220112 for: JT640 【online】导入 无论一对一还是一对多 如果子表只有一个字段 则子表无数据
						// 上一行不是合并的情况下另有一种特殊的场景： 如果当前单元格和上面的单元格同一列 即子表字段只有一个 所以标题没有出现跨列
						String prefixTitle = titlemap.get(cell.getColumnIndex());
						if(prefixTitle!=null && !"".equals(prefixTitle)){
							titlemap.put(cell.getColumnIndex(), prefixTitle + "_" +value);
						}else{
							titlemap.put(cell.getColumnIndex(), value);
						}
						//update-end-author:taoyan date:20220112 for: JT640 【online】导入 无论一对一还是一对多 如果子表只有一个字段 则子表无数据
					}
					/*int i = cell.getColumnIndex();
					// 用以支持重名导入
					if (titlemap.containsKey(i)) {
						collectionName = titlemap.get(i);
						collectionParams = getCollectionParams(excelCollection, collectionName);
						titlemap.put(i, collectionName + "_" + value);
					} else if (StringUtils.isNotEmpty(collectionName) && collectionParams.getExcelParams().containsKey(collectionName + "_" + value)) {
						titlemap.put(i, collectionName + "_" + value);
					} else {
						collectionName = null;
						collectionParams = null;
					}
					if (StringUtils.isEmpty(collectionName)) {
						titlemap.put(i, value);
					}*/
				}
			}
		}
		return titlemap;
	}
	//update-end--Author:xuelin  Date:20171205 for：TASK #2098 【excel问题】 Online 一对多导入失败--------------------
	/**
	 * 获取这个名称对应的集合信息
	 * 
	 * @param excelCollection
	 * @param collectionName
	 * @return
	 */
	private ExcelCollectionParams getCollectionParams(List<ExcelCollectionParams> excelCollection, String collectionName) {
		for (ExcelCollectionParams excelCollectionParams : excelCollection) {
			if (collectionName.equals(excelCollectionParams.getExcelName())) {
				return excelCollectionParams;
			}
		}
		return null;
	}

	/**
	 * Excel 导入 field 字段类型 Integer,Long,Double,Date,String,Boolean
	 * 
	 * @param inputstream
	 * @param pojoClass
	 * @param params
	 * @return
	 * @throws Exception
	 */
	public ExcelImportResult importExcelByIs(InputStream inputstream, Class<?> pojoClass, ImportParams params) throws Exception {
		if (LOGGER.isDebugEnabled()) {
			LOGGER.debug("Excel import start ,class is {}", pojoClass);
		}
		List<T> result = new ArrayList<T>();
		Workbook book = null;
		boolean isXSSFWorkbook = false;
		//update-begin---author:chenrui ---date:20240403  for：[issue/#5987]嵌入单元格图片无法导入------------
		// 复制输入流,防止在读取嵌入图片时流为空
        ByteArrayOutputStream inCopy = new ByteArrayOutputStream();
        IOUtils.copy(inputstream, inCopy);
        inputstream = new ByteArrayInputStream(inCopy.toByteArray());
        if (!(inputstream.markSupported())) {
            inputstream = new PushbackInputStream(inputstream, 8);
        }
		//update-end---author:chenrui ---date:20240403  for：[issue/#5987]嵌入单元格图片无法导入------------
		//begin-------author:liusq------date:20210129-----for:-------poi3升级到4兼容改造工作【重要敏感修改点】--------
		//------poi4.x begin----
//		FileMagic fm = FileMagic.valueOf(FileMagic.prepareToCheckMagic(inputstream));
//		if(FileMagic.OLE2 == fm){
//			isXSSFWorkbook=false;
//		}
		book = WorkbookFactory.create(inputstream);
		if(book instanceof XSSFWorkbook){
			isXSSFWorkbook=true;
		}
		LOGGER.info("  >>>  poi3升级到4.0兼容改造工作, isXSSFWorkbook = " +isXSSFWorkbook);
		//end-------author:liusq------date:20210129-----for:-------poi3升级到4兼容改造工作【重要敏感修改点】--------

		//begin-------author:liusq------date:20210313-----for:-------多sheet导入改造点--------
		//获取导入文本的sheet数
		//update-begin-author:taoyan date:20211210 for:https://gitee.com/jeecg/jeecg-boot/issues/I45C32 导入空白sheet报错
		if(params.getSheetNum()==0){
			int sheetNum = book.getNumberOfSheets();
			if(sheetNum>0){
				params.setSheetNum(sheetNum);
			}
		}
		//update-end-author:taoyan date:20211210 for:https://gitee.com/jeecg/jeecg-boot/issues/I45C32 导入空白sheet报错
		//end-------author:liusq------date:20210313-----for:-------多sheet导入改造点--------
		createErrorCellStyle(book);
		Map<String, PictureData> pictures;
		// 获取指定的sheet名称
		String sheetName = params.getSheetName();

		//update-begin-author:liusq date:20220609 for:issues/I57UPC excel导入 ImportParams 中没有startSheetIndex参数
		for (int i = params.getStartSheetIndex(); i < params.getStartSheetIndex()
				+ params.getSheetNum(); i++) {
		//update-end-author:liusq date:20220609 for:issues/I57UPC excel导入 ImportParams 中没有startSheetIndex参数

			//update-begin-author:taoyan date:2023-3-4 for: 导入数据支持指定sheet名称
			if(sheetName!=null && !"".equals(sheetName)){
				Sheet tempSheet = book.getSheetAt(i);
				if(!sheetName.equals(tempSheet.getSheetName())){
					continue;
				}
			}
			//update-end-author:taoyan date:2023-3-4 for: 导入数据支持指定sheet名称
				
			if (LOGGER.isDebugEnabled()) {
				LOGGER.debug(" start to read excel by is ,startTime is {}", System.currentTimeMillis());
			}
			if (isXSSFWorkbook) {
				pictures = PoiPublicUtil.getSheetPictrues07((XSSFSheet) book.getSheetAt(i), (XSSFWorkbook) book);
				//update-begin---author:chenrui ---date:20240403  for：[issue/#5987]嵌入单元格图片无法导入------------
                Map<String, PictureData> cellImages = PoiPublicUtil.getCellImages(book.getSheetAt(i), inCopy, book);
                if (!cellImages.isEmpty()) {
                    pictures.putAll(cellImages);
                }
				//update-end---author:chenrui ---date:20240403  for：[issue/#5987]嵌入单元格图片无法导入------------
			} else {
				pictures = PoiPublicUtil.getSheetPictrues03((HSSFSheet) book.getSheetAt(i), (HSSFWorkbook) book);
			}
			if (LOGGER.isDebugEnabled()) {
				LOGGER.debug(" end to read excel by is ,endTime is {}", new Date().getTime());
			}
			result.addAll(importExcel(result, book.getSheetAt(i), pojoClass, params, pictures));
			if (LOGGER.isDebugEnabled()) {
				LOGGER.debug(" end to read excel list by pos ,endTime is {}", new Date().getTime());
			}
		}
		if (params.isNeedSave()) {
			saveThisExcel(params, pojoClass, isXSSFWorkbook, book);
		}
		return new ExcelImportResult(result, verfiyFail, book);
	}
	/**
	 *
	 * @param is
	 * @return
	 * @throws IOException
	 */
	public static byte[] getBytes(InputStream is) throws IOException {
		ByteArrayOutputStream buffer = new ByteArrayOutputStream();

		int len;
		byte[] data = new byte[100000];
		while ((len = is.read(data, 0, data.length)) != -1) {
			buffer.write(data, 0, len);
		}

		buffer.flush();
		return buffer.toByteArray();
	}

	/**
	 * 保存字段值(获取值,校验值,追加错误信息)
	 * 
	 * @param params
	 * @param object
	 * @param cell
	 * @param excelParams
	 * @param titleString
	 * @param row
	 * @throws Exception
	 * @return
	 */
	private Object saveFieldValue(ImportParams params, Object object, Cell cell, Map<String, ExcelImportEntity> excelParams, String titleString, Row row) throws Exception {
		Object value = cellValueServer.getValue(params.getDataHanlder(), object, cell, excelParams, titleString);
		if (object instanceof Map) {
			if (params.getDataHanlder() != null) {
				params.getDataHanlder().setMapValue((Map) object, titleString, value);
			} else {
				((Map) object).put(titleString, value);
			}
		} else {
			ExcelVerifyHanlderResult verifyResult = verifyHandlerServer.verifyData(object, value, titleString, excelParams.get(titleString).getVerify(), params.getVerifyHanlder());
			if (verifyResult.isSuccess()) {
				setValues(excelParams.get(titleString), object, value);
			} else {
				Cell errorCell = row.createCell(row.getLastCellNum());
				errorCell.setCellValue(verifyResult.getMsg());
				errorCell.setCellStyle(errorCellStyle);
				verfiyFail = true;
				throw new ExcelImportException(ExcelImportEnum.VERIFY_ERROR);
			}
		}
		return value;
	}

	/**
	 * 
	 * @param object
	 * @param picId
	 * @param excelParams
	 * @param titleString
	 * @param pictures
	 * @param params
	 * @throws Exception
	 */
	private void saveImage(Object object, String picId, Map<String, ExcelImportEntity> excelParams, String titleString, Map<String, PictureData> pictures, ImportParams params) throws Exception {
		if (pictures == null || pictures.get(picId)==null) {
			return;
		}
		PictureData image = pictures.get(picId);
		byte[] data = image.getData();
		String fileName = "pic" + Math.round(Math.random() * 100000000000L);
		fileName += "." + PoiPublicUtil.getFileExtendName(data);
		//update-beign-author:taoyan date:20200302 for:【多任务】online 专项集中问题 LOWCOD-159
		int saveType = excelParams.get(titleString).getSaveType();
		if ( saveType == 1) {
			String path = PoiPublicUtil.getWebRootPath(getSaveUrl(excelParams.get(titleString), object));
			File savefile = new File(path);
			if (!savefile.exists()) {
				savefile.mkdirs();
			}
			savefile = new File(path + "/" + fileName);
			FileOutputStream fos = new FileOutputStream(savefile);
			fos.write(data);
			fos.close();
			setValues(excelParams.get(titleString), object, getSaveUrl(excelParams.get(titleString), object) + "/" + fileName);
		} else if(saveType==2) {
			setValues(excelParams.get(titleString), object, data);
		} else {
			ImportFileServiceI importFileService = null;
			try {
				importFileService = ApplicationContextUtil.getContext().getBean(ImportFileServiceI.class);
			} catch (Exception e) {
				System.err.println(e.getMessage());
			}
			if(importFileService!=null){
				//update-beign-author:liusq date:20230411 for:【issue/4415】autopoi-web 导入图片字段时无法指定保存路径
				String saveUrl = excelParams.get(titleString).getSaveUrl();
				String dbPath;
				if(StringUtils.isNotBlank(saveUrl)){
					LOGGER.debug("图片保存路径saveUrl = "+saveUrl);
					Matcher matcher = lettersAndNumbersPattern.matcher(saveUrl);
					if(!matcher.matches()){
						LOGGER.warn("图片保存路径格式错误，只能设置字母和数字的组合!");
						dbPath = importFileService.doUpload(data);
					}else{
						dbPath = importFileService.doUpload(data,saveUrl);
					}
				}else{
					dbPath = importFileService.doUpload(data);
				}
				//update-end-author:liusq date:20230411 for:【issue/4415】autopoi-web 导入图片字段时无法指定保存路径
				setValues(excelParams.get(titleString), object, dbPath);
			}
		}
		//update-end-author:taoyan date:20200302 for:【多任务】online 专项集中问题 LOWCOD-159
	}

	private void createErrorCellStyle(Workbook workbook) {
		errorCellStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setColor(Font.COLOR_RED);
		errorCellStyle.setFont(font);
	}

}
