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
package org.jeecgframework.poi.word.parse;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.jeecgframework.poi.cache.WordCache;
import org.jeecgframework.poi.util.PoiPublicUtil;
import org.jeecgframework.poi.word.entity.MyXWPFDocument;
import org.jeecgframework.poi.word.entity.WordImageEntity;
import org.jeecgframework.poi.word.entity.params.ExcelListEntity;
import org.jeecgframework.poi.word.parse.excel.ExcelEntityParse;
import org.jeecgframework.poi.word.parse.excel.ExcelMapParse;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 解析07版的Word,替换文字,生成表格,生成图片
 * 
 * @author JEECG
 * @date 2013-11-16
 * @version 1.0
 */
@SuppressWarnings({ "unchecked", "rawtypes" })
public class ParseWord07 {

	private static final Logger LOGGER = LoggerFactory.getLogger(ParseWord07.class);

	/**
	 * 添加图片
	 * 
	 * @Author JEECG
	 * @date 2013-11-20
	 * @param obj
	 * @param currentRun
	 * @throws Exception
	 */
	private void addAnImage(WordImageEntity obj, XWPFRun currentRun) throws Exception {
		Object[] isAndType = PoiPublicUtil.getIsAndType(obj);
		String picId;
		try {
			picId = currentRun.getParagraph().getDocument().addPictureData((byte[]) isAndType[0], (Integer) isAndType[1]);
			((MyXWPFDocument) currentRun.getParagraph().getDocument()).createPicture(currentRun, picId, currentRun.getParagraph().getDocument().getNextPicNameNumber((Integer) isAndType[1]), obj.getWidth(), obj.getHeight());

		} catch (Exception e) {
			LOGGER.error(e.getMessage(), e);
		}

	}

	/**
	 * 根据条件改变值
	 * 
	 * @param map
	 * @Author JEECG
	 * @date 2013-11-16
	 */
	private void changeValues(XWPFParagraph paragraph, XWPFRun currentRun, String currentText, List<Integer> runIndex, Map<String, Object> map) throws Exception {
		Object obj = PoiPublicUtil.getRealValue(currentText, map);
		if (obj instanceof WordImageEntity) {// 如果是图片就设置为图片
			currentRun.setText("", 0);
			addAnImage((WordImageEntity) obj, currentRun);
		} else {
			currentText = obj.toString();
			currentRun.setText(currentText, 0);
		}
		for (int k = 0; k < runIndex.size(); k++) {
			paragraph.getRuns().get(runIndex.get(k)).setText("", 0);
		}
		runIndex.clear();
	}

	/**
	 * 判断是不是迭代输出
	 * 
	 * @Author JEECG
	 * @date 2013-11-18
	 * @return
	 * @throws Exception
	 */
	private Object checkThisTableIsNeedIterator(XWPFTableCell cell, Map<String, Object> map) throws Exception {
		String text = cell.getText().trim();
        //begin-------author:liusq------date:20210129-----for:-------poi3升级到4兼容改造工作【重要敏感修改点】--------
		// 判断是不是迭代输出
		if (text != null&&text.startsWith("{{") && text.indexOf("$fe:") != -1) {
			return PoiPublicUtil.getRealValue(text.replace("$fe:", "").trim(), map);
		}
        //end-------author:liusq------date:20210129-----for:-------poi3升级到4兼容改造工作【重要敏感修改点】--------
		return null;
	}

	/**
	 * 解析所有的文本
	 * 
	 * @Author JEECG
	 * @date 2013-11-17
	 * @param paragraphs
	 * @param map
	 */
	private void parseAllParagraphic(List<XWPFParagraph> paragraphs, Map<String, Object> map) throws Exception {
		XWPFParagraph paragraph;
		for (int i = 0; i < paragraphs.size(); i++) {
			paragraph = paragraphs.get(i);
			if (paragraph.getText().indexOf("{{") != -1) {
				parseThisParagraph(paragraph, map);
			}

		}

	}

	/**
	 * 解析这个段落
	 * 
	 * @Author JEECG
	 * @date 2013-11-16
	 * @param paragraph
	 * @param map
	 */
	private void parseThisParagraph(XWPFParagraph paragraph, Map<String, Object> map) throws Exception {
		XWPFRun run;
		XWPFRun currentRun = null;// 拿到的第一个run,用来set值,可以保存格式
		String currentText = "";// 存放当前的text
		String text;
		Boolean isfinde = false;// 判断是不是已经遇到{{
		List<Integer> runIndex = new ArrayList<Integer>();// 存储遇到的run,把他们置空
		for (int i = 0; i < paragraph.getRuns().size(); i++) {
			run = paragraph.getRuns().get(i);
			text = run.getText(0);
			if (StringUtils.isEmpty(text)) {
				continue;
			}// 如果为空或者""这种这继续循环跳过
			if (isfinde) {
				currentText += text;
				if (currentText.indexOf("{{") == -1) {
					isfinde = false;
					runIndex.clear();
				} else {
					runIndex.add(i);
				}
				if (currentText.indexOf("}}") != -1) {
					changeValues(paragraph, currentRun, currentText, runIndex, map);
					currentText = "";
					isfinde = false;
				}
			} else if (text.indexOf("{") >= 0) {// 判断是不是开始
				currentText = text;
				isfinde = true;
				currentRun = run;
			} else {
				currentText = "";
			}
			if (currentText.indexOf("}}") != -1) {
				changeValues(paragraph, currentRun, currentText, runIndex, map);
				isfinde = false;
			}
		}

	}

	private void parseThisRow(List<XWPFTableCell> cells, Map<String, Object> map) throws Exception {
		for (XWPFTableCell cell : cells) {
			parseAllParagraphic(cell.getParagraphs(), map);
		}
	}

	/**
	 * 解析这个表格
	 * 
	 * @Author JEECG
	 * @date 2013-11-17
	 * @param table
	 * @param map
	 */
	private void parseThisTable(XWPFTable table, Map<String, Object> map) throws Exception {
		XWPFTableRow row;
		List<XWPFTableCell> cells;
		Object listobj;
		for (int i = 0; i < table.getNumberOfRows(); i++) {
			row = table.getRow(i);
			cells = row.getTableCells();
			//begin-------author:liusq------date:20210129-----for:-------poi3升级到4兼容改造工作【重要敏感修改点】--------
			listobj = checkThisTableIsNeedIterator(cells.get(0), map);
			if (listobj == null) {
				parseThisRow(cells, map);
			} else if (listobj instanceof ExcelListEntity) {
				new ExcelEntityParse().parseNextRowAndAddRow(table, i, (ExcelListEntity) listobj);
				i = i + ((ExcelListEntity) listobj).getList().size() - 1;//删除之后要往上挪一行,然后加上跳过新建的行数
			} else {
				ExcelMapParse.parseNextRowAndAddRow(table, i, (List) listobj);
				i = i + ((List) listobj).size() - 1;//删除之后要往上挪一行,然后加上跳过新建的行数
			}
			/*if (cells.size() == 1) {
				listobj = checkThisTableIsNeedIterator(cells.get(0), map);
				if (listobj == null) {
					parseThisRow(cells, map);
				} else if (listobj instanceof ExcelListEntity) {
					table.removeRow(i);// 删除这一行
					excelEntityParse.parseNextRowAndAddRow(table, i, (ExcelListEntity) listobj);
				} else {
					table.removeRow(i);// 删除这一行
					ExcelMapParse.parseNextRowAndAddRow(table, i, (List) listobj);
				}
			} else {
				parseThisRow(cells, map);
			}*/
			//end-------author:liusq------date:20210129-----for:-------poi3升级到4兼容改造工作【重要敏感修改点】--------
		}
	}

	/**
	 * 解析07版的Word并且进行赋值
	 * 
	 * @Author JEECG
	 * @date 2013-11-16
	 * @return
	 * @throws Exception
	 */
	public XWPFDocument parseWord(String url, Map<String, Object> map) throws Exception {
		MyXWPFDocument doc = WordCache.getXWPFDocumen(url);
		parseWordSetValue(doc, map);
		return doc;
	}

	/**
	 * 解析07版的Word并且进行赋值
	 * 
	 * @Author JEECG
	 * @date 2013-11-16
	 * @return
	 * @throws Exception
	 */
	public void parseWord(XWPFDocument document, Map<String, Object> map) throws Exception {
		parseWordSetValue((MyXWPFDocument) document, map);
	}

	private void parseWordSetValue(MyXWPFDocument doc, Map<String, Object> map) throws Exception {
		// 第一步解析文档
		parseAllParagraphic(doc.getParagraphs(), map);
		// 第二步解析页眉,页脚
		parseHeaderAndFoot(doc, map);
		// 第三步解析所有表格
		XWPFTable table;
		Iterator<XWPFTable> itTable = doc.getTablesIterator();
		while (itTable.hasNext()) {
			table = itTable.next();
			if (table.getText().indexOf("{{") != -1) {
				parseThisTable(table, map);
			}
		}

	}

	/**
	 * 解析页眉和页脚
	 * 
	 * @param doc
	 * @param map
	 * @throws Exception
	 */
	private void parseHeaderAndFoot(MyXWPFDocument doc, Map<String, Object> map) throws Exception {
		List<XWPFHeader> headerList = doc.getHeaderList();
		for (XWPFHeader xwpfHeader : headerList) {
			for (int i = 0; i < xwpfHeader.getListParagraph().size(); i++) {
				parseThisParagraph(xwpfHeader.getListParagraph().get(i), map);
			}
		}
		List<XWPFFooter> footerList = doc.getFooterList();
		for (XWPFFooter xwpfFooter : footerList) {
			for (int i = 0; i < xwpfFooter.getListParagraph().size(); i++) {
				parseThisParagraph(xwpfFooter.getListParagraph().get(i), map);
			}
		}

	}
}
