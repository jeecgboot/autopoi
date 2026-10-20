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
package org.jeecgframework.poi.excel.imports.sax;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.model.SharedStrings;
import org.jeecgframework.poi.excel.entity.enmus.CellValueType;
import org.jeecgframework.poi.excel.entity.sax.SaxReadCellEntity;
import org.jeecgframework.poi.excel.imports.sax.parse.ISaxRowRead;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import com.google.common.collect.Lists;

/**
 * 回调接口
 * 
 * @author JEECG
 * @date 2014年12月29日 下午9:50:09
 */
public class SheetHandler extends DefaultHandler {

	private SharedStrings sst;
	private String lastContents;

	// 当前行
	private int curRow = 0;
	// 当前列
	private int curCol = 0;

	private CellValueType type;

	//update-begin---author:lvdandan ---date:20261020  for：[issues/189]当前单元格的v是否是共享字符串的下标------------
	private boolean sstIndex = false;
	//update-end---author:lvdandan ---date:20261020  for：[issues/189]当前单元格的v是否是共享字符串的下标------------

	private ISaxRowRead read;

	// 存储行记录的容器
	private List<SaxReadCellEntity> rowlist = Lists.newArrayList();

	public SheetHandler(SharedStrings sst, ISaxRowRead rowRead) {
		this.sst = sst;
		this.read = rowRead;
	}

	@Override
	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
		// 置空
		lastContents = "";
		//update-begin---author:lvdandan ---date:20261020  for：[issues/189]SAX导入数字单元格沿用了上一个单元格的类型，被当成共享字符串下标取到别的单元格(如表头)的文本------------
		if ("c".equals(name)) {
			// 每个单元格的类型由它自己的t属性决定，不能沿用上一个单元格的类型
			String t = attributes.getValue("t");
			sstIndex = "s".equals(t);
			if (t == null || "n".equals(t)) {
				// 数字单元格：Excel保存时通常不写t属性
				type = CellValueType.Number;
			} else if ("s".equals(t) || "str".equals(t) || "d".equals(t)) {
				// s：v是共享字符串的下标；str(公式的字符串结果)、d(ISO日期)：v就是文本本身
				type = CellValueType.String;
			} else {
				// inlineStr(文本在t元素里)、b(布尔)、e(错误)等：v 不按数字或共享字符串处理
				type = CellValueType.None;
			}
		}
		//update-end---author:lvdandan ---date:20261020  for：[issues/189]SAX导入数字单元格沿用了上一个单元格的类型，被当成共享字符串下标取到别的单元格(如表头)的文本------------
		// c => 单元格
		if ("c".equals(name)) {
			// 如果下一个元素是 SST 的索引，则将nextIsString标记为true
			String cellType = attributes.getValue("t");
			if ("s".equals(cellType)) {
				type = CellValueType.String;
				return;
			}
			// 日期格式
			cellType = attributes.getValue("s");
			if ("1".equals(cellType)) {
				type = CellValueType.Date;
			} else if ("2".equals(cellType)) {
				type = CellValueType.Number;
			}
		} else if ("t".equals(name)) {// 当元素为t时
			type = CellValueType.TElement;
		}

	}

	@Override
	public void endElement(String uri, String localName, String name) throws SAXException {

		// 根据SST的索引值的到单元格的真正要存储的字符串
		// 这时characters()方法可能会被调用多次
		//update-begin---author:lvdandan ---date:20261020  for：[issues/189]只有共享字符串(t="s")的v才是下标，公式的字符串结果(t="str")的v就是文本本身------------
		if (CellValueType.String.equals(type) && sstIndex) {
		//update-end---author:lvdandan ---date:20261020  for：[issues/189]只有共享字符串(t="s")的v才是下标，公式的字符串结果(t="str")的v就是文本本身------------
			try {
				int idx = Integer.parseInt(lastContents);
				lastContents = sst.getItemAt(idx).toString();
			} catch (Exception e) {

			}
		}
		// t元素也包含字符串
		if (CellValueType.TElement.equals(type)) {
			String value = lastContents.trim();
			rowlist.add(curCol, new SaxReadCellEntity(CellValueType.String, value));
			curCol++;
			type = CellValueType.None;
			// v => 单元格的值，如果单元格是字符串则v标签的值为该字符串在SST中的索引
			// 将单元格内容加入rowlist中，在这之前先去掉字符串前后的空白符
		} else if ("v".equals(name)) {
			String value = lastContents.trim();
			value = value.equals("") ? " " : value;
			if (CellValueType.Date.equals(type)) {
				Date date = DateUtil.getJavaDate(Double.valueOf(value));
				rowlist.add(curCol, new SaxReadCellEntity(CellValueType.Date, date));
			} else if (CellValueType.Number.equals(type)) {
				BigDecimal bd = new BigDecimal(value);
				rowlist.add(curCol, new SaxReadCellEntity(CellValueType.Number, bd));
			} else if (CellValueType.String.equals(type)) {
				rowlist.add(curCol, new SaxReadCellEntity(CellValueType.String, value));
			}
			curCol++;
		} else if (name.equals("row")) {// 如果标签名称为 row ，这说明已到行尾，调用 optRows() 方法
			read.parse(curRow, rowlist);
			rowlist.clear();
			curRow++;
			curCol = 0;
		}

	}

	@Override
	public void characters(char[] ch, int start, int length) throws SAXException {
		// 得到单元格内容的值
		lastContents += new String(ch, start, length);
	}

}
