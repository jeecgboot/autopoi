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
package org.jeecgframework.poi.word.entity;

import java.io.InputStream;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 扩充document,修复图片插入失败问题问题
 * 
 * @author JEECG
 * @date 2013-11-20
 * @version 1.0
 */
public class MyXWPFDocument extends XWPFDocument {

	private static final Logger LOGGER = LoggerFactory.getLogger(MyXWPFDocument.class);

	private static String PICXML = "" + "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" + "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" + "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" + "         <pic:nvPicPr>" + "            <pic:cNvPr id=\"%s\" name=\"Generated\"/>" + "            <pic:cNvPicPr/>" + "         </pic:nvPicPr>" + "         <pic:blipFill>" + "            <a:blip r:embed=\"%s\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" + "            <a:stretch>" + "               <a:fillRect/>" + "            </a:stretch>" + "         </pic:blipFill>" + "         <pic:spPr>" + "            <a:xfrm>"
			+ "               <a:off x=\"0\" y=\"0\"/>" + "               <a:ext cx=\"%s\" cy=\"%s\"/>" + "            </a:xfrm>" + "            <a:prstGeom prst=\"rect\">" + "               <a:avLst/>" + "            </a:prstGeom>" + "         </pic:spPr>" + "      </pic:pic>" + "   </a:graphicData>" + "</a:graphic>";

	public MyXWPFDocument() {
		super();
	}

	public MyXWPFDocument(InputStream in) throws Exception {
		super(in);
	}

	public MyXWPFDocument(OPCPackage opcPackage) throws Exception {
		super(opcPackage);
	}

	public void createPicture(String blipId, int id, int width, int height) {
		final int EMU = 9525;
		width *= EMU;
		height *= EMU;
		CTInline inline = createParagraph().createRun().getCTR().addNewDrawing().addNewInline();
		String picXml = String.format(PICXML, id, blipId, width, height);
		XmlToken xmlToken = null;
		try {
			xmlToken = XmlToken.Factory.parse(picXml);
		} catch (XmlException xe) {
			LOGGER.error(xe.getMessage(), xe.fillInStackTrace());
		}
		inline.set(xmlToken);

		inline.setDistT(0);
		inline.setDistB(0);
		inline.setDistL(0);
		inline.setDistR(0);

		CTPositiveSize2D extent = inline.addNewExtent();
		extent.setCx(width);
		extent.setCy(height);

		CTNonVisualDrawingProps docPr = inline.addNewDocPr();
		docPr.setId(id);
		docPr.setName("Picture " + id);
		docPr.setDescr("Generated");
	}

	public void createPicture(XWPFRun run, String blipId, int id, int width, int height) {
		final int EMU = 9525;
		width *= EMU;
		height *= EMU;
		CTInline inline = run.getCTR().addNewDrawing().addNewInline();
		String picXml = String.format(PICXML, id, blipId, width, height);
		XmlToken xmlToken = null;
		try {
			xmlToken = XmlToken.Factory.parse(picXml);
		} catch (XmlException xe) {
			LOGGER.error(xe.getMessage(), xe.fillInStackTrace());
		}
		inline.set(xmlToken);

		inline.setDistT(0);
		inline.setDistB(0);
		inline.setDistL(0);
		inline.setDistR(0);

		CTPositiveSize2D extent = inline.addNewExtent();
		extent.setCx(width);
		extent.setCy(height);

		CTNonVisualDrawingProps docPr = inline.addNewDocPr();
		docPr.setId(id);
		docPr.setName("Picture " + id);
		docPr.setDescr("Generated");
	}

}
