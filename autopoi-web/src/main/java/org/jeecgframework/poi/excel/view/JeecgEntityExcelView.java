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
package org.jeecgframework.poi.excel.view;

import java.util.Collection;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.def.NormalExcelConstants;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.export.ExcelExportServer;
import org.springframework.stereotype.Controller;

/**
 * Entity 实体数据对象导出
 * @Author JEECG
 * 
 */
@SuppressWarnings("unchecked")
@Controller(NormalExcelConstants.JEECG_ENTITY_EXCEL_VIEW)
public class JeecgEntityExcelView extends MiniAbstractExcelView {

	public JeecgEntityExcelView() {
		super();
	}

	@Override
	protected void renderMergedOutputModel(Map<String, Object> model, HttpServletRequest request, HttpServletResponse response) throws Exception {
		String codedFileName = "临时文件";
		Workbook workbook = null;
		//---update-end-----autor:scott------date:20191016-------for:导出字段支持自定义--------
		String[] exportFields = null;

		Object exportFieldStr =  model.get(NormalExcelConstants.EXPORT_FIELDS);
		if(exportFieldStr!=null && exportFieldStr!=""){
			exportFields = exportFieldStr.toString().split(",");
		}
        //---update-end-----autor:scott------date:20191016-------for:导出字段支持自定义--------
		if (model.containsKey(NormalExcelConstants.MAP_LIST)) {
			List<Map<String, Object>> list = (List<Map<String, Object>>) model.get(NormalExcelConstants.MAP_LIST);
			if (list.size() == 0) {
				throw new RuntimeException("MAP_LIST IS NULL");
			}
			workbook = ExcelExportUtil.exportExcel((ExportParams) list.get(0).get(NormalExcelConstants.PARAMS), (Class<?>) list.get(0).get(NormalExcelConstants.CLASS), (Collection<?>) list.get(0).get(NormalExcelConstants.DATA_LIST),exportFields);
			for (int i = 1; i < list.size(); i++) {
				new ExcelExportServer().createSheet(workbook, (ExportParams) list.get(i).get(NormalExcelConstants.PARAMS), (Class<?>) list.get(i).get(NormalExcelConstants.CLASS), (Collection<?>) list.get(i).get(NormalExcelConstants.DATA_LIST),exportFields);
			}
		} else {
			workbook = ExcelExportUtil.exportExcel((ExportParams) model.get(NormalExcelConstants.PARAMS), (Class<?>) model.get(NormalExcelConstants.CLASS), (Collection<?>) model.get(NormalExcelConstants.DATA_LIST),exportFields);
		}
		if (model.containsKey(NormalExcelConstants.FILE_NAME)) {
			codedFileName = (String) model.get(NormalExcelConstants.FILE_NAME);
		}
		if (workbook instanceof HSSFWorkbook) {
			codedFileName += HSSF;
		} else {
			codedFileName += XSSF;
		}
		if (isIE(request)) {
			codedFileName = java.net.URLEncoder.encode(codedFileName, "UTF8");
		} else {
			codedFileName = new String(codedFileName.getBytes("UTF-8"), "ISO-8859-1");
		}
		response.setHeader("content-disposition", "attachment;filename=" + codedFileName);
		ServletOutputStream out = response.getOutputStream();
		workbook.write(out);
		out.flush();
	}
}
