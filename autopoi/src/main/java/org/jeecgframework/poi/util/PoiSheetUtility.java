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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * 国外高手的,不过也不好,慎用,效率不行 Helper functions to aid in the management of sheets
 */
public class PoiSheetUtility extends Object {

	/**
	 * Given a sheet, this method deletes a column from a sheet and moves all
	 * the columns to the right of it to the left one cell.
	 * 
	 * Note, this method will not update any formula references.
	 * 
	 * @param sheet
	 * @param column
	 */
	public static void deleteColumn(Sheet sheet, int columnToDelete) {
		int maxColumn = 0;
		for (int r = 0; r < sheet.getLastRowNum() + 1; r++) {
			Row row = sheet.getRow(r);

			// if no row exists here; then nothing to do; next!
			if (row == null)
				continue;

			// if the row doesn't have this many columns then we are good; next!
			int lastColumn = row.getLastCellNum();
			if (lastColumn > maxColumn)
				maxColumn = lastColumn;

			if (lastColumn < columnToDelete)
				continue;

			for (int x = columnToDelete + 1; x < lastColumn + 1; x++) {
				Cell oldCell = row.getCell(x - 1);
				if (oldCell != null)
					row.removeCell(oldCell);

				Cell nextCell = row.getCell(x);
				if (nextCell != null) {
					Cell newCell = row.createCell(x - 1, nextCell.getCellType());
					cloneCell(newCell, nextCell);
				}
			}
		}

		// Adjust the column widths
		for (int c = 0; c < maxColumn; c++) {
			sheet.setColumnWidth(c, sheet.getColumnWidth(c + 1));
		}
	}

	/*
	 * Takes an existing Cell and merges all the styles and forumla into the new
	 * one
	 */
	private static void cloneCell(Cell cNew, Cell cOld) {
		cNew.setCellComment(cOld.getCellComment());
		cNew.setCellStyle(cOld.getCellStyle());

		switch (cNew.getCellTypeEnum()) {
		case BOOLEAN: {
			cNew.setCellValue(cOld.getBooleanCellValue());
			break;
		}
		case NUMERIC: {
			cNew.setCellValue(cOld.getNumericCellValue());
			break;
		}
		case STRING: {
			cNew.setCellValue(cOld.getStringCellValue());
			break;
		}
		case ERROR: {
			cNew.setCellValue(cOld.getErrorCellValue());
			break;
		}
		case FORMULA: {
			cNew.setCellFormula(cOld.getCellFormula());
			break;
		}
		}

	}
}