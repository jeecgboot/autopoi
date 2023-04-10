package org.jeecgframework.poi.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * poi 4.0 07版本在 shift操作下有bug,不移动了单元格以及单元格样式,没有移动cell
 * cell还是复用的原理的cell,导致wb输出的时候没有输出值
 * 等待修复的时候删除这个问题
 *
 * @author by jueyue on 19-6-17.
 */
public class PoiExcelTempUtil {

    /**
     * 把这N行的数据,cell重新设置下,修复因为shift的浅复制问题,导致文本不显示的错误
     *
     * @param sheet
     * @param startRow
     * @param endRow
     */
    public static void reset(Sheet sheet, int startRow, int endRow) {
        if (sheet.getWorkbook() instanceof HSSFWorkbook) {
            return;
        }
        for (int i = startRow; i <= endRow; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            int cellNum = row.getLastCellNum();
            for (int j = 0; j < cellNum; j++) {
                if (row.getCell(j) == null) {
                    continue;
                }
                Map<String, Object> map = copyCell(row.getCell(j));
                row.removeCell(row.getCell(j));
                Cell cell = row.createCell(j);
                cell.setCellStyle((CellStyle) map.get("cellStyle"));
                if ((boolean) map.get("isDate")) {
                    cell.setCellValue((Date) map.get("value"));
                } else {
                    CellType cellType = (CellType) map.get("cellType");
                    switch (cellType) {
                        case NUMERIC:
                            cell.setCellValue((double) map.get("value"));
                            break;
                        case STRING:
                            cell.setCellValue((String) map.get("value"));
                        case FORMULA:
                            break;
                        case BLANK:
                            break;
                        case BOOLEAN:
                            cell.setCellValue((boolean) map.get("value"));
                        case ERROR:
                            break;
                    }
                }
            }
        }

    }

    private static Map copyCell(Cell cell) {
        Map<String, Object> map = new HashMap<>();
        map.put("cellType", cell.getCellType());
        map.put("isDate", CellType.NUMERIC == cell.getCellType() && DateUtil.isCellDateFormatted(cell));
        map.put("value", getValue(cell));
        map.put("cellStyle", cell.getCellStyle());
        return map;
    }

    private static Object getValue(Cell cell) {
        if (CellType.NUMERIC == cell.getCellType() && DateUtil.isCellDateFormatted(cell)) {
            return cell.getDateCellValue();
        }
        switch (cell.getCellType()) {
            case _NONE:
                return null;
            case NUMERIC:
                return cell.getNumericCellValue();
            case STRING:
                return cell.getStringCellValue();
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                break;
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case ERROR:
                break;
        }
        return null;
    }


}
