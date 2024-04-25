package org.jeecgframework.poi.util;
/**
 * @author Link Xue 
 * @version 20171025
 * POI对EXCEL操作工具
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelUtil {

    public static void main(String[] args){
        //读取excel数据
        ArrayList<Map<String,String>> result = ExcelUtil.readExcelToObj("D:\\上传表.xlsx");
        for(Map<String,String> map:result){
            System.out.println(map);
        }

    }
    /**
     * 读取excel数据
     * @param path
     */
    public static ArrayList<Map<String,String>> readExcelToObj(String path) {

        Workbook wb = null;
        ArrayList<Map<String,String>> result = null;
        try {
            wb = WorkbookFactory.create(new File(path));
            result = readExcel(wb, 0, 2, 0);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    /**
     * 读取excel文件
     * @param wb
     * @param sheetIndex sheet页下标：从0开始
     * @param startReadLine 开始读取的行:从0开始
     * @param tailLine 去除最后读取的行
     */
    public static ArrayList<Map<String,String>> readExcel(Workbook wb,int sheetIndex, int startReadLine, int tailLine) {
        Sheet sheet = wb.getSheetAt(sheetIndex);
        Row row = null;
        ArrayList<Map<String,String>> result = new ArrayList<Map<String,String>>();
        for(int i=startReadLine; i<sheet.getLastRowNum()-tailLine+1; i++) {

            row = sheet.getRow(i);
            Map<String,String> map = new HashMap<String,String>();
            for(Cell c : row) {
                String returnStr = "";

                boolean isMerge = isMergedRegion(sheet, i, c.getColumnIndex());
                //判断是否具有合并单元格
                if(isMerge) {
                    String rs = getMergedRegionValue(sheet, row.getRowNum(), c.getColumnIndex());
//                    System.out.print(rs + "------ ");
                    returnStr = rs;
                }else {
//                    System.out.print(c.getRichStringCellValue()+"++++ ");
                    returnStr = c.getRichStringCellValue().getString();
                }
                if(c.getColumnIndex()==0){
                    map.put("id",returnStr);
                }else if(c.getColumnIndex()==1){
                    map.put("base",returnStr);
                }else if(c.getColumnIndex()==2){
                    map.put("siteName",returnStr);
                }else if(c.getColumnIndex()==3){
                    map.put("articleName",returnStr);
                }else if(c.getColumnIndex()==4){
                    map.put("mediaName",returnStr);
                }else if(c.getColumnIndex()==5){
                    map.put("mediaUrl",returnStr);
                }else if(c.getColumnIndex()==6){
                    map.put("newsSource",returnStr);
                }else if(c.getColumnIndex()==7){
                    map.put("isRecord",returnStr);
                }else if(c.getColumnIndex()==8){
                    map.put("recordTime",returnStr);
                }else if(c.getColumnIndex()==9){
                    map.put("remark",returnStr);
                }

            }
            result.add(map);
//            System.out.println();

        }
        return result;

    }

    /**
     * 获取合并单元格的值
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static String getMergedRegionValue(Sheet sheet ,int row , int column){
        int sheetMergeCount = sheet.getNumMergedRegions();

        for(int i = 0 ; i < sheetMergeCount ; i++){
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if(row >= firstRow && row <= lastRow){

                if(column >= firstColumn && column <= lastColumn){
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return getCellValue(fCell) ;
                }
            }
        }

        return null ;
    }

    /**
     * 判断合并了行
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static boolean isMergedRow(Sheet sheet,int row ,int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if(row == firstRow && row == lastRow){
                if(column >= firstColumn && column <= lastColumn){
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 判断指定的单元格是否是合并单元格
     * @param sheet
     * @param row 行下标
     * @param column 列下标
     * @return
     */
    public static boolean isMergedRegion(Sheet sheet,int row ,int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if(row >= firstRow && row <= lastRow){
                if(column >= firstColumn && column <= lastColumn){
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 判断sheet页中是否含有合并单元格
     * @param sheet
     * @return
     */
    public static boolean hasMerged(Sheet sheet) {
        return sheet.getNumMergedRegions() > 0 ? true : false;
    }

    /**
     * 合并单元格
     * @param sheet
     * @param firstRow 开始行
     * @param lastRow 结束行
     * @param firstCol 开始列
     * @param lastCol 结束列
     */
    public static void mergeRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        //update-begin-author:wangshuai date:20201118 for:一对多导出needMerge 子表数据对应数量小于2时报错 github#1840、gitee I1YH6B
        try{
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        }catch (IllegalArgumentException e){
            e.fillInStackTrace();
        }
        //update-end-author:wangshuai date:20201118 for:一对多导出needMerge 子表数据对应数量小于2时报错 github#1840、gitee I1YH6B
    }

    /**
     * 获取单元格的值
     * @param cell
     * @return
     */
    public static String getCellValue(Cell cell){

        if(cell == null) {
            return "";
        }

        if(cell.getCellTypeEnum() == CellType.STRING){

            return cell.getStringCellValue();

        }else if(cell.getCellTypeEnum() == CellType.BOOLEAN){

            return String.valueOf(cell.getBooleanCellValue());

        }else if(cell.getCellTypeEnum() ==  CellType.FORMULA){

            return cell.getCellFormula() ;

        }else if(cell.getCellTypeEnum() == CellType.NUMERIC){

            return String.valueOf(cell.getNumericCellValue());

        }
        return "";
    }
    
    /**
     * 数字值，去掉.0后缀
     * @param cell
     * @return
     */
    public static String remove0Suffix(Object value){
    	if(value!=null) {
    	    // update-begin-author:taoyan date:20210526 for:对于特殊的字符串 V1.0 也进行了去.0操作 这是不合理的
			String val = value.toString();
			if(val.endsWith(".0") && isNumberString(val)) {
				val = val.replace(".0", "");
			}
            // update-end-author:taoyan date:20210526 for:对于特殊的字符串 V1.0 也进行了去.0操作 这是不合理的
			return val;
		}
        return null;
    }

    /**
     * 判断给定的字符串是不是只有数字
     * @param str
     * @return
     */
    private static boolean isNumberString(String str){
        //update-begin---author:chenrui ---date:20240424  for：[QQYUN-9048]负数被识别成非数字------------
        String regex = "^-?[0-9]+\\.0+$";
        //update-end---author:chenrui ---date:20240424  for：[QQYUN-9048]负数被识别成非数字------------
        Pattern pattern = Pattern.compile(regex);
        Matcher m = pattern.matcher(str);
        if (m.find()) {
            return true;
        }
        return false;
    }
    
    /**
     * 从excel读取内容
     */
    public static void readContent(String fileName)  {
        boolean isE2007 = false;    //判断是否是excel2007格式
        if(fileName.endsWith("xlsx"))
            isE2007 = true;
        try {
            InputStream input = new FileInputStream(fileName);  //建立输入流
            Workbook wb  = null;
            //根据文件格式(2003或者2007)来初始化
            if(isE2007)
                wb = new XSSFWorkbook(input);
            else
                wb = new HSSFWorkbook(input);
            Sheet sheet = wb.getSheetAt(0);     //获得第一个表单
            Iterator<Row> rows = sheet.rowIterator(); //获得第一个表单的迭代器
            while (rows.hasNext()) {
                Row row = rows.next();  //获得行数据
                System.out.println("Row #" + row.getRowNum());  //获得行号从0开始
                Iterator<Cell> cells = row.cellIterator();    //获得第一行的迭代器
                while (cells.hasNext()) {
                    Cell cell = cells.next();
                    System.out.println("Cell #" + cell.getColumnIndex());
                    switch (cell.getCellTypeEnum()) {   //根据cell中的类型来输出数据
                        case NUMERIC:
                            System.out.println(cell.getNumericCellValue());
                            break;
                        case STRING:
                            System.out.println(cell.getStringCellValue());
                            break;
                        case BOOLEAN:
                            System.out.println(cell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            System.out.println(cell.getCellFormula());
                            break;
                        default:
                            System.out.println("unsuported sell type======="+cell.getCellTypeEnum());
                            break;
                    }
                }
            }
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

}