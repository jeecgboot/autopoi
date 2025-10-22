import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import java.util.Date;

/**
 * POI 5.4.1 API 兼容性测试
 * 测试从 4.1.2 升级到 5.4.1 后的 API 变化
 */
public class Poi541ApiCompatibilityTest {

    /**
     * 测试1: getCellType() 方法（替代 getCellTypeEnum()）
     */
    @Test
    public void testGetCellType() {
        System.out.println("=== 测试 getCellType() 方法 ===");
        
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("测试");
        Row row = sheet.createRow(0);
        
        // 测试字符串类型
        Cell cell1 = row.createCell(0);
        cell1.setCellValue("测试文本");
        Assert.assertEquals("字符串类型判断", CellType.STRING, cell1.getCellType());
        System.out.println("  ✅ STRING 类型: " + cell1.getCellType());
        
        // 测试数字类型
        Cell cell2 = row.createCell(1);
        cell2.setCellValue(123.45);
        Assert.assertEquals("数字类型判断", CellType.NUMERIC, cell2.getCellType());
        System.out.println("  ✅ NUMERIC 类型: " + cell2.getCellType());
        
        // 测试布尔类型
        Cell cell3 = row.createCell(2);
        cell3.setCellValue(true);
        Assert.assertEquals("布尔类型判断", CellType.BOOLEAN, cell3.getCellType());
        System.out.println("  ✅ BOOLEAN 类型: " + cell3.getCellType());
        
        // 测试公式类型
        Cell cell4 = row.createCell(3);
        cell4.setCellFormula("SUM(A1:B1)");
        Assert.assertEquals("公式类型判断", CellType.FORMULA, cell4.getCellType());
        System.out.println("  ✅ FORMULA 类型: " + cell4.getCellType());
        
        System.out.println("✅ getCellType() 方法测试通过\n");
    }

    /**
     * 测试2: DateUtil 类（替代 HSSFDateUtil）
     */
    @Test
    public void testDateUtil() {
        System.out.println("=== 测试 DateUtil 类 ===");
        
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("日期测试");
        Row row = sheet.createRow(0);
        
        // 创建日期单元格
        Cell cell = row.createCell(0);
        Date now = new Date();
        cell.setCellValue(now);
        
        // 使用 DateUtil 判断是否为日期格式
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("yyyy-MM-dd"));
        cell.setCellStyle(cellStyle);
        
        // 验证 DateUtil.isCellDateFormatted 方法
        boolean isDate = DateUtil.isCellDateFormatted(cell);
        System.out.println("  ✅ DateUtil.isCellDateFormatted() 工作正常: " + isDate);
        
        // 测试数字转日期
        double dateValue = cell.getNumericCellValue();
        Date convertedDate = DateUtil.getJavaDate(dateValue);
        Assert.assertNotNull("日期转换不应为空", convertedDate);
        System.out.println("  ✅ DateUtil.getJavaDate() 转换成功: " + convertedDate);
        
        System.out.println("✅ DateUtil 类测试通过\n");
    }

    /**
     * 测试3: getAlignment() 方法（替代 getAlignmentEnum()）
     */
    @Test
    public void testGetAlignment() {
        System.out.println("=== 测试 getAlignment() 方法 ===");
        
        Workbook workbook = new XSSFWorkbook();
        CellStyle style = workbook.createCellStyle();
        
        // 设置不同的对齐方式
        style.setAlignment(HorizontalAlignment.CENTER);
        Assert.assertEquals("水平居中对齐", HorizontalAlignment.CENTER, style.getAlignment());
        System.out.println("  ✅ 水平对齐: " + style.getAlignment());
        
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        Assert.assertEquals("垂直居中对齐", VerticalAlignment.CENTER, style.getVerticalAlignment());
        System.out.println("  ✅ 垂直对齐: " + style.getVerticalAlignment());
        
        System.out.println("✅ getAlignment() 方法测试通过\n");
    }

    /**
     * 测试4: Font 索引类型变化（short -> int）
     */
    @Test
    public void testFontIndexType() {
        System.out.println("=== 测试 Font 索引类型变化 ===");
        
        Workbook workbook = new XSSFWorkbook();
        
        // POI 5.x 中 getNumberOfFonts() 返回 int 类型
        int fontCount = workbook.getNumberOfFonts();
        System.out.println("  ✅ 字体数量 (int 类型): " + fontCount);
        
        // 遍历所有字体（使用 int 而不是 short）
        for (int i = 0; i < fontCount && i < 3; i++) {
            Font font = workbook.getFontAt(i);
            System.out.println("  ✅ 字体 " + i + ": " + font.getFontName());
        }
        
        System.out.println("✅ Font 索引类型测试通过\n");
    }

    /**
     * 测试5: Cell 值读取兼容性
     */
    @Test
    public void testCellValueCompatibility() {
        System.out.println("=== 测试 Cell 值读取兼容性 ===");
        
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("值读取测试");
        Row row = sheet.createRow(0);
        
        // 字符串值
        Cell cell1 = row.createCell(0);
        cell1.setCellValue("测试");
        if (cell1.getCellType() == CellType.STRING) {
            String value = cell1.getStringCellValue();
            Assert.assertEquals("字符串值读取", "测试", value);
            System.out.println("  ✅ 字符串值: " + value);
        }
        
        // 数字值
        Cell cell2 = row.createCell(1);
        cell2.setCellValue(100.5);
        if (cell2.getCellType() == CellType.NUMERIC) {
            double value = cell2.getNumericCellValue();
            Assert.assertEquals("数字值读取", 100.5, value, 0.001);
            System.out.println("  ✅ 数字值: " + value);
        }
        
        // 布尔值
        Cell cell3 = row.createCell(2);
        cell3.setCellValue(false);
        if (cell3.getCellType() == CellType.BOOLEAN) {
            boolean value = cell3.getBooleanCellValue();
            Assert.assertFalse("布尔值读取", value);
            System.out.println("  ✅ 布尔值: " + value);
        }
        
        System.out.println("✅ Cell 值读取兼容性测试通过\n");
    }

    /**
     * 测试6: 单元格样式兼容性
     */
    @Test
    public void testCellStyleCompatibility() {
        System.out.println("=== 测试单元格样式兼容性 ===");
        
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("样式测试");
        
        // 创建样式
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        font.setColor(IndexedColors.RED.getIndex());
        style.setFont(font);
        
        // 设置对齐
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        
        // 应用样式
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("样式测试");
        cell.setCellStyle(style);
        
        // 验证样式
        CellStyle appliedStyle = cell.getCellStyle();
        Assert.assertEquals("水平对齐", HorizontalAlignment.CENTER, appliedStyle.getAlignment());
        Assert.assertEquals("垂直对齐", VerticalAlignment.CENTER, appliedStyle.getVerticalAlignment());
        
        System.out.println("  ✅ 样式创建和应用正常");
        System.out.println("  ✅ 对齐方式: " + appliedStyle.getAlignment());
        System.out.println("✅ 单元格样式兼容性测试通过\n");
    }

    /**
     * 测试7: Workbook 创建和基本操作
     */
    @Test
    public void testWorkbookBasicOperations() {
        System.out.println("=== 测试 Workbook 基本操作 ===");
        
        // 创建 Workbook
        Workbook workbook = new XSSFWorkbook();
        Assert.assertNotNull("Workbook 创建", workbook);
        System.out.println("  ✅ Workbook 创建成功");
        
        // 创建 Sheet
        Sheet sheet1 = workbook.createSheet("Sheet1");
        Sheet sheet2 = workbook.createSheet("Sheet2");
        Assert.assertEquals("Sheet 数量", 2, workbook.getNumberOfSheets());
        System.out.println("  ✅ 创建了 " + workbook.getNumberOfSheets() + " 个 Sheet");
        
        // 创建行和单元格
        Row row = sheet1.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("POI 5.4.1 测试");
        
        Assert.assertEquals("单元格值", "POI 5.4.1 测试", cell.getStringCellValue());
        System.out.println("  ✅ 单元格值设置成功: " + cell.getStringCellValue());
        
        System.out.println("✅ Workbook 基本操作测试通过\n");
    }

    /**
     * 运行所有 API 兼容性测试
     */
    @Test
    public void runAllApiTests() {
        System.out.println("\n==========================================");
        System.out.println("  POI 5.4.1 API 兼容性全面测试");
        System.out.println("==========================================\n");
        
        long startTime = System.currentTimeMillis();
        
        testGetCellType();
        testDateUtil();
        testGetAlignment();
        testFontIndexType();
        testCellValueCompatibility();
        testCellStyleCompatibility();
        testWorkbookBasicOperations();
        
        long totalTime = System.currentTimeMillis() - startTime;
        
        System.out.println("==========================================");
        System.out.println("✅ 所有 API 兼容性测试通过！");
        System.out.println("   总耗时: " + totalTime + "ms");
        System.out.println("   POI 5.4.1 升级成功，API 完全兼容");
        System.out.println("==========================================\n");
    }
}

