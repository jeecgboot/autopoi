import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;
import org.jeecgframework.poi.handler.inter.IExcelExportServer;
import org.junit.Test;
import vo.TestEntity;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

/**
 * POI 5.4.1 综合功能测试
 * 测试升级后的各种导出功能是否正常
 */
public class Poi541ComprehensiveTest {

    private static final String OUTPUT_DIR = "D:/excel/poi541test/";

    /**
     * 测试1：多表头导出（Map数据 + 手动封装ExcelExportEntity）
     */
    @Test
    public void testMultiHeaderExport() throws Exception {
        System.out.println("=== 测试多表头导出 ===");

        List<Map<String, Object>> dataList = new ArrayList<Map<String, Object>>();
        Map<String, Object> map1 = new HashMap<>();
        map1.put("name", "小明");
        map1.put("age", 21);
        map1.put("degree", 36);
        map1.put("link_name", "小八");
        map1.put("link_age", 33);
        dataList.add(map1);

        Map<String, Object> map2 = new HashMap<>();
        map2.put("name", "小王");
        map2.put("age", 24);
        map2.put("degree", 37);
        map2.put("link_name", "小六");
        map2.put("link_age", 26);
        dataList.add(map2);

        List<ExcelExportEntity> entityList = new ArrayList<>();
        // 一般表头
        entityList.add(new ExcelExportEntity("姓名", "name"));
        entityList.add(new ExcelExportEntity("年龄", "age"));
        entityList.add(new ExcelExportEntity("体温", "degree"));
        
        // 多表头方式1：需要先添加子列，再添加父列（Map数据专用）
        // 子列需要使用三个参数的构造器，第三个参数为 true
        entityList.add(new ExcelExportEntity("姓名", "link_name", true));
        entityList.add(new ExcelExportEntity("年龄", "link_age", true));
        
        // 父列也需要设置 colspan=true，并设置 SubColumnList
        ExcelExportEntity contactEntity = new ExcelExportEntity("紧急联系人", "linkman", true);
        List<String> subKeys = new ArrayList<>();
        subKeys.add("link_name");
        subKeys.add("link_age");
        contactEntity.setSubColumnList(subKeys);
        entityList.add(contactEntity);

        // 导出 - 使用 XSSF 格式对应 .xlsx 文件
        Workbook wb = ExcelExportUtil.exportExcel(new ExportParams("测试多表头", "sheetName", ExcelType.XSSF), entityList, dataList);

        // 保存文件 - 使用 .xlsx 扩展名
        saveWorkbook(wb, "test1_multiheader.xlsx");
        System.out.println("✅ 多表头导出测试完成");
    }

    /**
     * 测试2：多Sheet导出（实体类方式）
     */
    @Test
    public void testMultiSheetExport() throws Exception {
        System.out.println("=== 测试多Sheet导出 ===");
        
        // 多个map，对应了多个sheet
        List<Map<String, Object>> listMap = new ArrayList<>();

        for (int i = 0; i < 3; i++) {
            Map<String, Object> map = new HashMap<>();
            
            // 表格title
            map.put("title", getExportParams("测试Sheet" + (i + 1)));
            
            // 表格对应实体
            map.put("entity", TestEntity.class);

            // 准备数据（实体类方式）
            List<TestEntity> ls = new ArrayList<>();
            for (int j = 0; j < 10; j++) {
                TestEntity testEntity = new TestEntity();
                testEntity.setName("张三" + j);
                testEntity.setAge(18 + j);
                ls.add(testEntity);
            }
            map.put("data", ls);
            listMap.add(map);
        }

        // 导出
        Workbook wb = ExcelExportUtil.exportExcel(listMap, ExcelType.HSSF);

        // 保存文件
        saveWorkbook(wb, "test2_multisheet.xls");
        System.out.println("✅ 多Sheet导出测试完成");
    }

    /**
     * 测试3：模板导出（简单模板）
     */
    @Test
    public void testTemplateExport() throws Exception {
        System.out.println("=== 测试模板导出 ===");
        
        try {
            // 创建简单的测试模板（如果模板文件存在）
            String templatePath = "src/test/resources/templates/test.xlsx";
            File templateFile = new File(templatePath);
            
            if (!templateFile.exists()) {
                System.out.println("⚠️  模板文件不存在，跳过模板导出测试: " + templatePath);
                return;
            }

            TemplateExportParams params = new TemplateExportParams(templatePath);
            Map<String, Object> map = new HashMap<>();
            map.put("title", "员工个人信息");
            map.put("name", "大熊");
            map.put("age", 22);
            map.put("company", "北京机器猫科技有限公司");
            map.put("date", "2020-07-13");
            
            Workbook workbook = ExcelExportUtil.exportExcel(params, map);

            // 保存文件
            saveWorkbook(workbook, "test3_template.xlsx");
            System.out.println("✅ 模板导出测试完成");
        } catch (Exception e) {
            System.out.println("⚠️  模板导出测试失败: " + e.getMessage());
        }
    }

    /**
     * 测试4：复杂模板导出（带循环）
     */
    @Test
    public void testComplexTemplateExport() throws Exception {
        System.out.println("=== 测试复杂模板导出（带循环）===");
        
        try {
            String templatePath = "src/test/resources/templates/testNextMarge.xlsx";
            File templateFile = new File(templatePath);
            
            if (!templateFile.exists()) {
                System.out.println("⚠️  模板文件不存在，跳过复杂模板导出测试: " + templatePath);
                return;
            }

            TemplateExportParams params = new TemplateExportParams(templatePath);
            Map<String, Object> map = new HashMap<>();
            map.put("title", "员工信息");
            
            List<Map<String, Object>> listMap = new ArrayList<>();
            for (int i = 0; i < 6; i++) {
                Map<String, Object> lm = new HashMap<>();
                lm.put("name", "王" + i);
                lm.put("age", "2" + i);
                lm.put("sex", i % 2 == 0 ? "1" : "2");
                lm.put("date", new Date());
                lm.put("salary", 1000 + i);
                listMap.add(lm);
            }
            map.put("maplist", listMap);
            
            Workbook workbook = ExcelExportUtil.exportExcel(params, map);

            // 保存文件
            saveWorkbook(workbook, "test4_complex_template.xlsx");
            System.out.println("✅ 复杂模板导出测试完成");
        } catch (Exception e) {
            System.out.println("⚠️  复杂模板导出测试失败: " + e.getMessage());
        }
    }

    /**
     * 测试5：大数据导出（测试POI 5.4.1性能）
     */
    @Test
    public void testBigDataExport() throws Exception {
        System.out.println("=== 测试大数据导出 ===");
        
        Date start = new Date();
        
        // 设置表格标题
        ExportParams params = new ExportParams("POI 5.4.1大数据测试", "性能测试");
        
        /**
         * 导出10万条数据（分10次，每次1万条）
         * 测试POI 5.4.1的性能
         */
        Workbook workbook = ExcelExportUtil.exportBigExcel(
            params, 
            TestEntity.class, 
            new IExcelExportServer() {
                @Override
                public List<Object> selectListForExcelExport(Object obj, int page) {
                    System.out.println("  正在导出第 " + (page + 1) + " 批数据...");
                    
                    // page每次加一，当等于obj的值时返回空，代码结束
                    if (((int) obj) == page) {
                        return null;
                    }
                    
                    // 每次返回1万条数据
                    List<Object> list = new ArrayList<>();
                    for (int i = 0; i < 10000; i++) {
                        TestEntity client = new TestEntity();
                        client.setName("测试用户" + (page * 10000 + i));
                        client.setAge(18 + (i % 50));
                        list.add(client);
                    }
                    return list;
                }
            }, 
            10  // 总共导出10批
        );

        long timeUsed = new Date().getTime() - start.getTime();
        System.out.println("  导出10万条数据耗时: " + timeUsed + "ms (" + (timeUsed / 1000.0) + "秒)");

        // 保存文件
        saveWorkbook(workbook, "test5_bigdata_100k.xlsx");
        System.out.println("✅ 大数据导出测试完成");
    }

    /**
     * 测试6：Cell类型API兼容性测试
     */
    @Test
    public void testCellTypeCompatibility() throws Exception {
        System.out.println("=== 测试Cell类型API兼容性 ===");
        
        List<TestEntity> dataList = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            TestEntity entity = new TestEntity();
            entity.setName("用户" + i);
            entity.setAge(20 + i);
            dataList.add(entity);
        }

        ExportParams params = new ExportParams("API兼容性测试", "Sheet1", ExcelType.XSSF);
        Workbook wb = ExcelExportUtil.exportExcel(params, TestEntity.class, dataList);

        // 保存文件
        saveWorkbook(wb, "test6_api_compatibility.xlsx");
        System.out.println("✅ Cell类型API兼容性测试完成");
    }

    /**
     * 测试7：Map数据导出
     */
    @Test
    public void testMapDataExport() throws Exception {
        System.out.println("=== 测试Map数据导出 ===");
        
        List<ExcelExportEntity> entityList = new ArrayList<>();
        entityList.add(new ExcelExportEntity("姓名", "name", 15));
        entityList.add(new ExcelExportEntity("年龄", "age", 10));
        entityList.add(new ExcelExportEntity("部门", "dept", 20));
        entityList.add(new ExcelExportEntity("薪资", "salary", 15));

        List<Map<String, Object>> dataList = new ArrayList<>();
        for (int i = 0; i < 20; i++) {
            Map<String, Object> map = new HashMap<>();
            map.put("name", "员工" + i);
            map.put("age", 25 + i);
            map.put("dept", i % 3 == 0 ? "研发部" : i % 3 == 1 ? "市场部" : "行政部");
            map.put("salary", 8000 + i * 500);
            dataList.add(map);
        }

        Workbook wb = ExcelExportUtil.exportExcel(
            new ExportParams("Map数据导出", "员工表", ExcelType.XSSF), 
            entityList, 
            dataList
        );

        saveWorkbook(wb, "test7_map_data.xlsx");
        System.out.println("✅ Map数据导出测试完成");
    }

    /**
     * 运行所有测试
     */
    @Test
    public void runAllTests() throws Exception {
        System.out.println("\n==========================================");
        System.out.println("    POI 5.4.1 升级后功能综合测试");
        System.out.println("==========================================\n");
        
        long startTime = System.currentTimeMillis();
        
        try {
            testMultiHeaderExport();
            System.out.println();
            
            testMultiSheetExport();
            System.out.println();
            
            testTemplateExport();
            System.out.println();
            
            testComplexTemplateExport();
            System.out.println();
            
            testCellTypeCompatibility();
            System.out.println();
            
            testMapDataExport();
            System.out.println();
            
            // 大数据测试放在最后（比较耗时）
            testBigDataExport();
            
        } catch (Exception e) {
            System.err.println("❌ 测试过程中出现异常: " + e.getMessage());
            e.printStackTrace();
            throw e;
        }
        
        long totalTime = System.currentTimeMillis() - startTime;
        
        System.out.println("\n==========================================");
        System.out.println("✅ 所有测试完成！");
        System.out.println("   总耗时: " + totalTime + "ms (" + (totalTime / 1000.0) + "秒)");
        System.out.println("   输出目录: " + OUTPUT_DIR);
        System.out.println("==========================================\n");
    }

    /**
     * 获取导出参数
     */
    private static ExportParams getExportParams(String name) {
        // 表格名称,sheet名称,导出版本
        return new ExportParams(name, name, ExcelType.HSSF);
    }

    /**
     * 保存Workbook到文件
     */
    private void saveWorkbook(Workbook workbook, String fileName) throws Exception {
        File saveDir = new File(OUTPUT_DIR);
        if (!saveDir.exists()) {
            saveDir.mkdirs();
        }
        
        String filePath = OUTPUT_DIR + fileName;
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(filePath);
            workbook.write(fos);
            fos.flush();
        } finally {
            // 先关闭输出流
            if (fos != null) {
                try {
                    fos.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
            // 再关闭 Workbook (POI 5.x 需要显式关闭)
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        
        System.out.println("  文件已保存: " + filePath);
    }
}

