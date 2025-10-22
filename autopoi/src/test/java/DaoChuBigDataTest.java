import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.handler.inter.IExcelExportServer;
import vo.TestEntity;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @Description: 大数据量 Excel 导出测试
 * 测试功能：
 * 1. 测试百万级数据的 Excel 导出性能
 * 2. 测试分批次导出大数据，避免内存溢出
 * 3. 验证 exportBigExcel 方法的实际应用效果
 * 
 * 技术要点：
 * - 使用 IExcelExportServer 接口实现分批数据加载
 * - 每批次处理 10万条数据，避免一次性加载导致内存溢出
 * - 适用于数据量超过 10万条的大数据导出场景
 * 
 * 参考文档：
 * - http://doc.wupaas.com/docs/easypoi/easypoi-1c10lbsojh62f
 * - https://blog.csdn.net/weixin_45214729/article/details/118552415
 * 
 * @author: liusq
 * @date: 2022年1月4日
 */
public class DaoChuBigDataTest {
    // 生成文件的保存路径
    private static final String generatePath = "D:/excel/";
    
    /**
     * 大数据导出测试方法
     * 
     * 测试内容：
     * - 模拟生成 100万条测试数据
     * - 使用分批导出方式，每批处理 10万条数据
     * - 记录导出耗时，验证性能表现
     * 
     * 核心逻辑：
     * 1. 准备100万条测试数据（实际应用中应从数据库分批查询）
     * 2. 计算总页数（100万 / 10万 = 10页）
     * 3. 通过 IExcelExportServer 接口实现分批数据返回
     * 4. 每次返回10万条数据，POI 自动写入 Excel
     * 5. 当返回 null 时，表示数据处理完成
     * 
     * 性能优化建议：
     * - 每批数据量建议控制在 1-10万条之间
     * - 数据量过大（如30万/批）可能导致内存溢出
     * - 实际应用中应从数据库分页查询，而不是一次性加载到内存
     * 
     * @throws Exception 文件操作或数据处理异常
     */
    public static void bigDataExport() throws Exception {
        Workbook workbook = null;
        List<TestEntity> aList = new ArrayList<TestEntity>();
        ExportParams exportParams = new ExportParams();
        Date start = new Date();
        
        // 模拟100万条数据（实际应用中应该分批从数据库查询，避免内存溢出）
        System.out.println("开始准备100万条测试数据...");
        for(int j=0;j<1000000;j++){
            TestEntity testEntity = new TestEntity();
            testEntity.setName("李四"+j);
            testEntity.setAge(j);
            aList.add(testEntity);
        }
        
        // 计算分页参数：总页数和每页大小
        int totalPage = (aList.size() / 100000) + 1;  // 总共10页
        int pageSize = 100000;  // 每页10万条数据
        
        System.out.println("数据准备完成，开始分批导出，总页数：" + totalPage);
        
        /**
         * 使用 exportBigExcel 方法进行大数据导出
         * 
         * @param exportParams 导出参数配置
         * @param TestEntity.class 数��实体类
         * @param IExcelExportServer 数据服务接口，用于分批返回数据
         * @param totalPage 总页数，传递给 selectListForExcelExport 方法作为终止条件
         */
        workbook = ExcelExportUtil.exportBigExcel(exportParams, TestEntity.class, new IExcelExportServer() {
            /**
             * 分批数据查询接口实现
             * 
             * 该方法会被多次调用，每次返回一批数据，直到返回 null 为止
             * 
             * @param obj 就是上面传入的 totalPage，用于控制循环终止条件
             * @param page 当前页码（从1开始），每次自动+1
             * @return 当前批次的数据列表，返回 null 表示数据处理完成
             */
            @Override
            public List<Object> selectListForExcelExport(Object obj, int page) {
                System.out.println("正在处理第 " + page + " 页数据...");
                
                // 终止条件：当页码超过总页数时，返回 null
                if (page > totalPage) {
                    return null;
                }
                
                // 计算当前批次的数据范围
                int fromIndex = (page - 1) * pageSize;  // 起始索引
                int toIndex = page != totalPage ? fromIndex + pageSize : aList.size();  // 结束索引
                
                // 返回当前批次的数据（使用 subList 截取）
                // 重要提示：实际应用中应该从数据库分页查询，而不是从内存中截取
                List<Object> list = new ArrayList<>();
                list.addAll(aList.subList(fromIndex, toIndex));
                return list;
            }
        }, totalPage);
        
        // 保存文件
        File savefile = new File(generatePath);
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("D:/excel/ExcelExportBigData.bigDataExport.xlsx");
        workbook.write(fos);
        fos.close();
        
        // 输出性能统计
        long timeUsed = new Date().getTime() - start.getTime();
        System.out.println("导出完成！耗时(秒)：" + (timeUsed / 1000) + "，文件保存在：D:/excel/ExcelExportBigData.bigDataExport.xlsx");
    }

    /**
     * 主方法：执行大数据导出测试
     * 
     * 预期结果：
     * - 在 D:/excel/ 目录下生成 ExcelExportBigData.bigDataExport.xlsx 文件
     * - 文件包含100万行数据
     * - 控制台输出各批次处理进度和总耗时
     * 
     * @param args 命令行参数
     * @throws Exception 导出过程异常
     */
    public static void main(String[] args) throws Exception {
       bigDataExport();
    }
}
