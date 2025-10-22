import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;
import vo.TestEntity;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Description: 多Sheet导出测试
 * 测试功能：
 * 1. 测试一个 Workbook 中创建多个 Sheet 页签
 * 2. 测试使用实体类（TestEntity）方式导出数据到不同 Sheet
 * 3. 测试 XSSF 格式（.xlsx）的多 Sheet 导出
 * 
 * 参考文档：http://doc.jeecg.com/2044223
 * 
 * @author: scott
 * @date: 2020年09月16日 11:46
 */
public class DaoChuSheetTest {
    // 生成文件的保存路径
    private static final String generatePath = "D:/excel/";

    /**
     * 获取导出参数配置
     * @param name 表格名称和Sheet名称
     * @return ExportParams 配置对象，使用 XSSF 格式
     */
    public static ExportParams getExportParams(String name) {
        return  new ExportParams(name,name,ExcelType.XSSF);
    }
    
    /**
     * 测试多Sheet导出功能
     * 
     * 测试内容：
     * - 创建3个Sheet页签
     * - 每个Sheet使用相同的实体类（TestEntity）
     * - 每个Sheet包含10条测试数据
     * - 验证多Sheet在同一个Excel文件中的导出效果
     * 
     * @return Workbook 对象，包含3个Sheet页签
     */
    public static Workbook test() {
        /**
         * 多个Map配置：
         * - title: 对应表格标题和Sheet名称（ExportParams对象）
         * - entity: 对应表格数据的实体类（如 TestEntity.class）
         * - data: 对应实际数据集合（Collection类型）
         * 
         * 注意：也可以使用 Map 数据替代实体类，示例中 ls2 展示了这种方式
         */
        List<Map<String, Object>> listMap = new ArrayList<Map<String, Object>>();
        
        // 创建3个Sheet
        for(int i=0;i<3;i++){
            Map<String, Object> map = new HashMap<String, Object>();
            map.put("title", getExportParams("测试"+i));//表格Title
            map.put("entity", TestEntity.class);//表格对应实体
            
            // 方式1：使用实体类数据
            List<TestEntity> ls=new ArrayList<TestEntity> ();
            for(int j=0;j<10;j++){
                TestEntity testEntity = new TestEntity();
                testEntity.setName("张三"+i+j);
                testEntity.setAge(18+i+j);
                ls.add(testEntity);
            }
            
            // 方式2：使用Map数据（示例代码，未使用）
            List<Map> ls2=new ArrayList<Map> ();
            for(int j=0;j<10;j++){
                Map map1 = new HashMap();
                map1.put("name","李四"+i+j);
                map1.put("age",18+i+j);
                ls2.add(map1);
            }
            
            map.put("data", ls);//可选：ls（实体类数据） or ls2（Map数据）
            listMap.add(map);
        }
        
        // 导出多个Sheet到同一个Workbook
        Workbook workbook = ExcelExportUtil.exportExcel(listMap, ExcelType.XSSF);
        return workbook;
    }

    /**
     * 主方法：执行测试并保存Excel文件
     * 
     * 执行步骤：
     * 1. 调用test()方法生成包含3个Sheet的Workbook
     * 2. 检查保存目录是否存在，不存在则创建
     * 3. 将Workbook写入到文件 testSheet.xlsx
     * 4. 关闭文件输出流
     * 
     * 预期结果：
     * - 在 D:/excel/ 目录下生成 testSheet.xlsx 文件
     * - 文件包含3个Sheet：测试0、测试1、测试2
     * - 每个Sheet包含10行数据
     * 
     * @param args 命令行参数
     * @throws IOException 文件写入异常
     */
    public static void main(String[] args) throws IOException {
        Workbook workbook = test();
        File savefile = new File(generatePath);
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream(generatePath + "testSheet.xlsx");
        workbook.write(fos);
        fos.close();
    }
}
