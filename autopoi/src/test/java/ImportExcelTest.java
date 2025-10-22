import org.apache.commons.lang3.builder.ReflectionToStringBuilder;
import org.jeecgframework.poi.excel.ExcelImportUtil;
import org.jeecgframework.poi.excel.entity.ImportParams;
import vo.TestDateEntity;

import java.io.File;
import java.util.List;

/**
 * @Description: Excel 数据导入测试
 * 测试功能：
 * 1. 测试从 Excel 文件中读取数据并转换为 Java 对象
 * 2. 测试日期类型数据的导入和解析
 * 3. 验证导入参数配置（标题行、表头行）的正确性
 *
 * 测试场景：
 * - 读取 ExcelImportDateTest.xlsx 文件
 * - 将 Excel 数据映射到 TestDateEntity 实体类
 * - 支持自动类型转换和日期格式解析
 *
 * @author: scott
 * @date: 2020年09月16日 11:46
 */
public class ImportExcelTest {
    // 模板文件路径：项目根目录/autopoi/src/test/resources/templates/
    private static final String TEMPLATE_PATH = System.getProperty("user.dir") + File.separator + "autopoi" + File.separator + "src" + File.separator + "test" + File.separator + "resources" + File.separator + "templates" + File.separator;

    /**
     * 主方法：执行 Excel 导入测试
     *
     * 测试内容：
     * - 配置导入参数（标题行数、表头行数）
     * - 读取 Excel 文件并解析为实体对象列表
     * - 验证导入的数据数量和内容正确性
     *
     * 导入参数说明：
     * - TitleRows(1): 标题行占1行（通常是表格大标题）
     * - HeadRows(1): 表头行占1行（字段名称行）
     * - 实际数据从第3行开始读取
     *
     * 预期结果：
     * - 成功读取 Excel 文件中的所有数据行
     * - 数据正确映射到 TestDateEntity 对象
     * - 日期字段正确解析为 Date 类型
     * - 控制台输出数据总数和第2条数据的详细信息
     *
     * @param args 命令行参数
     * @throws Exception 文件读取或数据解析异常
     */
    public static void main(String[] args) throws Exception {
        // 配置导入参数
        ImportParams params = new ImportParams();
        params.setTitleRows(1);  // 标题行数：1行
        params.setHeadRows(1);   // 表头行数：1行

        // 指定要导入的 Excel 文件
        File importFile = new File(TEMPLATE_PATH + "ExcelImportDateTest.xlsx");

        // 执行导入：将 Excel 数据转换为 TestDateEntity 对象列表
        List<TestDateEntity> list = ExcelImportUtil.importExcel(importFile, TestDateEntity.class, params);

        // 输出导入结果
        System.out.println("导入数据总数：" + list.size());

        // 输出第2条数据的详细信息（索引为1）
        if (list.size() > 1) {
            System.out.println("第2条数据详情：");
            System.out.println(ReflectionToStringBuilder.toString(list.get(1)));
        }
    }
}
