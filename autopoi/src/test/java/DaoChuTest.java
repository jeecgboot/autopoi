import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;
import org.jeecgframework.poi.excel.entity.params.ExcelExportEntity;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Description: Excel 模板导出测试
 * 测试功能：
 * 1. 测试基于 Excel 模板文件的数据导出功能
 * 2. 测试模板中的变量替换和列表数据循环填充
 * 3. 验证模板导出在实际业务场景中的应用
 *
 * 测试场景：
 * - 使用预定义的 Excel 模板（test.xlsx 或 testNextMarge.xlsx）
 * - 将 Map 数据填充到模板的占位符位置
 * - 支持单个变量替换和列表数据的循环生成
 *
 * @author: scott
 * @date: 2020年09月16日 11:46
 */
public class DaoChuTest {
    // 模板文件路径：项目根目录/autopoi/src/test/resources/templates/
    private static final String TEMPLATE_PATH = System.getProperty("user.dir") + File.separator + "autopoi" + File.separator + "src" + File.separator + "test" + File.separator + "resources" + File.separator + "templates" + File.separator;

    /**
     * 获取模板导出参数配置
     *
     * @param name 模板文件名（不含扩展名），如 "test" 或 "testNextMarge"
     * @return TemplateExportParams 模板导出参数对象
     */
    public static TemplateExportParams getTemplateParams(String name) {
        return new TemplateExportParams(TEMPLATE_PATH + name + ".xlsx");
    }

    /**
     * 测试模板导出功能
     *
     * 测试内容：
     * - 加载指定名称的 Excel 模板文件
     * - 构造测试数据（包含3条记录的列表）
     * - 将数据填充到模板的 autoList 循环区域
     * - 生成包含实际数据的 Excel 文件
     *
     * 数据结构：
     * - autoList: 列表数据，会在模板中循环生成多行
     *   - name: 姓名字段
     *   - isTts: 序号字段
     *   - sname: 简称字段
     *   - ttsContent: 内容字段
     *   - rate: 费率字段
     *
     * @param name 模板文件名
     * @return Workbook 填充数据后的工作簿对象
     */
    public static Workbook test(String name) {
        TemplateExportParams params = getTemplateParams(name);
        Map<String, Object> map = new HashMap<String, Object>();

        // 构造列表数据（3条记录）
        List<Map<String, Object>> listMap = new ArrayList<Map<String, Object>>();
        for (int i = 0; i < 3; i++) {
            Map<String, Object> lm = new HashMap<String, Object>();
            lm.put("name", "姓名1" + i);
            lm.put("isTts", i);
            lm.put("sname", "s姓名");
            lm.put("ttsContent", "ttsContent内容");
            lm.put("rate", 1000 + i);
            listMap.add(lm);
        }

        // 将列表数据放入 map，对应模板中的 {{$fe: autoList}} 循环标记
        map.put("autoList", listMap);

        // 基于模板和数据生成 Excel
        Workbook workbook = ExcelExportUtil.exportExcel(params, map);
        return workbook;
    }

    /**
     * 测试 Map 数据导出功能（使用 ExcelExportEntity）
     *
     * 测试内容：
     * - 手动构造 ExcelExportEntity 定义 Excel 列结构
     * - 使用 Map 方式填充数据（适用于动态列场景）
     * - 设置 sheetName 和 ExcelType.XSSF 格式
     * - 验证基于 ExportParams + ExcelExportEntity + Map 的导出方式
     *
     * 数据结构：
     * - ExcelExportEntity：定义列名、字段名、列宽
     * - Map数据：key 对应 ExcelExportEntity 的字段名
     *
     * @return Workbook 填充数据后的工作簿对象
     */
    public static Workbook testMapDataExport() {
        // 定义 Excel 列结构
        List<ExcelExportEntity> entityList = new ArrayList<>();
        entityList.add(new ExcelExportEntity("姓名", "name", 15));
        entityList.add(new ExcelExportEntity("年龄", "age", 10));
        entityList.add(new ExcelExportEntity("部门", "dept", 20));
        entityList.add(new ExcelExportEntity("薪资", "salary", 15));

        // 构造 Map 数据列表
        List<Map<String, Object>> result = new ArrayList<>();
        for (int i = 0; i < 20; i++) {
            Map<String, Object> map = new HashMap<>();
            map.put("name", "员工" + i);
            map.put("age", 25 + i);
            map.put("dept", i % 3 == 0 ? "研发部" : i % 3 == 1 ? "市场部" : "行政部");
            map.put("salary", 8000 + i * 500);
            result.add(map);
        }

        // 设置导出参数
        String sheetName = "员工信息表";
        ExportParams exportParams = new ExportParams(null, sheetName);
        exportParams.setType(ExcelType.XSSF);

        // 导出 Excel
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, entityList, result);
        return workbook;
    }

    /**
     * 主方法：执行模板导出测试并保存文件
     *
     * 执行步骤：
     * 1. 选择模板文件（test 或 testNextMarge）
     * 2. 调用 test() 方法基于模板生成 Workbook
     * 3. 检查保存目录是否存在
     * 4. 将生成的 Workbook 保存到 D:/excel/testNew.xlsx
     *
     * 预期结果：
     * - 在 D:/excel/ 目录下生成 testNew.xlsx 文件
     * - 文件内容基于模板结构，包含3行数据记录
     * - 模板中的占位符被实际数据替换
     *
     * @param args 命令行参数
     * @throws IOException 文件操作异常
     */
    public static void main(String[] args) throws IOException {
        String temName = "test";  // 简单模板
        String temNameNextM = "testNextMarge";  // 带合并单元格的复杂模板

        // 测试1：使用复杂模板进行测试
        Workbook workbook = test(temNameNextM);

        File savefile = new File(TEMPLATE_PATH);
        if (!savefile.exists()) {
            savefile.mkdirs();
        }

        FileOutputStream fos = new FileOutputStream("D:/excel/testNew.xlsx");
        workbook.write(fos);
        fos.close();

        // 测试2：Map 数据导出测试
        Workbook workbook2 = testMapDataExport();
        FileOutputStream fos2 = new FileOutputStream("D:/excel/testMapDataExport.xlsx");
        workbook2.write(fos2);
        fos2.close();
        
        System.out.println("✅ 所有测试完成，文件已保存到 D:/excel/ 目录");
    }
}
