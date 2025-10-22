import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jeecgframework.poi.word.WordExportUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Description: Word 文档导出测试
 * 测试功能：
 * 1. 测试基于 Word 模板的数据导出功能
 * 2. 测试 Word 文档中的变量替换和列表数据填充
 * 3. 验证 .docx 格式文档的模板导出效果
 *
 * 测试场景：
 * - 使用预定义的 Word 模板（纳税信息.docx）
 * - 将 Map 数据填充到模板的占位符位置
 * - 支持单个变量替换和列表数据循环生成
 *
 * 应用场景：
 * - 生成标准化的合同文档
 * - 批量生成报表文档
 * - 根据数据自动生成通知、证明等文档
 *
 * @author: scott
 * @date: 2020年09月16日 11:46
 */
public class DaoChuWordTest {
    // 模板文件路径：项目根目录/autopoi/src/test/resources/templates/
    private static final String TEMPLATE_PATH = System.getProperty("user.dir") + File.separator + "autopoi" + File.separator + "src" + File.separator + "test" + File.separator + "resources" + File.separator + "templates" + File.separator;

    /**
     * 主方法：执行 Word 文档导出测试
     *
     * 测试内容：
     * - 准备纳税信息数据（单条记录和列表数据）
     * - 使用 Word 模板填充数据
     * - 生成新的 Word 文档
     *
     * 数据结构说明：
     * - taxlist: 纳税列表数据（可循环生成多行，当前为空）
     * - totalpreyear: 去年总额
     * - totalthisyear: 今年总额
     * - type: 税种类型
     * - presum: 上期金额
     * - thissum: 本期金额
     * - curmonth: 当前月份
     * - now: 当前时间
     *
     * 模板语法：
     * - 单个变量：{{variableName}}
     * - 列表循环：{{$fe: listName t.field}}
     *
     * 预期结果：
     * - 在 D:/excel/ 目录下生成 纳税信息new.docx 文件
     * - 文档中的占位符被实际数据替换
     * - 如果有列表数据，会在文档中循环生成多行
     *
     * @param args 命令行参数
     * @throws Exception 文件操作异常
     */
    public static void main(String[] args) throws Exception {
        // 格式化当前时间
        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String curTime = format.format(new Date());

        // 准备数据 Map
        Map<String, Object> map = new HashMap<String, Object>();

        // 准备列表数据（示例中被注释，可用于循环生成表格行）
        List<Map<String, Object>> mapList = new ArrayList<Map<String, Object>>();
        // 示例：添加多条纳税记录到列表
        // Map<String, Object> map1 = new HashMap<String, Object>();
        // map1.put("type", "个人所得税");
        // map1.put("presum", "1580");
        // map1.put("thissum", "1750");
        // map1.put("curmonth", "1-11月");
        // map1.put("now", curTime);
        // mapList.add(map1);

        map.put("taxlist", mapList);  // 列表数据（当前为空）
        map.put("totalpreyear", "2660");  // 去年总额
        map.put("totalthisyear", "3400");  // 今年总额

        // 单条记录数据
        map.put("type", "增值税");
        map.put("presum", "1080");
        map.put("thissum", "1650");
        map.put("curmonth", "1-11月");
        map.put("now", curTime);

        // 基于模板导出 Word 文档（.docx 格式）
        XWPFDocument document = WordExportUtil.exportWord07(TEMPLATE_PATH + "纳税信息.docx", map);

        // 创建保存目录
        File savefile = new File("D:\\excel");
        if (!savefile.exists()) {
            savefile.mkdirs();
        }

        // 保存生成的 Word 文档
        FileOutputStream fos = new FileOutputStream("D:\\excel\\纳税信息new.docx");
        document.write(fos);
        fos.close();

        System.out.println("Word 文档导出完成！文件保存在：D:\\excel\\纳税信息new.docx");
    }
}
