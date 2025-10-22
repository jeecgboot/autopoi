import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelToHtmlUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;

/**
 * @Description: Excel 转 HTML 测试
 * 测试功能：
 * 1. 测试将 Excel 文件转换为 HTML 表格格式
 * 2. 验证 Excel 样式（字体、颜色、边框等）在 HTML 中的呈现
 * 3. 测试复杂表格结构（合并单元格、多行表头）的转换效果
 *
 * 应用场景：
 * - 在网页中预览 Excel 内容
 * - 将 Excel 报表转换为 HTML 格式发送邮件
 * - 在线展示 Excel 数据，无需下载文件
 *
 * @author: autopoi
 */
public class ExcelToHtmlTest {

    // 模板文件路径：项目根目录/autopoi/src/test/resources/templates/
    private static final String TEMPLATE_PATH = System.getProperty("user.dir") + File.separator + "autopoi" + File.separator + "src" + File.separator + "test" + File.separator + "resources" + File.separator + "templates" + File.separator;

    /**
     * 主方法：执行 Excel 转 HTML 测试
     *
     * 测试内容：
     * - 读取 .xls 格式的 Excel 文件（专项支出用款申请书.xls）
     * - 将 Excel 内容转换为 HTML 表格格式
     * - 保留原 Excel 的样式和布局
     * - 将生成的 HTML 保存到文件
     *
     * 转换特性：
     * - 自动处理合并单元格
     * - 保留字体样式（大小、颜色、粗体等）
     * - 保留单元格边框和背景色
     * - 保持表格布局和对齐方式
     *
     * 预期结果：
     * - 在 D:/excel/ 目录下生成 专项支出用款申请书.html 文件
     * - HTML 文件可在浏览器中打开查看
     * - 表格样式与原 Excel 文件基本一致
     *
     * @param args 命令行参数
     * @throws Exception 文件读取或转换异常
     */
    public static void main(String[] args) throws Exception {
        // 读取 Excel 文件
        File file = new File(TEMPLATE_PATH + "专项支出用款申请书.xls");
        Workbook wb = new HSSFWorkbook(new FileInputStream(file));

        // 将 Excel 转换为 HTML 表格
        String html = ExcelToHtmlUtil.toTableHtml(wb);

        // 保存 HTML 文件
        FileWriter fw = new FileWriter("D:/excel/专项支出用款申请书.html");
        fw.write(html);
        fw.close();

        System.out.println("Excel 转 HTML 完成！文件保存在：D:/excel/专项支出用款申请书.html");
    }
}
