import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelToHtmlUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;

public class ExcelToHtmlTest {

    private static final String TEMPLATE_PATH = System.getProperty("user.dir") + File.separator + "autopoi" + File.separator + "src" + File.separator + "test" + File.separator + "resources" + File.separator + "templates" + File.separator;

    public static void main(String[] args) throws Exception {
        File file = new File(TEMPLATE_PATH + "专项支出用款申请书.xls");
        Workbook wb = new HSSFWorkbook(new FileInputStream(file));
        String     html = ExcelToHtmlUtil.toTableHtml(wb);
        FileWriter fw   = new FileWriter("D:/excel/专项支出用款申请书.html");
        fw.write(html);
        fw.close();
    }
}
