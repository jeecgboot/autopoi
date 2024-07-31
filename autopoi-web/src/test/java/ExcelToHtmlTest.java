import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelToHtmlUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;

public class ExcelToHtmlTest {

    private static final String basePath = "D:\\idea_project_2023\\autopoi_lsq\\autopoi-web\\src\\test\\resources\\templates\\";

    public static void main(String[] args) throws Exception {
        File file = new File(basePath + "专项支出用款申请书.xls");
        Workbook wb = new HSSFWorkbook(new FileInputStream(file));
        String     html = ExcelToHtmlUtil.toTableHtml(wb);
        FileWriter fw   = new FileWriter("D:/home/excel/专项支出用款申请书.html");
        fw.write(html);
        fw.close();
    }
}
