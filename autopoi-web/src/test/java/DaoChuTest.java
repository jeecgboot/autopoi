import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Description: TODO
 * @author: scott
 * @date: 2020年09月16日 11:46
 */
public class DaoChuTest {
    private static final String basePath = "G:\\needtodeplay\\autopoi-framework-sy-4.0\\autopoi-web\\src\\test\\resources\\templates\\";

    public static TemplateExportParams getTemplateParams(String name) {
        return new TemplateExportParams(basePath + name + ".xlsx");
    }

    public static Workbook test(String name) {
        TemplateExportParams params = getTemplateParams(name);
        Map<String, Object> map = new HashMap<String, Object>();
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
        map.put("autoList", listMap);
        Workbook workbook = ExcelExportUtil.exportExcel(params, map);
        return workbook;
    }

    public static void main(String[] args) throws IOException {
        String temName = "test";
        String temNameNextM = "testNextMarge";
        Workbook workbook = test(temNameNextM);
        File savefile = new File(basePath);
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream(basePath + "testNew.xlsx");
        workbook.write(fos);
        fos.close();
    }
}
