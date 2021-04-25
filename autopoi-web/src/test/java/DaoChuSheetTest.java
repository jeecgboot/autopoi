import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.excel.entity.TemplateExportParams;
import org.jeecgframework.poi.excel.entity.enmus.ExcelType;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Description: 参考文档：http://doc.jeecg.com/2044223
 * @author: scott
 * @date: 2020年09月16日 11:46
 */
public class DaoChuSheetTest {
    private static final String basePath = "G:\\needtodeplay\\autopoi-framework-sy-4.0\\autopoi-web\\src\\test\\resources\\templates\\";

    public static ExportParams getExportParams(String name) {
        return  new ExportParams(name,name,ExcelType.XSSF);
    }
    public static Workbook test() {
        /**
         * 多个Map
         * title:对应表格Title
         * entity：对应表格对应实体
         * data：Collection 数据
         */
        List<Map<String, Object>> listMap = new ArrayList<Map<String, Object>>();
        for(int i=0;i<3;i++){
            Map<String, Object> map = new HashMap<String, Object>();
            map.put("title", getExportParams("测试"+i));//表格Title
            map.put("entity",TestEntity.class);//表格对应实体
            List<TestEntity> ls=new ArrayList<TestEntity> ();
            for(int j=0;j<10;j++){
                TestEntity testEntity = new TestEntity();
                testEntity.setName("张三"+i+j);
                testEntity.setAge(18+i+j);
                ls.add(testEntity);
            }
            List<Map> ls2=new ArrayList<Map> ();
            for(int j=0;j<10;j++){
                Map map1 = new HashMap();
                map1.put("name","李四"+i+j);
                map1.put("age",18+i+j);
                ls2.add(map1);
            }
            map.put("data", ls);//ls or ls2
            listMap.add(map);
        }
        Workbook workbook = ExcelExportUtil.exportExcel(listMap, ExcelType.XSSF);
        return workbook;
    }

    public static void main(String[] args) throws IOException {
        Workbook workbook = test();
        File savefile = new File(basePath);
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream(basePath + "testSheet.xlsx");
        workbook.write(fos);
        fos.close();
    }
}
