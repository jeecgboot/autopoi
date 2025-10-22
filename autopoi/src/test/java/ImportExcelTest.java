import org.apache.commons.lang3.builder.ReflectionToStringBuilder;
import org.jeecgframework.poi.excel.ExcelImportUtil;
import org.jeecgframework.poi.excel.entity.ImportParams;
import vo.TestDateEntity;

import java.io.File;
import java.util.List;

/**
 * @Description: TODO
 * @author: scott
 * @date: 2020年09月16日 11:46
 */
public class ImportExcelTest {
    private static final String TEMPLATE_PATH = System.getProperty("user.dir") + File.separator + "autopoi" + File.separator + "src" + File.separator + "test" + File.separator + "resources" + File.separator + "templates" + File.separator;

    public static void main(String[] args) throws Exception {
        ImportParams params = new ImportParams();
        params.setTitleRows(1);
        params.setHeadRows(1);
        File importFile = new File(TEMPLATE_PATH + "ExcelImportDateTest.xlsx");
        List<TestDateEntity> list = ExcelImportUtil.importExcel(importFile, TestDateEntity.class, params);
        System.out.println(list.size());
        System.out.println(ReflectionToStringBuilder.toString(list.get(1)));
    }
}
