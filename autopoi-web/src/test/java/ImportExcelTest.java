import org.apache.commons.lang3.builder.ReflectionToStringBuilder;
import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.ExcelImportUtil;
import org.jeecgframework.poi.excel.entity.ImportParams;
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
public class ImportExcelTest {
    private static final String basePath = "D:\\idea_project_2023\\autopoi_lsq\\autopoi-web\\src\\test\\resources\\templates\\";

    public static void main(String[] args) throws Exception {
        ImportParams params = new ImportParams();
        params.setTitleRows(1);
        params.setHeadRows(1);
        File importFile = new File(basePath+"ExcelImportDateTest.xlsx");
        List<TestDateEntity> list = ExcelImportUtil.importExcel(importFile, TestDateEntity.class, params);
        System.out.println(list.size());
        System.out.println(ReflectionToStringBuilder.toString(list.get(1)));
    }
}
