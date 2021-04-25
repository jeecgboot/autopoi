import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.jeecgframework.poi.word.WordExportUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Description: TODO
 * @author: scott
 * @date: 2020年09月16日 11:46
 */
public class DaoChuWordTest {
    private static final String basePath = "G:\\needtodeplay\\autopoi-framework-sy-4.0\\autopoi-web\\src\\test\\resources\\templates\\";

    public static void main(String[] args) throws Exception {
        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String curTime = format.format(new Date());

        Map<String, Object> map = new HashMap<String, Object>();
        List<Map<String, Object>> mapList = new ArrayList<Map<String, Object>>();
//        Map<String, Object> map1 = new HashMap<String, Object>();
//        map1.put("type", "个人所得税");
//        map1.put("presum", "1580");
//        map1.put("thissum", "1750");
//        map1.put("curmonth", "1-11月");
//        map1.put("now", curTime);
//        mapList.add(map1);
//        Map<String, Object> map2 = new HashMap<String, Object>();
//        map2.put("type", "增值税");
//        map2.put("presum", "1080");
//        map2.put("thissum", "1650");
//        map2.put("curmonth", "1-11月");
//        map2.put("now", curTime);
//        mapList.add(map2);
        map.put("taxlist", mapList);
        map.put("totalpreyear", "2660");
        map.put("totalthisyear", "3400");

        map.put("type", "增值税");
        map.put("presum", "1080");
        map.put("thissum", "1650");
        map.put("curmonth", "1-11月");
        map.put("now", curTime);

        //单列
        XWPFDocument document = WordExportUtil.exportWord07(basePath + "纳税信息.docx", map);
        File savefile = new File("D:\\poi");
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream(basePath + "纳税信息new.docx");
        document.write(fos);
        fos.close();
    }
}
