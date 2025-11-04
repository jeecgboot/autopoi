import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.consts.ImageScaleMode;
import org.jeecgframework.poi.entity.ImageEntity;
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
 * 图片缩放功能测试类
 * 测试模板导出中的图片缩放功能
 * for [issues/8892] AutoPoi ImageEntity建议添加scale属性，控制图片导出缩放模式
 * 
 * @author chenrui
 * @date 2025-10-28
 */
public class TestImageScale {
    
    // 模板文件路径 - 使用相对路径，确保所有人都能使用
    private static final String TEMPLATE_PATH = "autopoi/src/test/resources/templates/";

    // 测试图片路径 - 使用相对路径，图片已放在 resources/templates 下
    private static final String TEST_IMAGE_PATH = "autopoi/src/test/resources/templates/dakytot.jpeg";
    
    /**
     * 获取模板导出参数配置
     */
    public static TemplateExportParams getTemplateParams(String name) {
        return new TemplateExportParams(TEMPLATE_PATH + name + ".xlsx");
    }
    
    /**
     * 测试图片缩放功能
     * 测试3种不同的缩放模式：
     * - imageLSTC: 拉伸填充 (scaleMode = ImageScaleMode.STRETCH)
     * - imageDBL: 等比例缩放适应 (scaleMode = ImageScaleMode.FIT) 
     * - imageYT: 不缩放（原始大小） (scaleMode = ImageScaleMode.ORIGINAL)
     */
    public static Workbook testImageScaling(String templateName) {
        try {
            System.out.println("开始测试图片缩放功能，模板: " + templateName);
            TemplateExportParams params = getTemplateParams(templateName);
            System.out.println("模板路径: " + params.getTemplateUrl());
            
            Map<String, Object> map = new HashMap<String, Object>();

            // 构造列表数据（3条记录）
            List<Map<String, Object>> listMap = new ArrayList<Map<String, Object>>();
            for (int i = 0; i < 3; i++) {
                Map<String, Object> lm = new HashMap<String, Object>();
                lm.put("name", "用户" + (i + 1));
                lm.put("id", i + 1);
                
                // 1. 拉伸填充
                ImageEntity imageLSTC = new ImageEntity();
                imageLSTC.setHeight(200);
                imageLSTC.setWidth(300);
                imageLSTC.setUrl(TEST_IMAGE_PATH);
                imageLSTC.setScaleModeEnum(ImageScaleMode.STRETCH); // 拉伸填充
                lm.put("imageLSTC", imageLSTC);
                
                // 2. 等比例缩放适应
                ImageEntity imageDBL = new ImageEntity();
                imageDBL.setHeight(200);
                imageDBL.setWidth(300);
                imageDBL.setUrl(TEST_IMAGE_PATH);
                imageDBL.setScaleModeEnum(ImageScaleMode.FIT); // 等比例缩放适应
                lm.put("imageDBL", imageDBL);
                
                // 3. 不缩放（原始大小）
                ImageEntity imageYT = new ImageEntity();
                imageYT.setHeight(200);
                imageYT.setWidth(300);
                imageYT.setUrl(TEST_IMAGE_PATH);
                imageYT.setScaleModeEnum(ImageScaleMode.ORIGINAL); // 不缩放（原始大小）
                lm.put("imageYT", imageYT);
                
                listMap.add(lm);
            }

            // 将列表数据放入 map，对应模板中的 {{$fe: autoList}} 循环标记
            map.put("autoList", listMap);
            System.out.println("数据准备完成，开始导出Excel...");

            // 基于模板和数据生成 Excel
            Workbook workbook = ExcelExportUtil.exportExcel(params, map);
            System.out.println("Excel导出完成，workbook: " + (workbook != null ? "成功" : "失败"));
            return workbook;
        } catch (Exception e) {
            System.err.println("测试图片缩放功能失败: " + e.getMessage());
            e.printStackTrace();
            return null;
        }
    }
    
    /**
     * 主方法：执行图片缩放测试
     */
    public static void main(String[] args) throws IOException {
        System.out.println("=== 图片缩放功能测试开始 ===");
        
        String templateName = "testImageScale";  // 使用test模板

        String outputDir = "/Users/chenrui/Downloads";

        try {
            // 测试1：本地图片缩放
            System.out.println("测试1：本地图片缩放功能");
            Workbook workbook1 = testImageScaling(templateName);

            File saveDir = new File(outputDir);
            if (!saveDir.exists()) {
                saveDir.mkdirs();
            }

            FileOutputStream fos1 = new FileOutputStream(outputDir + "/ImageScaleTest_Local.xlsx");
            workbook1.write(fos1);
            fos1.close();
            workbook1.close();

            System.out.println("=== 图片缩放功能测试完成 ===");
            System.out.println("测试说明：");
            System.out.println("- imageLSTC: 拉伸填充 (scaleMode = ImageScaleMode.STRETCH)");
            System.out.println("- imageDBL: 等比例缩放适应 (scaleMode = ImageScaleMode.FIT)");
            System.out.println("- imageYT: 不缩放（原始大小） (scaleMode = ImageScaleMode.ORIGINAL)");

        } catch (Exception e) {
            System.err.println("❌ 测试失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
