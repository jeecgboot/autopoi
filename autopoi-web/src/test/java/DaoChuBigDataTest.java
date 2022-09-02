import org.apache.poi.ss.usermodel.Workbook;
import org.jeecgframework.poi.excel.ExcelExportUtil;
import org.jeecgframework.poi.excel.entity.ExportParams;
import org.jeecgframework.poi.handler.inter.IExcelExportServer;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @Description: 大数据导出示例
 * @author: liusq
 * @date: 2022年1月4日
 */
public class DaoChuBigDataTest {
    /**
     * 导出测试方法
     * 参考：http://doc.wupaas.com/docs/easypoi/easypoi-1c10lbsojh62f
     * 参考：https://blog.csdn.net/weixin_45214729/article/details/118552415
     * @author liusq
     * @date: 2022年1月4日
     * @throws Exception
     */
    public static void bigDataExport() throws Exception {
        Workbook workbook = null;
        List<TestEntity> aList = new ArrayList<TestEntity>();
        ExportParams exportParams = new ExportParams();
        Date   start    = new Date();
        //模拟100w数据
        for(int j=0;j<1000000;j++){
            TestEntity testEntity = new TestEntity();
            testEntity.setName("李四"+j);
            testEntity.setAge(j);
            aList.add(testEntity);
        }
        //分别是 totalPage是总页数，pageSize 页码长度
        int totalPage = (aList.size() / 100000) + 1;
        int pageSize = 100000;
        /**
         * params:（表格标题属性）筛选条件，sheet值
         * TestEntity：表格的实体类
         */
        workbook = ExcelExportUtil.exportBigExcel(exportParams,TestEntity.class, new IExcelExportServer() {
            /**
             * obj 就是下面的totalPage，限制条件
             * page 是页数，他是在分页进行文件转换，page每次+1
             */
            @Override
            public List<Object> selectListForExcelExport(Object obj, int page) {
                //很重要！！这里面整个方法体，其实就是将所有的数据aList分批返回处理
                //分批的方式很多，我直接用了subList。然后 每批不能太大。我试了30000一批，
                //特别注意，最好每次10000条，否则，可能有内存溢出风险
                if (page > totalPage) {
                    return null;
                }
                // fromIndex开始索引，toIndex结束索引
                int fromIndex = (page - 1) * pageSize;
                int toIndex = page != totalPage ? fromIndex + pageSize :aList.size();
                //不是空时：一直循环运行selectListForExcelExport。每次返回1万条数据。
                List<Object> list = new ArrayList<>();
                list.addAll(aList.subList(fromIndex, toIndex));
                return list;
            }
        }, totalPage);
        File savefile = new File("D:/excel/");
        if (!savefile.exists()) {
            savefile.mkdirs();
        }
        FileOutputStream fos = new FileOutputStream("D:/excel/ExcelExportBigData.bigDataExport.xlsx");
        workbook.write(fos);
        fos.close();
        System.out.println("耗时(秒)："+((new Date().getTime() - start.getTime())/ 1000));
    }

    public static void main(String[] args) throws Exception {
       bigDataExport();
    }
}
