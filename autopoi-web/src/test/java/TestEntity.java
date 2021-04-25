import org.jeecgframework.poi.excel.annotation.Excel;

/**
 * @Description: TODO
 * @author: lsq
 * @date: 2021年02月02日 16:31
 */
public class TestEntity {
    @Excel(name = "姓名", width = 15)
    private String name;
    @Excel(name = "年龄", width = 15)
    private Integer age;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }
}
