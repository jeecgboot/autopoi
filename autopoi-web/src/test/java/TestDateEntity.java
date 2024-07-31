import org.jeecgframework.poi.excel.annotation.Excel;

import java.time.LocalDate;
import java.time.LocalDateTime;

/**
 * @Description: TODO
 * @author: lsq
 * @date: 2024年07月31日 10:31
 */
public class TestDateEntity {
    @Excel(name = "localdate", format = "yyyy-MM-dd")
    private LocalDate localDate;

    @Excel(name = "localdatetime", format = "yyyy-MM-dd")
    private LocalDateTime localDateTime;


    public LocalDate getLocalDate() {
        return localDate;
    }

    public void setLocalDate(LocalDate localDate) {
        this.localDate = localDate;
    }

    public LocalDateTime getLocalDateTime() {
        return localDateTime;
    }

    public void setLocalDateTime(LocalDateTime localDateTime) {
        this.localDateTime = localDateTime;
    }
}
