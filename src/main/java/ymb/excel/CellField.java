package ymb.excel;

import lombok.Data;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.List;
import java.util.function.Function;

/**
 * @author YinMingBin
 */
@Data
class CellField {
    private int index;
    private String title;
    private CellType cellType;
    private XSSFCellStyle cellStyle;
    private Function<Object, Object> valueFun;
    private List<CellField> cellFields;
    private int width;

    void setTitle(String title, String title2) {
        if (title == null || title.isEmpty()) {
            title = title2;
        }
        this.title = title;
    }
}