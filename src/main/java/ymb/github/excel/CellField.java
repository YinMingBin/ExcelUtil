package ymb.github.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;

import java.util.List;
import java.util.function.BiConsumer;
import java.util.function.Function;

/**
 * @author YinMingBin
 */
class CellField {
    private int index;
    private String title;
    private CellType cellType;
    private Class<?> fieldType;
    private CellStyle cellStyle;
    private Function<Object, Object> valueFun;
    private BiConsumer<Object, Object> settingFun;
    private List<CellField> cellFields;
    private int width;

    void setTitle(String title, String title2) {
        if (title == null || title.isEmpty()) {
            title = title2;
        }
        this.title = title;
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public String getTitle() {
        return title;
    }

    public CellType getCellType() {
        return cellType;
    }

    public void setCellType(CellType cellType) {
        this.cellType = cellType;
    }

    public Class<?> getFieldType() {
        return fieldType;
    }

    public void setFieldType(Class<?> fieldType) {
        this.fieldType = fieldType;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public Function<Object, Object> getValueFun() {
        return valueFun;
    }

    public void setValueFun(Function<Object, Object> valueFun) {
        this.valueFun = valueFun;
    }

    public BiConsumer<Object, Object> getSettingFun() {
        return settingFun;
    }

    public void setSettingFun(BiConsumer<Object, Object> settingFun) {
        this.settingFun = settingFun;
    }

    public List<CellField> getCellFields() {
        return cellFields;
    }

    public void setCellFields(List<CellField> cellFields) {
        this.cellFields = cellFields;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }
}