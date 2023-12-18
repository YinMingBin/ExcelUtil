package ymb.github.excel;

/**
 * Cell类型
 * @author WuLiao
 */

public enum CellType {
    /** 数字（byte, short, int, long, float, double） */
    NUMBER(org.apache.poi.ss.usermodel.CellType.NUMERIC),
    /** true/false */
    BOOLEAN(org.apache.poi.ss.usermodel.CellType.BOOLEAN),
    /** 字符（char, String） */
    STRING(org.apache.poi.ss.usermodel.CellType.STRING),
    /** 公式 */
    FORMULA(org.apache.poi.ss.usermodel.CellType.FORMULA),
    /** 空 */
    BLANK(org.apache.poi.ss.usermodel.CellType.BLANK),
    /** 对象（读取Object中添加了ExcelColumn注解的属性） */
    OBJECT(null),
    /** 集合（不需要设置，程序会自动判断：Collection.class.isAssignableFrom(field.getType())） */
    LIST(null);

    private final org.apache.poi.ss.usermodel.CellType cellType;

    CellType(org.apache.poi.ss.usermodel.CellType cellType) {
        this.cellType = cellType;
    }

    public org.apache.poi.ss.usermodel.CellType getCellType() {
        return cellType;
    }
}
