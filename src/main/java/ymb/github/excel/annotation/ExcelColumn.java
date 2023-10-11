package ymb.github.excel.annotation;

import org.apache.poi.ss.usermodel.*;

import java.lang.annotation.*;

/**
 * @author YinMingBin
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
@AllFieldColumn
public @interface ExcelColumn {
    /** 同 title */
    String value() default "";
    /** 表头 */
    String title() default "";
    /** 所属位置 */
    short index() default -1;
    /** 数据类型 */
    CellType type() default CellType.STRING;
    /** 数据格式 */
    String format() default "";
    /** 列宽 */
    int width() default 10;
    /** 是否自动换行（默认：否） */
    boolean wrapText() default false;
    /** 背景色（默认：无背景） */
    IndexedColors background() default IndexedColors.AUTOMATIC;
    /** 字体颜色（默认：黑色） */
    IndexedColors color() default IndexedColors.BLACK;
    /** 字体大小（默认：11） */
    short size() default 0;
    /** 背景格式（默认：SOLID_FOREGROUND） */
    FillPatternType pattern() default FillPatternType.SOLID_FOREGROUND;
    /** 数据水平对齐方式（默认靠左） **/
    HorizontalAlignment horizontal() default HorizontalAlignment.LEFT;
    /** 数据垂直对其方式（默认靠居中） **/
    VerticalAlignment vertical() default VerticalAlignment.CENTER;
}
