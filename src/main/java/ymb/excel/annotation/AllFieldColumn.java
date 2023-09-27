package ymb.excel.annotation;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.*;

/**
 * @author YinMingBin
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
@Documented
public @interface AllFieldColumn {
    /** 列宽 */
    int width() default 10;
    /** 是否自动换行（默认：否） */
    boolean wrapText() default false;
    /** 背景色（默认：无背景） */
    IndexedColors background() default IndexedColors.AUTOMATIC;
    /** 背景格式（默认：SOLID_FOREGROUND） */
    FillPatternType pattern() default FillPatternType.SOLID_FOREGROUND;
    /** 数据对齐方式（默认靠左） **/
    HorizontalAlignment horizontal() default HorizontalAlignment.LEFT;
    /** 数据垂直对其方式（默认靠居中） **/
    VerticalAlignment vertical() default VerticalAlignment.CENTER;
}
