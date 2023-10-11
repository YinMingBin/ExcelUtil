package ymb.github.excel.annotation;

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
public @interface ExcelClass {
    /** 同 title */
    String value() default "";
    /** 标题（仅最上层有效） */
    String title() default "";
    /** 背景色（默认：无背景） */
    IndexedColors background() default IndexedColors.AUTOMATIC;
    /** 背景格式（默认：SOLID_FOREGROUND） */
    FillPatternType pattern() default FillPatternType.SOLID_FOREGROUND;
    /** 标题水平对齐方式（默认居中） **/
    HorizontalAlignment horizontal() default HorizontalAlignment.CENTER;
    /** 标题垂直对齐方式（默认居中） **/
    VerticalAlignment vertical() default VerticalAlignment.CENTER;
    /** 字体大小（默认：20） */
    short fontSize() default 20;
}
