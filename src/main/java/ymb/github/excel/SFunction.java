package ymb.github.excel;

import java.io.Serializable;
import java.lang.invoke.SerializedLambda;
import java.lang.reflect.Method;
import java.util.function.Function;

/**
 * 序列化的Function
 * @author WuLiao
 */
@SuppressWarnings("AlibabaClassNamingShouldBeCamel")
public interface SFunction<T, R> extends Function<T, R>, Serializable {

    /**
     * 通过属性的get方法获取属性名
     * @param fn  SFunction对象
     * @param <T> 类型
     * @return 属性名
     * @throws ReflectiveOperationException 反射异常
     */
    static <T, R> String getFieldName(SFunction<T, R> fn) throws ReflectiveOperationException {
        // 从function取出序列化方法
        Method writeReplaceMethod = fn.getClass().getDeclaredMethod("writeReplace");

        // 从序列化方法取出序列化的lambda信息
        boolean isAccessible = writeReplaceMethod.isAccessible();
        writeReplaceMethod.setAccessible(true);
        SerializedLambda serializedLambda = (SerializedLambda) writeReplaceMethod.invoke(fn);
        writeReplaceMethod.setAccessible(isAccessible);

        // 从lambda信息取出method、field、class等
        String fieldName = serializedLambda.getImplMethodName().substring("get".length());
        fieldName = fieldName.replaceFirst(fieldName.charAt(0) + "", (fieldName.charAt(0) + "").toLowerCase());
        return fieldName;
    }
}
