package ymb.github.excel;

import java.io.Serializable;
import java.util.function.Function;

/**
 * @author WuLiao
 */
@SuppressWarnings("AlibabaClassNamingShouldBeCamel")
public interface SFunction<T, R> extends Function<T, R>, Serializable {
}
