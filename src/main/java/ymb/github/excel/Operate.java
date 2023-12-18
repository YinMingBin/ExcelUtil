package ymb.github.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.List;
import java.util.function.BiConsumer;
import java.util.function.Consumer;

/**
 * @author WuLiao
 */
@SuppressWarnings({"UnusedReturnValue", "unchecked"})
public interface Operate<T, R> {

    /**
     * 设置数据源
     *
     * @param data 数据
     * @return this
     */
    R setData(List<T> data);

    /**
     * 设置表头的字体大小
     *
     * @param titleSize 字体大小
     * @return this
     */
    R setTitleSize(short titleSize);

    /**
     * 设置单元格（Cell）的字体大小
     * @param fontSize 字体大小
     * @return this
     */
    R setFontSize(short fontSize);

    /**
     * 设置表头的行高
     * @param titleHeight 行高
     * @return this
     */
    R setTitleHeight(short titleHeight);

    /**
     * 设置（Row）行高
     * @param rowHeight 行高
     * @return this
     */
    R setRowHeight(short rowHeight);

    /**
     * 设置列宽
     * @param columnWidth 列宽
     * @return this
     */
    R setColumnWidth(int columnWidth);

    /**
     * 设置表头样式
     * @param titleStyleFun (CellStyle) -> void
     * @return this
     */
    R setTitleStyle(Consumer<CellStyle> titleStyleFun);

    /**
     * 设置单元格（Cell）样式
     * @param cellStyleFun (CellStyle) -> void
     * @return this
     */
    R setCellStyle(Consumer<CellStyle> cellStyleFun);

    /**
     * 操作某一列单元格（Cell）的样式 (设置Cell时调用)
     * @param index 列索引
     * @param cellStyle (CellStyle, RowData) -> void
     * @return this
     */
    R operateCellStyle(int index, BiConsumer<CellStyle, Object> cellStyle);

    /**
     * 操作表头，每次设置表头之后执行
     * @param operateTitle (Cell) -> void
     * @return this
     */
    R operateTitle(Consumer<SXSSFCell> operateTitle);

    /**
     * 操作单元格（Cell），每次设置数据之后执行
     * @param operateCell (Cell, RowData) -> void
     * @return this
     */
    R operateCell(BiConsumer<SXSSFCell, Object> operateCell);

    /**
     * 操作Row，每次设置完一行数据之后执行
     * @param operateRow (Row, RowData) -> void
     * @return this
     */
    R operateRow(BiConsumer<SXSSFRow, Object> operateRow);

    /**
     * 操作Sheet，在数据生成完之后执行
     * @param operateSheet (Sheet, dataList) -> void
     * @return this
     */
    R operateSheet(BiConsumer<SXSSFSheet, List<T>> operateSheet);

    /**
     * 设置列
     * @param functions 字段的get方法（不定项参数）
     * @return this
     */
    R settingColumn(SFunction<T, ?>... functions);

    /**
     * 设置列
     * @param function 字段的get方法
     * @param columnClass 列属性
     * @return this
     */
    R settingColumn(SFunction<T, ?> function, ExcelColumnClass columnClass);

    /**
     * 启用自适应列宽
     * @return this
     */
    R autoColumnWidth();
}