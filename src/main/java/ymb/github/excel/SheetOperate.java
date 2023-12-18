package ymb.github.excel;

import javafx.util.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import ymb.github.excel.annotation.AllFieldColumn;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.util.*;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.stream.Collectors;

/**
 * @author YinMingBin
 */
@SuppressWarnings({"unused", "UnusedReturnValue"})
public final class SheetOperate<T> implements Operate<T, SheetOperate<T>>{
    private SXSSFWorkbook workbook;
    private SXSSFSheet sheet;
    private String sheetName;
    private List<T> data;
    private Consumer<SXSSFCell> operateTitle;
    private BiConsumer<SXSSFCell, Object> operateCell;
    private BiConsumer<SXSSFRow, Object> operateRow;
    private BiConsumer<SXSSFSheet, List<T>> operateSheet;
    private List<CellField> fields;
    private final Class<T> tClass;
    private short titleSize = 12;
    private short fontSize = 11;
    private float titleHeight = 25;
    private float rowHeight = 20;
    private int columnWidth = 10;
    private Map<Integer, BiConsumer<CellStyle, Object>> cellStyleFunMap;
    private Consumer<CellStyle> cellStyleFun = cell -> {};
    private Consumer<CellStyle> titleStyleFun = cell -> {};
    private boolean autoColumnWidth = false;
    private List<Pair<SFunction<T, ?>, ExcelColumnClass>> columnFunctions;

    private SheetOperate(Class<T> tClass) {
        this.tClass = tClass;
    }

    private SheetOperate(Class<T> tClass, String sheetName) {
        this.tClass = tClass;
        this.sheetName = sheetName;
    }

    public static <R> SheetOperate<R> create(Class<R> tClass) {
        return new SheetOperate<>(tClass);
    }

    public static <R> SheetOperate<R> create(Class<R> tClass, String sheetName) {
        return new SheetOperate<>(tClass, sheetName);
    }

    void setWorkbook(SXSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    /**
     * 设置数据源
     * @param data 数据源
     * @return this
     */
    @Override
    public SheetOperate<T> setData(List<T> data) {
        this.data = data;
        return this;
    }

    /**
     * 设置表头的字体大小
     * @param titleSize 字体大小
     * @return this
     */
    @Override
    public SheetOperate<T> setTitleSize(short titleSize) {
        this.titleSize = titleSize;
        return this;
    }

    /**
     * 设置Cell的字体大小
     * @param fontSize 字体大小
     * @return this
     */
    @Override
    public SheetOperate<T> setFontSize(short fontSize) {
        this.fontSize = fontSize;
        return this;
    }

    /**
     * 设置表头的行高
     * @param titleHeight 行高
     * @return this
     */
    @Override
    public SheetOperate<T> setTitleHeight(short titleHeight) {
        this.titleHeight = titleHeight;
        return this;
    }

    /**
     * 设置(Row)行高
     * @param rowHeight 行高
     * @return this
     */
    @Override
    public SheetOperate<T> setRowHeight(short rowHeight) {
        this.rowHeight = rowHeight;
        return this;
    }

    /**
     * 设置列宽
     * @param columnWidth 列宽
     * @return this
     */
    @Override
    public SheetOperate<T> setColumnWidth(int columnWidth) {
        this.columnWidth = columnWidth;
        return this;
    }

    /**
     * 设置表头样式
     * @param titleStyleFun (CellStyle) -> void
     * @return this
     */
    @Override
    public SheetOperate<T> setTitleStyle(Consumer<CellStyle> titleStyleFun) {
        this.titleStyleFun = titleStyleFun;
        return this;
    }

    /**
     * 设置Cell样式
     * @param cellStyleFun (CellStyle) -> void
     * @return this
     */
    @Override
    public SheetOperate<T> setCellStyle(Consumer<CellStyle> cellStyleFun) {
        this.cellStyleFun = cellStyleFun;
        return this;
    }

    /**
     * 操作某一列单元格（Cell）的样式 (设置Cell时调用)
     * @param index 列索引
     * @param cellStyle (CellStyle, rowData) -> void
     * @return this
     */
    @Override
    public SheetOperate<T> operateCellStyle(int index, BiConsumer<CellStyle, Object> cellStyle) {
        if (cellStyleFunMap == null) {
            cellStyleFunMap = new HashMap<>(5);
        }
        cellStyleFunMap.put(index, cellStyle);
        return this;
    }

    /**
     * 操作表头，每次设置表头之后执行
     * @param operateTitle (Cell) -> void
     * @return this
     */
    @Override
    public SheetOperate<T> operateTitle(Consumer<SXSSFCell> operateTitle) {
        this.operateTitle = operateTitle;
        return this;
    }

    /**
     * 操作单元格（Cell），每次设置数据之后执行
     * @param operateCell (Cell, RowData) -> void
     * @return this
     */
    @Override
    public SheetOperate<T> operateCell(BiConsumer<SXSSFCell, Object> operateCell) {
        this.operateCell = operateCell;
        return this;
    }

    /**
     * 操作Row，每次设置完一行数据之后执行
     * @param operateRow (Row, RowData) -> void
     * @return this
     */
    @Override
    public SheetOperate<T> operateRow(BiConsumer<SXSSFRow, Object> operateRow) {
        this.operateRow = operateRow;
        return this;
    }

    /**
     * 操作Sheet，在数据生成完之后执行
     * @param operateSheet (Sheet, dataList) -> void
     * @return this
     */
    @Override
    public SheetOperate<T> operateSheet(BiConsumer<SXSSFSheet, List<T>> operateSheet) {
        this.operateSheet = operateSheet;
        return this;
    }

    /**
     * 设置列
     * @param functions 字段的get方法（不定项参数）
     * @return this
     */
    @SafeVarargs
    @Override
    public final SheetOperate<T> settingColumn(SFunction<T, ?>... functions) {
        if (functions == null) {
            return this;
        }

        if (columnFunctions == null) {
            columnFunctions = new ArrayList<>(5);
        }

        for (SFunction<T, ?> function : functions) {
            columnFunctions.add(new Pair<>(function, null));
        }

        return this;
    }

    /**
     * 设置列
     * @param function 字段的get方法
     * @param columnClass 列属性
     * @return this
     */
    @Override
    public SheetOperate<T> settingColumn(SFunction<T, ?> function, ExcelColumnClass columnClass) {
        if (function != null) {

            if (columnFunctions == null) {
                columnFunctions = new ArrayList<>(5);
            }

            columnFunctions.add(new Pair<>(function, columnClass));
        }
        return this;
    }

    @Override
    public SheetOperate<T> autoColumnWidth() {
        this.autoColumnWidth = true;
        return this;
    }

    SXSSFSheet getSheet() {
        if (sheet == null) {
            sheet = sheetName == null ? workbook.createSheet() : workbook.createSheet(sheetName);
            sheetName = sheet.getSheetName();
        }
        return sheet;
    }

    void clearSheet() {
        this.sheet = null;
    }

    CellStyle getCellStyle() {
        CellStyle cellStyle = workbook.createCellStyle();
        // 设置字体
        Font font = workbook.createFont();
        font.setFontHeightInPoints(getFontSize());
        cellStyle.setFont(font);
        // 设置边框
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

        cellStyleFun.accept(cellStyle);
        return cellStyle;
    }

    CellStyle getTitleStyle() {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        // 设置字体
        Font font = workbook.createFont();
        font.setFontHeightInPoints(getTitleSize());
        cellStyle.setFont(font);
        // 设置边框
        cellStyle.setBorderTop(BorderStyle.MEDIUM);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderBottom(BorderStyle.MEDIUM);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderLeft(BorderStyle.MEDIUM);
        cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyle.setBorderRight(BorderStyle.MEDIUM);
        cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        titleStyleFun.accept(cellStyle);
        return cellStyle;
    }

    List<CellField> getFields() {
        if (this.fields == null) {
            if (columnFunctions == null) {
                this.fields = getFields(tClass);
            } else {
                this.fields = new ArrayList<>(columnFunctions.size());
                for (Pair<SFunction<T, ?>, ExcelColumnClass> pair : columnFunctions) {
                    try {
                        String fieldName = SFunction.getFieldName(pair.getKey());
                        try {
                            ExcelColumnClass columnClass = pair.getValue();
                            Field field = gettClass().getDeclaredField(fieldName);
                            if (columnClass == null) {
                                AllFieldColumn fieldColumn = tClass.getAnnotation(AllFieldColumn.class);
                                columnClass = ExcelColumnClass.getExcelColumn(fieldColumn, field);
                                if (columnClass == null) {
                                    columnClass = ExcelColumnClass.build();
                                }
                            }
                            this.fields.add(getCellField(gettClass(), field, columnClass));
                        } catch (NoSuchFieldException e) {
                            System.err.println("Get Field: " + fieldName + " Fail!\n" + e.getMessage());
                        }
                    } catch (ReflectiveOperationException e) {
                        System.err.println("Get FieldName Fail!\n" + e.getMessage());
                    }
                }
            }
        }
        return this.fields;
    }

    private List<CellField> getFields(Class<?> tClass) {
        Field[] fields = tClass.getDeclaredFields();
        List<CellField> fieldList = new ArrayList<>();
        AllFieldColumn fieldColumn = tClass.getAnnotation(AllFieldColumn.class);
        for (Field field : fields) {
            ExcelColumnClass column = ExcelColumnClass.getExcelColumn(fieldColumn, field);
            if (column != null) {
                fieldList.add(getCellField(tClass, field, column));
            }
        }
        sortFields(fieldList);
        return fieldList;
    }

    private CellField getCellField(Class<?> tClass, Field field, ExcelColumnClass column) {
        CellField cellField = new CellField();
        cellField.setIndex(column.getIndex());
        String name = field.getName();
        char[] chars = name.toCharArray();
        chars[0] = Character.toUpperCase(chars[0]);
        String nameFormat = String.valueOf(chars);
        String title = column.getTitle();
        String methodName = "get" + nameFormat;
        try {
            final Method method = tClass.getDeclaredMethod(methodName);
            cellField.setValueFun(obj -> {
                try {
                    return method.invoke(obj);
                } catch (IllegalAccessException | InvocationTargetException e) {
                    System.err.println("Get " + name + " Field Get Method Fail：" + methodName + "\n" + e.getMessage());
                    return "";
                }
            });
        } catch (NoSuchMethodException e) {
            System.err.println("The " + methodName + " method call failure\n" + e.getMessage());
        }
        if (Collection.class.isAssignableFrom(field.getType())) {
            ParameterizedType genericType = (ParameterizedType) field.getGenericType();
            Class<?> fieldType = (Class<?>) genericType.getActualTypeArguments()[0];
            cellField.setCellFields(getFields(fieldType));
            cellField.setCellType(CellType.LIST);
            if (title != null && !title.isEmpty()) {
                cellField.setTitle(title);
            }
        } else if (CellType.OBJECT.equals(column.getType())) {
            Class<?> type = field.getType();
            cellField.setCellFields(getFields(type));
            cellField.setCellType(CellType.OBJECT);
            if (title != null && !title.isEmpty()) {
                cellField.setTitle(title);
            }
        } else {
            cellField.setTitle(title, nameFormat.replaceAll("(?<![A-Z]|^)[A-Z]", " $0"));
            cellField.setCellType(column.getType());
            CellStyle cellStyle = getCellStyle();
            Font font = workbook.getFontAt(cellStyle.getFontIndex());
            column.settingStyle(cellStyle, workbook.createDataFormat(), font);

            cellField.setCellStyle(cellStyle);
            int width = column.getWidth();
            cellField.setWidth(width > 0 ? width : getColumnWidth());
        }
        return cellField;
    }

    static void sortFields(List<CellField> fieldList) {
        Set<Integer> indexSet = fieldList.stream().map(CellField::getIndex)
                .filter(index -> index > -1).collect(Collectors.toSet());
        int index = 0;
        for (CellField cellField : fieldList) {
            if (cellField.getIndex() < 0) {
                while (!indexSet.add(index)) {
                    index++;
                }
                cellField.setIndex(index);
            }
        }
        fieldList.sort(Comparator.comparingInt(CellField::getIndex));
    }

    void operateSheet() {
        if (operateSheet != null) {
            operateSheet.accept(getSheet(), getData());
        }
    }

    void operateRow(SXSSFRow row, Object rowData) {
        if (operateRow != null) {
            operateRow.accept(row, rowData);
        }
    }

    void operateCell(SXSSFCell cell, Object rowData) {
        if (operateCell != null) {
            operateCell.accept(cell, rowData);
        }
    }

    void operateTitle(SXSSFCell cell) {
        if (operateTitle != null) {
            operateTitle.accept(cell);
        }
    }

    CellStyle operateCellStyle(CellField cellField, Object rowData) {
        CellStyle cellStyle = cellField.getCellStyle();
        if (cellStyleFunMap == null) {
            return cellStyle;
        }
        int index = cellField.getIndex();
        BiConsumer<CellStyle, Object> cellStyleFun = cellStyleFunMap.get(index);
        if (cellStyleFun != null) {
            CellStyle newCellStyle = this.getCellStyle();
            cellStyleFun.accept(newCellStyle, rowData);
            cellStyle = newCellStyle;
        }
        return cellStyle;
    }

    public List<T> getData() {
        return data;
    }

    public Class<T> gettClass() {
        return tClass;
    }

    public short getTitleSize() {
        return titleSize;
    }

    public short getFontSize() {
        return fontSize;
    }

    public float getTitleHeight() {
        return titleHeight;
    }

    public float getRowHeight() {
        return rowHeight;
    }

    public int getColumnWidth() {
        return columnWidth;
    }

    public boolean isAutoColumnWidth() {
        return autoColumnWidth;
    }
}