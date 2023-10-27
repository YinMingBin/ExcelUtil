package ymb.github.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import ymb.github.excel.annotation.AllFieldColumn;

import java.lang.invoke.SerializedLambda;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.util.*;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Function;
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
    private BiConsumer<SXSSFCell, Object> operateValue;
    private BiConsumer<SXSSFSheet, List<T>> operateSheet;
    private List<CellField> fields;
    private final Class<T> tClass;
    private short titleSize = 12;
    private short valueSize = 11;
    private float titleHeight = 25;
    private float valueHeight = 20;
    private int columnWidth = 10;
    private final Map<Integer, BiConsumer<CellStyle, Object>> valueStyleFunMap = new HashMap<>();
    private Consumer<CellStyle> valueStyleFun = cell -> {};
    private Consumer<CellStyle> titleStyleFun = cell -> {};
    private boolean autoColumnWidth = false;

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
     * 设置数据的字体大小
     * @param valueSize 字体大小
     * @return this
     */
    @Override
    public SheetOperate<T> setValueSize(short valueSize) {
        this.valueSize = valueSize;
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
     * 设置数据的行高
     * @param valueHeight 行高
     * @return this
     */
    @Override
    public SheetOperate<T> setValueHeight(short valueHeight) {
        this.valueHeight = valueHeight;
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
     * 设置数据样式
     * @param valueStyleFun (CellStyle) -> void
     * @return this
     */
    @Override
    public SheetOperate<T> setValueStyle(Consumer<CellStyle> valueStyleFun) {
        this.valueStyleFun = valueStyleFun;
        return this;
    }

    /**
     * 操作某一列数据的样式 (设置Cell时调用)
     * @param index 列索引
     * @param valueStyle (CellStyle, value) -> void
     * @return this
     */
    @Override
    public SheetOperate<T> operateValueStyle(int index, BiConsumer<CellStyle, Object> valueStyle) {
        valueStyleFunMap.put(index, valueStyle);
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
     * 操作数据，每次设置数据之后执行
     * @param operateValue (Cell, data) -> void
     * @return this
     */
    @Override
    public SheetOperate<T> operateValue(BiConsumer<SXSSFCell, Object> operateValue) {
        this.operateValue = operateValue;
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
    public final SheetOperate<T> settingColumn(SFunction<T, Object>... functions) {
        for (SFunction<T, ?> function : functions) {
            if (function != null) {
                settingColumn(function, null);
            }
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
        if (function == null) {
            return this;
        }

        try {
            String fieldName = getFieldName(function);
            try {
                Field field = gettClass().getDeclaredField(fieldName);
                if (this.fields == null) {
                    this.fields = new ArrayList<>();
                }
                if (columnClass == null) {
                    AllFieldColumn fieldColumn = tClass.getAnnotation(AllFieldColumn.class);
                    columnClass = ExcelColumnClass.getExcelColumn(fieldColumn, field);
                    if (columnClass == null) {
                        columnClass = ExcelColumnClass.build();
                    }
                }
                this.fields.add(getCellField(gettClass(), field, columnClass));
                sortFields(this.fields);
            } catch (NoSuchFieldException e) {
                System.err.println("Get Field: " + fieldName + " Fail!\n" + e.getMessage());
            }
        } catch (ReflectiveOperationException e) {
            System.err.println("Get FieldName Fail!\n" + e.getMessage());
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

    CellStyle getValueStyle() {
        CellStyle cellStyle = workbook.createCellStyle();
        // 设置字体
        Font font = workbook.createFont();
        font.setFontHeightInPoints(getValueSize());
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

        valueStyleFun.accept(cellStyle);
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
            this.fields = getFields(tClass);
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
        cellField.setTitle(column.getTitle(), nameFormat.replaceAll("(?<![A-Z]|^)[A-Z]", " $0"));
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
        } else {
            cellField.setCellType(column.getType());
            CellStyle cellStyle = getValueStyle();
            Font font = workbook.getFontAt(cellStyle.getFontIndex());
            column.settingStyle(cellStyle, workbook.createDataFormat(), font);

            cellField.setCellStyle(cellStyle);
            int width = column.getWidth();
            cellField.setWidth(width > 0 ? width : getColumnWidth());
        }
        return cellField;
    }

    private void sortFields(List<CellField> fieldList) {
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

    void operateValue(SXSSFCell cell, Object data) {
        if (operateValue != null) {
            operateValue.accept(cell, data);
        }
    }

    void operateTitle(SXSSFCell cell) {
        if (operateTitle != null) {
            operateTitle.accept(cell);
        }
    }

    CellStyle operateValueStyle(CellField cellField, Object value) {
        int index = cellField.getIndex();
        CellStyle cellStyle = cellField.getCellStyle();
        BiConsumer<CellStyle, Object> cellStyleFun = valueStyleFunMap.get(index);
        if (cellStyleFun != null) {
            CellStyle newCellStyle = workbook.createCellStyle();
            newCellStyle.cloneStyleFrom(cellStyle);
            cellStyleFun.accept(newCellStyle, value);
            cellStyle = newCellStyle;
        }
        return cellStyle;
    }

    public static <T> String getFieldName(Function<T, ?> fn) throws ReflectiveOperationException {
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

    public String getSheetName() {
        return sheetName;
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

    public short getValueSize() {
        return valueSize;
    }

    public float getTitleHeight() {
        return titleHeight;
    }

    public float getValueHeight() {
        return valueHeight;
    }

    public int getColumnWidth() {
        return columnWidth;
    }

    public boolean isAutoColumnWidth() {
        return autoColumnWidth;
    }
}