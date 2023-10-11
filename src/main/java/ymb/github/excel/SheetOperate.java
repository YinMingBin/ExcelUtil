package ymb.github.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.*;
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
public final class SheetOperate<T> {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private String sheetName;
    private List<T> data;
    private Consumer<XSSFCell> operateTitle;
    private BiConsumer<XSSFCell, Object> operateValue;
    private BiConsumer<XSSFSheet, List<T>> operateSheet;
    private List<CellField> fields;
    private final Class<T> tClass;
    private short titleSize = 12;
    private short valueSize = 11;
    private float titleHeight = 25;
    private float valueHeight = 20;
    private final Map<Integer, BiConsumer<XSSFCellStyle, Object>> valueStyleFunMap = new HashMap<>();
    private Consumer<XSSFCellStyle> valueStyleFun = cell -> {};
    private Consumer<XSSFCellStyle> titleStyleFun = cell -> {};

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

    void setWorkbook(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    /**
     * 设置数据源
     * @param data 数据源
     * @return this
     */
    public SheetOperate<T> setData(List<T> data) {
        this.data = data;
        return this;
    }

    /**
     * 设置表头的字体大小
     * @param titleSize 字体大小
     * @return this
     */
    public SheetOperate<T> setTitleSize(short titleSize) {
        this.titleSize = titleSize;
        return this;
    }

    /**
     * 设置数据的字体大小
     * @param valueSize 字体大小
     * @return this
     */
    public SheetOperate<T> setValueSize(short valueSize) {
        this.valueSize = valueSize;
        return this;
    }

    /**
     * 设置表头的行高
     * @param titleHeight 行高
     * @return this
     */
    public SheetOperate<T> setTitleHeight(short titleHeight) {
        this.titleHeight = titleHeight;
        return this;
    }

    /**
     * 设置数据的行高
     * @param valueHeight 行高
     * @return this
     */
    public SheetOperate<T> setValueHeight(short valueHeight) {
        this.valueHeight = valueHeight;
        return this;
    }

    /**
     * 设置表头样式
     * @param titleStyleFun (CellStyle) -> void
     * @return this
     */
    public SheetOperate<T> setTitleStyle(Consumer<XSSFCellStyle> titleStyleFun) {
        this.titleStyleFun = titleStyleFun;
        return this;
    }

    /**
     * 设置数据样式
     * @param valueStyleFun (CellStyle) -> void
     * @return this
     */
    public SheetOperate<T> setValueStyle(Consumer<XSSFCellStyle> valueStyleFun) {
        this.valueStyleFun = valueStyleFun;
        return this;
    }

    /**
     * 操作某一列数据的样式 (设置Cell时调用)
     * @param index 列索引
     * @param valueStyle (CellStyle, value) -> void
     * @return this
     */
    public SheetOperate<T> operateValueStyle(int index, BiConsumer<XSSFCellStyle, Object> valueStyle) {
        valueStyleFunMap.put(index, valueStyle);
        return this;
    }

    /**
     * 操作表头，每次设置表头之后执行
     * @param operateTitle (Cell) -> void
     * @return this
     */
    public SheetOperate<T> operateTitle(Consumer<XSSFCell> operateTitle) {
        this.operateTitle = operateTitle;
        return this;
    }

    /**
     * 操作数据，每次设置数据之后执行
     * @param operateValue (Cell, data) -> void
     * @return this
     */
    public SheetOperate<T> operateValue(BiConsumer<XSSFCell, Object> operateValue) {
        this.operateValue = operateValue;
        return this;
    }

    /**
     * 操作Sheet，在数据生成完之后执行
     * @param operateSheet (Sheet, dataList) -> void
     * @return this
     */
    public SheetOperate<T> operateSheet(BiConsumer<XSSFSheet, List<T>> operateSheet) {
        this.operateSheet = operateSheet;
        return this;
    }

    XSSFSheet getSheet() {
        if (sheet == null) {
            sheet = sheetName == null ? workbook.createSheet() : workbook.createSheet(sheetName);
            sheetName = sheet.getSheetName();
        }
        return sheet;
    }

    void clearSheet() {
        this.sheet = null;
    }

    XSSFCellStyle getValueStyle() {
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        // 设置字体
        XSSFFont font = workbook.createFont();
        font.setFontHeightInPoints(valueSize);
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

    XSSFCellStyle getTitleStyle() {
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        // 设置字体
        XSSFFont font = workbook.createFont();
        font.setFontHeightInPoints(titleSize);
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
        XSSFDataFormat dataFormat = workbook.createDataFormat();
        AllFieldColumn fieldColumn = tClass.getAnnotation(AllFieldColumn.class);
        for (Field field : fields) {
            ExcelColumnClass column = ExcelColumnClass.getExcelColumn(fieldColumn, field);
            if (column != null) {
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
                            System.err.println("Get Field: " + name + " Fail：" + methodName);
                            return "";
                        }
                    });
                } catch (NoSuchMethodException e) {
                    System.err.println("The " + methodName + " method call failure");
                }
                if (Collection.class.isAssignableFrom(field.getType())) {
                    ParameterizedType genericType = (ParameterizedType) field.getGenericType();
                    Class<?> fieldType = (Class<?>) genericType.getActualTypeArguments()[0];
                    cellField.setCellFields(getFields(fieldType));
                } else {
                    cellField.setCellType(column.getType());
                    XSSFCellStyle cellStyle = getValueStyle();

                    column.settingStyle(cellStyle, dataFormat);

                    cellField.setCellStyle(cellStyle);
                    cellField.setWidth(column.getWidth());
                }
                fieldList.add(cellField);
            }
        }
        sortFields(fieldList);
        return fieldList;
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
            operateSheet.accept(getSheet(), data);
        }
    }

    void operateValue(XSSFCell cell, Object data) {
        if (operateValue != null) {
            operateValue.accept(cell, data);
        }
    }

    void operateTitle(XSSFCell cell) {
        if (operateTitle != null) {
            operateTitle.accept(cell);
        }
    }

    XSSFCellStyle operateValueStyle(CellField cellField, Object value) {
        int index = cellField.getIndex();
        XSSFCellStyle cellStyle = cellField.getCellStyle();
        BiConsumer<XSSFCellStyle, Object> cellStyleFun = valueStyleFunMap.get(index);
        if (cellStyleFun != null) {
            XSSFCellStyle newCellStyle = workbook.createCellStyle();
            newCellStyle.cloneStyleFrom(cellStyle);
            cellStyleFun.accept(newCellStyle, value);
            cellStyle = newCellStyle;
        }
        return cellStyle;
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
}