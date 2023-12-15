package ymb.github.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ymb.github.excel.annotation.AllFieldColumn;
import ymb.github.excel.annotation.ExcelClass;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.function.Supplier;

/**
 * Excel导入工具类
 *
 * @author WuLiao
 */
@SuppressWarnings("unused")
public class ExcelImportUtil {
    private final Workbook workbook;
    private Sheet sheet;
    private int startRow;

    /**
     * 唯一的构造方法
     *
     * @param is Excel文件输入流
     * @throws IOException io异常
     */
    public ExcelImportUtil(InputStream is) throws IOException {
        this.workbook = new XSSFWorkbook(is);
    }

    /**
     * 获取Workbook对象
     *
     * @return Workbook对象
     */
    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * 读取Excel文件中的数据
     *
     * @param is        Excel文件输入流
     * @param tClass    数据类型对象
     * @param sheetName Excel中Sheet的名称
     * @param <T>       数据类型
     * @return 数据集
     * @throws IOException io异常
     */
    public static <T> List<T> read(InputStream is, Class<T> tClass, String sheetName) throws IOException {
        ExcelImportUtil util = new ExcelImportUtil(is);
        List<T> dataList = util.read(tClass, sheetName);
        util.close();
        return dataList;
    }

    /**
     * 读取Excel文件中的数据
     *
     * @param is         Excel文件输入流
     * @param tClass     数据类型对象
     * @param sheetIndex Excel中Sheet的下标
     * @param <T>        数据类型
     * @return 数据集
     * @throws IOException io异常
     */
    public static <T> List<T> read(InputStream is, Class<T> tClass, int sheetIndex) throws IOException {
        ExcelImportUtil util = new ExcelImportUtil(is);
        List<T> dataList = util.read(tClass, sheetIndex);
        util.close();
        return dataList;
    }

    /**
     * 读取Excel文件中的数据（读取第一个Sheet的数据）
     *
     * @param is     Excel文件输入流
     * @param tClass 数据类型对象
     * @param <T>    数据类型
     * @return 数据集
     * @throws IOException io异常
     */
    public static <T> List<T> read(InputStream is, Class<T> tClass) throws IOException {
        ExcelImportUtil util = new ExcelImportUtil(is);
        List<T> dataList = util.read(tClass, 0);
        util.close();
        return dataList;
    }

    /**
     * 读取Excel文件中的数据（读取第一个Sheet的数据）
     *
     * @param tClass 数据类型对象
     * @param <T>    数据类型
     * @return 数据集
     */
    public <T> List<T> read(Class<T> tClass) {
        return this.read(tClass, 0);
    }

    /**
     * 读取Excel文件中的数据
     *
     * @param tClass    数据类型对象
     * @param sheetName Excel中Sheet的名称
     * @param <T>       数据类型
     * @return 数据集
     */
    public <T> List<T> read(Class<T> tClass, String sheetName) {
        return this.read(tClass, workbook.getSheetIndex(sheetName));
    }

    /**
     * 读取Excel文件中的数据
     *
     * @param tClass     数据类型对象
     * @param sheetIndex Excel中Sheet的下标
     * @param <T>        数据类型
     * @return 数据集
     */
    public <T> List<T> read(Class<T> tClass, int sheetIndex) {
        return read(tClass, () -> {
            ExcelClass annotation = tClass.getAnnotation(ExcelClass.class);
            return getFields(tClass, annotation == null ? 0 : 1);
        }, sheetIndex);
    }

    /**
     * 根据对象中的某些属性，读取Excel文件中的数据
     *
     * @param tClass    数据类型对象
     * @param getFunArr 对象中属性的get方法
     * @param <T>       数据类型
     * @return 数据集
     */
    @SafeVarargs
    public final <T> List<T> read(Class<T> tClass, SFunction<T, ?>... getFunArr) {
        return read(tClass, 0, getFunArr);
    }

    /**
     * 根据对象中的某些属性，读取Excel文件中的数据
     *
     * @param tClass    数据类型对象
     * @param sheetName Excel中Sheet的名称
     * @param getFunArr 对象中属性的get方法
     * @param <T>       数据类型
     * @return 数据集
     */
    @SafeVarargs
    public final <T> List<T> read(Class<T> tClass, String sheetName, SFunction<T, ?>... getFunArr) {
        return read(tClass, workbook.getSheetIndex(sheetName), getFunArr);
    }

    /**
     * 根据对象中的某些属性，读取Excel文件中的数据
     *
     * @param tClass 数据类型对象
     * @param sheetIndex Excel中Sheet的下标
     * @param getFunArr 对象中属性的get方法
     * @param <T> 数据类型
     * @return 数据集
     */
    @SafeVarargs
    public final <T> List<T> read(Class<T> tClass, int sheetIndex, SFunction<T, ?>... getFunArr) {
        return read(tClass, () -> {
            Field[] fields = new Field[getFunArr.length];
            for (int i = 0; i < getFunArr.length; i++) {
                SFunction<T, ?> getFun = getFunArr[i];
                try {
                    String fieldName = SFunction.getFieldName(getFun);
                    Field field = tClass.getDeclaredField(fieldName);
                    fields[i] = field;
                } catch (ReflectiveOperationException e) {
                    throw new RuntimeException(e);
                }
            }
            ExcelClass annotation = tClass.getAnnotation(ExcelClass.class);
            return getFields(tClass, annotation == null ? 0 : 1, fields);
        }, sheetIndex);
    }

    private <T> List<T> read(Class<T> tClass, Supplier<List<CellField>> getCellFields, int sheetIndex) {
        // 重置数据
        this.sheet = workbook.getSheetAt(sheetIndex);
        this.startRow = 0;
        // 读取数据
        return getDataList(tClass, getCellFields.get(), startRow);
    }

    private List<CellField> getFields(Class<?> tClass, int rowIndex) {
        return getFields(tClass, rowIndex, tClass.getDeclaredFields());
    }

    private List<CellField> getFields(Class<?> tClass, int rowIndex, Field[] fields) {
        rowIndex++;
        List<CellField> fieldList = new ArrayList<>();
        AllFieldColumn fieldColumn = tClass.getAnnotation(AllFieldColumn.class);
        for (Field field : fields) {
            if (Collection.class.isAssignableFrom(field.getType())) {
                continue;
            }
            ExcelColumnClass column = ExcelColumnClass.getExcelColumn(fieldColumn, field);
            if (column != null) {
                fieldList.add(getCellField(tClass, field, column, rowIndex));
            }
        }
        SheetOperate.sortFields(fieldList);
        return fieldList;
    }

    private CellField getCellField(Class<?> tClass, Field field, ExcelColumnClass column, int rowIndex) {
        this.startRow = Math.max(rowIndex, this.startRow);
        CellField cellField = new CellField();
        Class<?> type = field.getType();
        cellField.setFieldType(type);
        cellField.setIndex(column.getIndex());
        String name = field.getName();
        char[] chars = name.toCharArray();
        chars[0] = Character.toUpperCase(chars[0]);
        String methodName = "set" + String.valueOf(chars);
        try {
            final Method method = tClass.getDeclaredMethod(methodName, type);
            cellField.setSettingFun((obj, val) -> {
                try {
                    method.invoke(obj, val);
                } catch (IllegalAccessException | InvocationTargetException e) {
                    System.err.println("Call " + name + " Field Set Method Fail：" + methodName + "\n" + e.getMessage());
                }
            });
        } catch (NoSuchMethodException e) {
            System.err.println("The " + methodName + " method call failure\n" + e.getMessage());
        }
        cellField.setCellType(column.getType());
        return cellField;
    }


    private <T> List<T> getDataList(Class<T> tClass, List<CellField> fields, Integer rowIndex) {
        List<T> list = new ArrayList<>();
        Row row = this.sheet.getRow(rowIndex);
        while (row != null) {
            T rowData = getRowData(tClass, fields, row);
            if (rowData != null) {
                list.add(rowData);
            }
            row = this.sheet.getRow(++rowIndex);
        }
        return list;
    }


    private <T> T getRowData(Class<T> tClass, List<CellField> fields, Row row) {
        try {
            T t = tClass.newInstance();
            for (CellField field : fields) {
                Class<?> fieldType = field.getFieldType();
                int index = field.getIndex();
                Cell cell = row.getCell(index);
                if (cell == null) {
                    continue;
                }
                CellType cellType = field.getCellType();
                Object value = getValue(cell, cellType, fieldType);
                field.getSettingFun().accept(t, value);
            }
            return t;
        } catch (InstantiationException | IllegalAccessException e) {
            System.err.println(tClass + " create fail:\n" + e.getMessage());
        }
        return null;
    }

    private Object getValue(Cell cell, CellType cellType, Class<?> fieldType) {
        Object value = null;
        if (LocalDate.class.isAssignableFrom(fieldType)) {
            LocalDateTime cellValue = cell.getLocalDateTimeCellValue();
            value = cellValue.toLocalDate();
        } else if (LocalDateTime.class.isAssignableFrom(fieldType)) {
            value = cell.getLocalDateTimeCellValue();
        } else if (Date.class.isAssignableFrom(fieldType)) {
            value = cell.getDateCellValue();
        } else if (RichTextString.class.isAssignableFrom(fieldType)) {
            value = cell.getRichStringCellValue();
        } else if (Character.class.isAssignableFrom(fieldType) || char.class.isAssignableFrom(fieldType)) {
            value = cell.getStringCellValue().charAt(0);
        } else if (CellType.NUMERIC.equals(cellType)) {
            double cellValue = cell.getNumericCellValue();
            if (Byte.class.isAssignableFrom(fieldType) || byte.class.isAssignableFrom(fieldType)) {
                value = (byte) cellValue;
            } else if (Short.class.isAssignableFrom(fieldType) || short.class.isAssignableFrom(fieldType)) {
                value = (short) cellValue;
            } else if (Integer.class.isAssignableFrom(fieldType) || int.class.isAssignableFrom(fieldType)) {
                value = (int) cellValue;
            } else if (Long.class.isAssignableFrom(fieldType) || long.class.isAssignableFrom(fieldType)) {
                value = (long) cellValue;
            } else if (Float.class.isAssignableFrom(fieldType) || float.class.isAssignableFrom(fieldType)) {
                value = (float) cellValue;
            } else if (Double.class.isAssignableFrom(fieldType) || double.class.isAssignableFrom(fieldType)) {
                value = cellValue;
            }
        } else if (CellType.BOOLEAN.equals(cellType)) {
            value = cell.getBooleanCellValue();
        } else if (!CellType.BLANK.equals(cellType)) {
            value = cell.getStringCellValue();
        }
        return value;
    }

    /**
     * 关闭Workbook流
     */
    public void close() {
        try {
            workbook.close();
        } catch (IOException e) {
            System.err.println("ExcelImportUtil -> workbook close error:" + e.getMessage());
        }
    }

    /**
     * 关闭Workbook流并将传入的输入流关闭
     * @throws IOException io异常
     */
    public void close(InputStream is) throws IOException {
        this.close();
        is.close();
    }
}
