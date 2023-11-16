package ymb.github.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ymb.github.excel.annotation.AllFieldColumn;
import ymb.github.excel.annotation.ExcelClass;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

/**
 * Excel导入工具类
 * @author WuLiao
 */
public class ExcelImportUtil {
    private final Workbook workbook;
    private Sheet sheet;
    private int startRow;

    public ExcelImportUtil(InputStream is) throws IOException {
        this.workbook = new XSSFWorkbook(is);
    }

    public <T> List<T> read(Class<T> tClass) {
        return this.read(tClass, 0);
    }

    public <T> List<T> read(Class<T> tClass, String sheetName) {
        return this.read(tClass, workbook.getSheetIndex(sheetName));
    }

    public <T> List<T> read(Class<T> tClass, int sheetIndex) {
        this.sheet = workbook.getSheetAt(sheetIndex);
        this.startRow = 0;
        // 获取CellField 及 数据开始行
        ExcelClass annotation = tClass.getAnnotation(ExcelClass.class);
        List<CellField> fields = getFields(tClass, annotation == null ? -1 : 0);
        // 循环读取数据
        getList(tClass, fields, startRow);
        return null;
    }

    private List<CellField> getFields(Class<?> tClass, int rowIndex) {
        rowIndex++;
        Field[] fields = tClass.getDeclaredFields();
        List<CellField> fieldList = new ArrayList<>();
        AllFieldColumn fieldColumn = tClass.getAnnotation(AllFieldColumn.class);
        for (Field field : fields) {
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
        if (Collection.class.isAssignableFrom(field.getType())) {
            ParameterizedType genericType = (ParameterizedType) field.getGenericType();
            Class<?> fieldType = (Class<?>) genericType.getActualTypeArguments()[0];
            cellField.setCellFields(getFields(fieldType, rowIndex));
        } else {
            cellField.setCellType(column.getType());
        }
        return cellField;
    }

    private <T> T getList(Class<T> tClass, List<CellField> fields, Integer rowIndex) {
        Row row = this.sheet.getRow(rowIndex);
        if (row != null) {
            try {
                T t = tClass.newInstance();
                for (CellField field : fields) {
                    Class<?> fieldType = field.getFieldType();
                    if (Collection.class.isAssignableFrom(fieldType)) {

                    }
                    List<CellField> cellFields = field.getCellFields();
                    if (cellFields != null) {
                        Object value = getList(fieldType, cellFields, rowIndex);
                        field.getSettingFun().accept(t, value);
                        continue;
                    }
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
        } else if (CellType.NUMERIC.equals(cellType)) {
            double cellValue = cell.getNumericCellValue();
            value = fieldType.cast(cellValue);
        } else if (CellType.BOOLEAN.equals(cellType)) {
            value = cell.getBooleanCellValue();
        } else if (!CellType.BLANK.equals(cellType)) {
            value = cell.getStringCellValue();
        }
        return value;
    }
}
