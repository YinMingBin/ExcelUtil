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
        List<CellField> fields = getFields(tClass, annotation == null ? 0 : 1);
        // 读取数据
        return getDataList(tClass, fields, startRow);
    }

    private List<CellField> getFields(Class<?> tClass, int rowIndex) {
        rowIndex++;
        Field[] fields = tClass.getDeclaredFields();
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
}
