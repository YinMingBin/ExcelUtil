package ymb.github.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import ymb.github.excel.annotation.ExcelClass;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.function.BiConsumer;
import java.util.function.Consumer;

/**
 * @author YinMingBin
 */
@SuppressWarnings({"unused", "UnusedReturnValue"})
public class ExcelUtil<T> implements Operate<T, ExcelUtil<T>> {
    private final SheetOperate<T> sheetOperate;
    private final Stack<SheetOperate<?>> otherSheet = new Stack<>();
    private final SXSSFWorkbook workbook;
    private SheetOperate<?> currentSheet;
    private String csv;
    private final Map<Integer, Integer> maxWidthMap = new HashMap<>();
    // int[0] = firstRow, int[1] = firstCol (firstCol == endCol)
    private Map<Integer, int[]> columnRangeMap = new HashMap<>();
    private Map<String, int[]> columnRangeByKeyMap = new HashMap<>();

    public ExcelUtil(Class<T> tClass) {
        this.workbook = new SXSSFWorkbook();
        this.sheetOperate = SheetOperate.create(tClass, workbook);
        otherSheet.add(this.sheetOperate);
    }

    public ExcelUtil(Class<T> tClass, String sheetName) {
        this.workbook = new SXSSFWorkbook();
        this.sheetOperate = SheetOperate.create(tClass, sheetName, workbook);
        otherSheet.add(this.sheetOperate);
    }

    public ExcelUtil(Class<T> tClass, List<T> data) {
        this.workbook = new SXSSFWorkbook();
        this.sheetOperate = SheetOperate.create(tClass, workbook);
        otherSheet.add(this.sheetOperate);

        this.setData(data);
    }

    public ExcelUtil(Class<T> tClass, String sheetName, List<T> data) {
        this.workbook = new SXSSFWorkbook();
        this.sheetOperate = SheetOperate.create(tClass, sheetName, workbook);
        otherSheet.add(this.sheetOperate);

        this.setData(data);
    }

    public SXSSFWorkbook getWorkbook() {
        return workbook;
    }

    public String getCsv() {
        return csv;
    }

    /**
     * 设置数据源
     * @param data 数据源
     * @return this
     */
    @Override
    public ExcelUtil<T> setData(List<T> data) {
        sheetOperate.setData(data);
        return this;
    }

    /**
     * 设置表头的字体大小
     * @param titleSize 字体大小
     * @return this
     */
    @Override
    public ExcelUtil<T> setTitleSize(short titleSize) {
        sheetOperate.setTitleSize(titleSize);
        return this;
    }

    /**
     * 设置Cell的字体大小
     * @param fontSize 字体大小
     * @return this
     */
    @Override
    public ExcelUtil<T> setFontSize(short fontSize) {
        sheetOperate.setFontSize(fontSize);
        return this;
    }

    /**
     * 设置表头的行高
     * @param titleHeight 行高
     * @return this
     */
    @Override
    public ExcelUtil<T> setTitleHeight(short titleHeight) {
        sheetOperate.setTitleHeight(titleHeight);
        return this;
    }

    /**
     * 设置数据的行高
     * @param rowHeight 行高
     * @return this
     */
    @Override
    public ExcelUtil<T> setRowHeight(short rowHeight) {
        sheetOperate.setRowHeight(rowHeight);
        return this;
    }

    /**
     * 设置列宽
     * @param columnWidth 列宽
     * @return this
     */
    @Override
    public ExcelUtil<T> setColumnWidth(int columnWidth) {
        sheetOperate.setColumnWidth(columnWidth);
        return this;
    }

    /**
     * 设置表头样式
     * @param titleStyle (CellStyle) -> void
     * @return this
     */
    @Override
    public ExcelUtil<T> setTitleStyle(Consumer<CellStyle> titleStyle) {
        sheetOperate.setTitleStyle(titleStyle);
        return this;
    }

    /**
     * 设置单元格（Cell）的样式
     * @param cellStyleFun (CellStyle) -> void
     * @return this
     */
    @Override
    public ExcelUtil<T> setCellStyle(Consumer<CellStyle> cellStyleFun) {
        sheetOperate.setCellStyle(cellStyleFun);
        return this;
    }

    /**
     * 操作某一列单元格（Cell）的样式 (设置完数据时调用)
     * @param columnIndex 列索引
     * @param cellStyle (CellStyle, rowData) -> void
     * @return this
     */
    @Override
    public ExcelUtil<T> operateCellStyle(int columnIndex, BiConsumer<CellStyle, Object> cellStyle) {
        sheetOperate.operateCellStyle(columnIndex, cellStyle);
        return this;
    }

    /**
     * 操作某一列单元格（Cell）的样式 (设置完数据时调用)
     * @param columnKey 列key
     * @param cellStyle (CellStyle, rowData) -> void
     * @return this
     */
    @Override
    public ExcelUtil<T> operateCellStyle(String columnKey, BiConsumer<CellStyle, Object> cellStyle) {
        sheetOperate.operateCellStyle(columnKey, cellStyle);
        return this;
    }

    /**
     * 操作表头，每次设置表头之后执行
     * @param operateTitle (Cell) -> void
     * @return this
     */
    @Override
    public ExcelUtil<T> operateTitle(Consumer<SXSSFCell> operateTitle) {
        sheetOperate.operateTitle(operateTitle);
        return this;
    }

    /**
     * 操作单元格（Cell），每次设置数据之后执行
     * @param operateCell (Cell, rowData) -> void
     * @return this
     */
    @Override
    public ExcelUtil<T> operateCell(BiConsumer<SXSSFCell, Object> operateCell) {
        sheetOperate.operateCell(operateCell);
        return this;
    }

    /**
     * 操作某一列的单元格（Cell），每次设置数据之后执行
     * @param index 列下标
     * @param operateCell (Cell, RowData) -> void
     * @return this
     */
    @Override
    public ExcelUtil<T> operateCell(int index, BiConsumer<SXSSFCell, Object> operateCell) {
        sheetOperate.operateCell(index, operateCell);
        return this;
    }

    /**
     * 操作某一列的单元格（Cell），每次设置数据之后执行
     * @param key 列key
     * @param operateCell (Cell, RowData) -> void
     * @return this
     */
    @Override
    public ExcelUtil<T> operateCell(String key, BiConsumer<SXSSFCell, Object> operateCell) {
        sheetOperate.operateCell(key, operateCell);
        return this;
    }

    /**
     * 操作单元格（Cell），每次设置完一行数据之后执行
     * @param operateRow (Row, rowData) -> void
     * @return this
     */
    @Override
    public ExcelUtil<T> operateRow(BiConsumer<SXSSFRow, Object> operateRow) {
        sheetOperate.operateRow(operateRow);
        return this;
    }

    /**
     * 操作Sheet，在数据生成完之后执行
     * @param operateSheet (Sheet, dataList) -> void
     * @return this
     */
    @Override
    public ExcelUtil<T> operateSheet(BiConsumer<SXSSFSheet, List<T>> operateSheet) {
        sheetOperate.operateSheet(operateSheet);
        return this;
    }

    /**
     * 设置列
     * @param functions 字段的get方法（不定项参数）
     * @return this
     */
    @SafeVarargs
    @Override
    public final ExcelUtil<T> settingColumn(SFunction<T, ?>... functions) {
        sheetOperate.settingColumn(functions);
        return this;
    }

    /**
     * 设置列
     * @param function 字段的get方法
     * @param columnClass 列属性
     * @return this
     */
    @Override
    public ExcelUtil<T> settingColumn(SFunction<T, ?> function, ExcelColumnClass columnClass) {
        sheetOperate.settingColumn(function, columnClass);
        return this;
    }

    /**
     * 启用自适应列宽（效率较低）
     * @return this
     */
    @Override
    public ExcelUtil<T> autoColumnWidth() {
        sheetOperate.autoColumnWidth();
        return this;
    }

    /**
     * 设置数据校验（下拉序列）
     * @param index 列下标
     * @param list 校验列表（下拉列表）
     * @return this
     */
    @Override
    public ExcelUtil<T> setDataValidationList(int index, Collection<String> list) {
        sheetOperate.setDataValidationList(index, list);
        return this;
    }

    /**
     * 设置数据校验（下拉序列）
     * @param key 列下标
     * @param list 校验列表（下拉列表）
     * @return this
     */
    @Override
    public ExcelUtil<T> setDataValidationList(String key, Collection<String> list) {
        sheetOperate.setDataValidationList(key, list);
        return this;
    }

    /**
     * 设置数据校验（下拉序列）
     * @param firstRow 开始行
     * @param firstCol 开始列
     * @param endRow 结束行
     * @param endCol 结束列
     * @param list 校验列表（下拉列表）
     * @return this
     */
    @Override
    public ExcelUtil<T> setDataValidationList(int firstRow, int firstCol, int endRow, int endCol, Collection<String> list) {
        sheetOperate.setDataValidationList(firstRow, firstCol, endRow, endCol, list);
        return this;
    }

    /**
     * 添加Sheet
     * @param sheetOperate SheetOperate.create
     * @return this
     */
    public ExcelUtil<T> addSheet(SheetOperate<?> sheetOperate) {
        sheetOperate.setWorkbook(this.workbook);
        otherSheet.add(sheetOperate);
        return this;
    }

    /**
     * 执行生成
     * @return this
     */
    public ExcelUtil<T> execute() {
        for (int i = workbook.getNumberOfSheets() - 1; i >= 0; i--) {
            workbook.removeSheetAt(i);
        }
        for (SheetOperate<?> operate : otherSheet) {
            this.currentSheet = operate;
            this.columnRangeMap = new HashMap<>();
            this.columnRangeByKeyMap = new HashMap<>();
            operate.clearSheet();
            List<CellField> fields = operate.getFields();
            if (fields.isEmpty()) {
                continue;
            }
            if (operate.isAutoColumnWidth()) {
                operate.getSheet().trackAllColumnsForAutoSizing();
                maxWidthMap.clear();
            }
            final int dataFirstRow = setExcelTitle(fields) + 1;
            final int dataEndRow = setExcelData(operate.getData(), fields, dataFirstRow);
            // 设置数据验证
            Map<Integer, Collection<String>> dataValidationMap = operate.getDataValidationMap();
            setDataValidation(operate, dataEndRow, dataValidationMap, columnRangeMap);
            Map<String, Collection<String>> dataValidationByKeyMap = operate.getDataValidationByKeyMap();
            setDataValidation(operate, dataEndRow, dataValidationByKeyMap, columnRangeByKeyMap);
            operate.operateSheet();
        }
        return this;
    }

    private <K> void setDataValidation(SheetOperate<?> operate,
                                       int endRow,
                                       Map<K, Collection<String>> dataValidationMap,
                                       Map<K, int[]> columnRangeMap) {
        if (dataValidationMap != null) {
            dataValidationMap.forEach((key, value) -> {
                int[] ranges = columnRangeMap.get(key);
                if (ranges != null) {
                    int firstRow = ranges[0];
                    int firstCol = ranges[1];
                    operate.setDataValidationList(firstRow, firstCol, endRow, firstCol, value);
                }
            });
        }
    }

    /**
     * 执行并转成byte数组
     * @return byte数组
     * @throws IOException write异常
     */
    public byte[] toByteArray() throws IOException {
        try (ByteArrayOutputStream os = new ByteArrayOutputStream()) {
            workbook.write(os);
            os.flush();
            return os.toByteArray();
        }
    }

    /**
     * 执行并写入OutputStream
     * @param os OutputStream流
     * @return this
     * @throws IOException write异常
     */
    public ExcelUtil<T> write(OutputStream os) throws IOException {
        workbook.write(os);
        os.flush();
        return this;
    }

    /**
     * 执行并写入文件
     * @param filePath 文件路径
     * @return this
     * @throws IOException write异常
     */
    public ExcelUtil<T> writeFile(String filePath) throws IOException {
        try (OutputStream os = new BufferedOutputStream(Files.newOutputStream(Paths.get(filePath)))) {
            write(os);
        }
        return this;
    }

    /**
     * 生成csv格式
     * @return this
     */
    public ExcelUtil<T> toCsv() {
        return toCsv(',', "\"", "\"\"", "\r\n");
    }

    /**
     * 生成csv格式
     * @param separator 分隔符
     * @param label 字段标识
     * @param escape 转义符
     * @param line 换行符
     * @return this
     */
    public ExcelUtil<T> toCsv(char separator, CharSequence label, CharSequence escape, String line) {
        StringBuilder sb = new StringBuilder();
        List<CellField> fields = sheetOperate.getFields();
        // title
        for (CellField field : fields) {
            String title = field.getTitle().replace(label, escape);
            sb.append(separator).append(label).append(title).append(label);
        }
        sb.deleteCharAt(0);
        // value
        List<T> data = sheetOperate.getData();
        if (data != null) {
            for (T datum : data) {
                StringBuilder vsb = new StringBuilder();
                for (CellField field : fields) {
                    String value = String.valueOf(field.getValueFun().apply(datum));
                    value = value.replace(label, escape);
                    vsb.append(separator).append(label).append(value).append(label);
                }
                sb.append(line).append(vsb.deleteCharAt(0));
            }
        }
        csv = sb.toString();
        return this;
    }

    /**
     * 将csv写入到输出流并指定字符集
     * @param os 输出流
     * @param charset 字符集
     * @return this
     * @throws IOException write异常
     */
    public ExcelUtil<T> writeCsv(OutputStream os, Charset charset) throws IOException {
        byte[] bytes = csv.getBytes(charset);
        os.write(bytes);
        os.flush();
        return this;
    }

    /**
     * 将csv写入到输出流（UTF-8）
     * @param os 输出流
     * @return this
     * @throws IOException write异常
     */
    public ExcelUtil<T> writeCsv(OutputStream os) throws IOException {
        return writeCsv(os, StandardCharsets.UTF_8);
    }

    /**
     * 将csv写入到文件并指定字符集
     * @param filePath 文件路径
     * @param charset 字符集
     * @return this
     * @throws IOException write异常
     */
    public ExcelUtil<T> writeCsv(String filePath, Charset charset) throws IOException {
        try (OutputStream os = new BufferedOutputStream(Files.newOutputStream(Paths.get(filePath)))) {
            writeCsv(os, charset);
        }
        return this;
    }

    /**
     * 将csv写入到文件（UTF-8）
     * @param filePath 文件路径
     * @return this
     * @throws IOException write异常
     */
    public ExcelUtil<T> writeCsv(String filePath) throws IOException {
        return writeCsv(filePath, StandardCharsets.UTF_8);
    }

    /**
     * 关闭Workbook流
     */
    public void close() {
        try {
            workbook.close();
        } catch (IOException e) {
            System.err.println("ExcelUtil -> workbook close error:" + e.getMessage());
        }
    }

    /**
     * 关闭Workbook流并将传入的输出流关闭
     */
    public void close(OutputStream os) throws IOException {
        this.close();
        os.close();
    }

    private int setExcelTitle(List<CellField> fieldList) {
        Class<?> tClass = currentSheet.gettClass();
        SXSSFSheet sheet = currentSheet.getSheet();
        ExcelClass excelClass = tClass.getAnnotation(ExcelClass.class);
        int startRow = excelClass != null ? 1 : 0;
        int[] ints = setExcelTitle(fieldList, startRow, 0);
        int maxRow = ints[0], maxCell = ints[1];
        mergeTitle(startRow, maxRow, fieldList);
        // 大标题
        if (excelClass != null) {
            SXSSFCell cell = sheet.createRow(0).createCell(0);
            String title = excelClass.title();
            if (title.isEmpty()) { title = excelClass.value(); }
            if (title.isEmpty()) {
                String className = tClass.getSimpleName();
                title = className.replaceAll("(?<![A-Z]|^)[A-Z]", " $0");
            }
            cell.setCellValue(title);
            // 设置标题样式
            CellStyle titleStyle = workbook.createCellStyle();
            IndexedColors background = excelClass.background();
            if (!IndexedColors.AUTOMATIC.equals(background)) {
                titleStyle.setFillForegroundColor(background.getIndex());
                titleStyle.setFillPattern(excelClass.pattern());
            }
            titleStyle.setAlignment(excelClass.horizontal());
            titleStyle.setVerticalAlignment(excelClass.vertical());
            // 设置标题字体
            Font font = workbook.createFont();
            font.setFontHeightInPoints(excelClass.fontSize());
            titleStyle.setFont(font);
            // 设置标题边框
            titleStyle.setBorderTop(BorderStyle.THICK);
            titleStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
            titleStyle.setBorderBottom(BorderStyle.THICK);
            titleStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            titleStyle.setBorderLeft(BorderStyle.THICK);
            titleStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            titleStyle.setBorderRight(BorderStyle.THICK);
            titleStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

            cell.setCellStyle(titleStyle);
            mergeRegion(0, 0, 0, maxCell, titleStyle);
        }
        return maxRow;
    }

    private int[] setExcelTitle(List<CellField> fields, int rowIndex, int cellIndex) {
        SXSSFSheet sheet = currentSheet.getSheet();
        CellStyle titleStyle = currentSheet.getTitleStyle();
        SXSSFRow row = sheet.getRow(rowIndex);
        int maxRow = rowIndex;
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        for (CellField field : fields) {
            SXSSFCell cell = row.getCell(cellIndex);
            if (cell == null) {
                cell = row.createCell(cellIndex);
            }
            String title = field.getTitle();
            cell.setCellValue(title);
            cell.setCellStyle(titleStyle);
            setColumnWidth(cell, field.getWidth() * 256);

            List<CellField> cellFields = field.getCellFields();
            if (cellFields != null) {
                if (title == null || title.isEmpty()) {
                    int[] area = setExcelTitle(cellFields, rowIndex, cellIndex);
                    maxRow = Math.max(maxRow, area[0]);
                    cellIndex = area[1] + 1;
                    continue;
                }
                int[] area = setExcelTitle(cellFields, rowIndex + 1, cellIndex);
                maxRow = Math.max(maxRow, area[0]);
                if (area[1] < cellIndex) {
                    continue;
                }
                mergeRegion(rowIndex, rowIndex, cellIndex, cellIndex = area[1], titleStyle);
            } else {
                field.setIndex(cellIndex);
            }
            currentSheet.operateTitle(cell);
            cellIndex++;
        }
        row.setHeightInPoints(currentSheet.getTitleHeight());
        return new int[]{maxRow, cellIndex - 1};
    }

    private void mergeTitle(int rowIndex, int maxRow, List<CellField> fields) {
        if (maxRow > rowIndex) {
            CellStyle titleStyle = currentSheet.getTitleStyle();
            for (CellField field : fields) {
                List<CellField> cellFields = field.getCellFields();
                if (cellFields == null) {
                    int index = field.getIndex();
                    mergeRegion(rowIndex, maxRow, index, index, titleStyle);
                } else {
                    String title = field.getTitle();
                    if (title == null || title.isEmpty()) {
                        mergeTitle(rowIndex, maxRow, cellFields);
                        continue;
                    }
                    mergeTitle(rowIndex + 1, maxRow, cellFields);
                }
            }
        }
    }

    private void mergeRegion(int startRow, int endRow, int startCell, int endCell, CellStyle style) {
        SXSSFSheet sheet = currentSheet.getSheet();
        for (int i = startRow; i <= endRow; i++) {
            SXSSFRow row = sheet.getRow(i);
            for (int j = startCell; j <= endCell; j++) {
                SXSSFCell cell = row.getCell(j);
                if (cell == null) {
                    cell = row.createCell(j);
                }
                cell.setCellStyle(style);
            }
        }
        sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, startCell, endCell));
    }

    private <R> int setExcelData(Collection<R> dataList, List<CellField> cellFields, int rowIndex) {
        if (dataList == null || dataList.isEmpty()) {
            return rowIndex;
        }
        SXSSFSheet sheet = currentSheet.getSheet();
        float rowHeight = currentSheet.getRowHeight();
        // 设置数据
        for (R data : dataList) {
            if (data == null) {
                continue;
            }
            int rowIndexCopy = rowIndex;
            SXSSFRow row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }
            for (CellField cellField : cellFields) {
                Object value = cellField.getValueFun().apply(data);
                List<CellField> cellFieldChi = cellField.getCellFields();
                CellType cellType = cellField.getCellType();
                if (CellType.LIST.equals(cellType)) {
                    int rowI = setExcelData((Collection<?>) value, cellFieldChi, rowIndexCopy);
                    rowIndex = Math.max(rowIndex, rowI);
                } else if (CellType.OBJECT.equals(cellType)) {
                    setExcelData(Collections.singletonList(value), cellFieldChi, rowIndexCopy);
                } else {
                    int cellIndex = cellField.getIndex();
                    SXSSFCell cell = row.getCell(cellIndex);
                    if (cell == null) {
                        cell = row.createCell(cellIndex);
                    }
                    cell.setCellStyle(currentSheet.operateCellStyle(cellField, data));
                    cellField.setCellStyle(cell.getCellStyle());
                    cell.setCellType(cellType.getCellType());
                    setValue(cell, value, cellType);
                    setColumnWidth(cell, 0);
                    currentSheet.operateCell(cell, data);
                    currentSheet.operateCell(cellField.getKey(), cell, data);
                    currentSheet.operateCell(cellIndex, cell, data);
                    int[] ranges = {rowIndexCopy, cellIndex};
                    columnRangeMap.put(cellIndex, ranges);
                    String key = cellField.getKey();
                    if (key != null && !key.isEmpty()) {
                        columnRangeByKeyMap.put(key, ranges);
                    }
                }
            }
            row.setHeightInPoints(rowHeight);
            mergeDataRow(cellFields, rowIndexCopy, rowIndex);
            currentSheet.operateRow(row, data);
            rowIndex++;
        }
        return rowIndex - 1;
    }

    private void mergeDataRow(List<CellField> cellFields,int startRowIndex, int maxRowIndex) {
        if (startRowIndex < maxRowIndex) {
            for (CellField cellField : cellFields) {
                List<CellField> cellFieldList = cellField.getCellFields();
                if (cellFieldList == null) {
                    int columnIndex = cellField.getIndex();
                    mergeRegion(startRowIndex, maxRowIndex, columnIndex, columnIndex, cellField.getCellStyle());
                } else if (CellType.OBJECT.equals(cellField.getCellType())){
                    mergeDataRow(cellFieldList, startRowIndex, maxRowIndex);
                }
            }
        }
    }

    public void setColumnWidth(SXSSFCell cell, int columnWidth) {
        if (currentSheet.isAutoColumnWidth()) {
            int cellIndex = cell.getColumnIndex();
            SXSSFSheet sheet = currentSheet.getSheet();
            String valStr = cell.toString();
            int length = valStr.getBytes().length;
            if (length < maxWidthMap.computeIfAbsent(cellIndex, key -> length)) {
                return;
            }
            int oldColumnWidth = sheet.getColumnWidth(cellIndex);
            short fontSize = workbook.getFontAt(cell.getCellStyle().getFontIndex()).getFontHeightInPoints();
            sheet.autoSizeColumn(cellIndex);
            int newColumnWidth = sheet.getColumnWidth(cellIndex) + ((length - valStr.length()) * 9 * fontSize);
            sheet.setColumnWidth(cellIndex, Math.max(oldColumnWidth, newColumnWidth));
            maxWidthMap.put(cellIndex, length);
        } else if (columnWidth > 0){
            int cellIndex = cell.getColumnIndex();
            SXSSFSheet sheet = currentSheet.getSheet();
            sheet.setColumnWidth(cellIndex, columnWidth);
        }
    }

    private void setValue(SXSSFCell cell, Object value, CellType cellType) {
        if (value != null && !CellType.BLANK.equals(cellType)) {
            if (value instanceof LocalDate) {
                cell.setCellValue((LocalDate) value);
            } else if (value instanceof LocalDateTime) {
                cell.setCellValue((LocalDateTime) value);
            } else if (value instanceof Date) {
                cell.setCellValue((Date) value);
            } else if (value instanceof Calendar) {
                cell.setCellValue((Calendar) value);
            } else if (value instanceof RichTextString) {
                cell.setCellValue((RichTextString) value);
            } else if (CellType.NUMBER.equals(cellType)) {
                cell.setCellValue(Double.parseDouble(String.valueOf(value)));
            } else if (CellType.BOOLEAN.equals(cellType)) {
                cell.setCellValue(Boolean.parseBoolean(String.valueOf(value)));
            } else if (CellType.STRING.equals(cellType)) {
                cell.setCellValue(String.valueOf(value));
            } else if(CellType.FORMULA.equals(cellType)) {
                cell.setCellValue(String.valueOf(value));
            } else {
                cell.setCellValue(String.valueOf(value));
            }
        }
    }

}
