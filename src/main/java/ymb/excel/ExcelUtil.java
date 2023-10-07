package ymb.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import ymb.excel.annotation.ExcelClass;

import java.io.*;
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
public final class ExcelUtil<T> {
    private final SheetOperate<T> sheetOperate;
    private final Stack<SheetOperate<?>> otherSheet = new Stack<>();
    private XSSFWorkbook workbook;
    private SheetOperate<?> currentSheet;

    public ExcelUtil(Class<T> tClass) {
        this.workbook = new XSSFWorkbook();
        this.sheetOperate = SheetOperate.create(tClass);
        this.sheetOperate.setWorkbook(this.workbook);
        otherSheet.add(this.sheetOperate);
    }

    public ExcelUtil(Class<T> tClass, String sheetName) {
        this.workbook = new XSSFWorkbook();
        this.sheetOperate = SheetOperate.create(tClass, sheetName);
        this.sheetOperate.setWorkbook(workbook);
        otherSheet.add(this.sheetOperate);
    }

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    /**
     * 设置数据源
     * @param data 数据源
     * @return this
     */
    public ExcelUtil<T> setDataList(List<T> data) {
        sheetOperate.setData(data);
        return this;
    }

    /**
     * 设置表头的字体大小
     * @param titleSize 字体大小
     * @return this
     */
    public ExcelUtil<T> setTitleSize(short titleSize) {
        sheetOperate.setTitleSize(titleSize);
        return this;
    }

    /**
     * 设置数据的字体大小
     * @param valueSize 字体大小
     * @return this
     */
    public ExcelUtil<T> setValueSize(short valueSize) {
        sheetOperate.setValueSize(valueSize);
        return this;
    }

    /**
     * 设置表头的行高
     * @param titleHeight 行高
     * @return this
     */
    public ExcelUtil<T> setTitleHeight(short titleHeight) {
        sheetOperate.setTitleHeight(titleHeight);
        return this;
    }

    /**
     * 设置数据的行高
     * @param valueHeight 行高
     * @return this
     */
    public ExcelUtil<T> setValueHeight(short valueHeight) {
        sheetOperate.setValueHeight(valueHeight);
        return this;
    }

    /**
     * 设置表头样式
     * @param titleStyle (CellStyle) -> void
     * @return this
     */
    public ExcelUtil<T> setTitleStyle(Consumer<XSSFCellStyle> titleStyle) {
        sheetOperate.setTitleStyle(titleStyle);
        return this;
    }

    /**
     * 设置数据样式
     * @param valueStyle (CellStyle) -> void
     * @return this
     */
    public ExcelUtil<T> setValueStyle(Consumer<XSSFCellStyle> valueStyle) {
        sheetOperate.setValueStyle(valueStyle);
        return this;
    }

    /**
     * 操作某一列数据的样式 (设置Cell时调用)
     * @param index 列索引
     * @param valueStyle (CellStyle, value) -> void
     * @return this
     */
    public ExcelUtil<T> operateValueStyle(int index, BiConsumer<XSSFCellStyle, Object> valueStyle) {
        sheetOperate.operateValueStyle(index, valueStyle);
        return this;
    }

    /**
     * 操作表头，每次设置表头之后执行
     * @param operateTitle (Cell) -> void
     * @return this
     */
    public ExcelUtil<T> operateTitle(Consumer<XSSFCell> operateTitle) {
        sheetOperate.operateTitle(operateTitle);
        return this;
    }

    /**
     * 操作数据，每次设置数据之后执行
     * @param operateValue (Cell, data) -> void
     * @return this
     */
    public ExcelUtil<T> operateValue(BiConsumer<XSSFCell, Object> operateValue) {
        sheetOperate.operateValue(operateValue);
        return this;
    }

    /**
     * 操作Sheet，在数据生成完之后执行
     * @param operateSheet (Sheet, dataList) -> void
     * @return this
     */
    public ExcelUtil<T> operateSheet(BiConsumer<XSSFSheet, List<T>> operateSheet) {
        sheetOperate.operateSheet(operateSheet);
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
        for (SheetOperate<?> operate : otherSheet) {
            List<CellField> fields = operate.getFields();
            this.currentSheet = operate;
            int maxRow = setExcelTitle(fields);
            setExcelData(operate.getData(), fields, maxRow + 1);
            operate.operateSheet();
        }
        return this;
    }

    /**
     * 执行并转成byte数组
     * @return byte数组
     * @throws IOException write异常
     */
    public byte[] toByteArray() throws IOException {
        try (XSSFWorkbook workbook = execute().workbook; ByteArrayOutputStream os = new ByteArrayOutputStream()) {
            workbook.write(os);
            os.flush();
            return os.toByteArray();
        }
    }

    /**
     * 执行并写入OutputStream
     * @param os OutputStream流
     * @throws IOException write异常
     */
    public void write(OutputStream os) throws IOException {
        try (XSSFWorkbook workbook = execute().workbook) {
            workbook.write(os);
            os.flush();
        }
    }

    /**
     * 执行并写入文件
     * @param filePath 文件路径
     * @throws IOException write异常
     */
    public void writeFile(String filePath) throws IOException {
        try (OutputStream os = new BufferedOutputStream(Files.newOutputStream(Paths.get(filePath)))) {
            write(os);
        }
    }

    private int setExcelTitle(List<CellField> fieldList) {
        Class<?> tClass = currentSheet.gettClass();
        XSSFSheet sheet = currentSheet.getSheet();
        ExcelClass excelClass = tClass.getAnnotation(ExcelClass.class);
        int startRow = excelClass != null ? 1 : 0;
        int[] ints = setExcelTitle(fieldList, startRow, 0);
        int maxRow = ints[0], maxCell = ints[1];
        mergeTitle(startRow, maxRow, fieldList);
        // 标题
        if (excelClass != null) {
            XSSFCell cell = sheet.createRow(0).createCell(0);
            String title = excelClass.title();
            if (title.isEmpty()) { title = excelClass.value(); }
            if (title.isEmpty()) {
                String className = tClass.getSimpleName();
                title = className.replaceAll("(?<![A-Z]|^)[A-Z]", " $0");
            }
            cell.setCellValue(title);
            // 设置标题样式
            XSSFCellStyle titleStyle = workbook.createCellStyle();
            IndexedColors background = excelClass.background();
            if (!IndexedColors.AUTOMATIC.equals(background)) {
                titleStyle.setFillForegroundColor(background.getIndex());
                titleStyle.setFillPattern(excelClass.pattern());
            }
            titleStyle.setAlignment(excelClass.horizontal());
            titleStyle.setVerticalAlignment(excelClass.vertical());
            // 设置标题字体
            XSSFFont font = workbook.createFont();
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
        XSSFSheet sheet = currentSheet.getSheet();
        XSSFCellStyle titleStyle = currentSheet.getTitleStyle();
        XSSFRow row = sheet.getRow(rowIndex);
        int maxRow = rowIndex;
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        for (CellField field : fields) {
            XSSFCell cell = row.getCell(cellIndex);
            if (cell == null) {
                cell = row.createCell(cellIndex);
            }
            cell.setCellValue(field.getTitle());
            cell.setCellStyle(titleStyle);
            sheet.setColumnWidth(cellIndex, field.getWidth() * 256);
            List<CellField> cellFields = field.getCellFields();
            if (cellFields != null) {
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
            XSSFCellStyle titleStyle = currentSheet.getTitleStyle();
            for (CellField field : fields) {
                List<CellField> cellFields = field.getCellFields();
                if (cellFields == null) {
                    int index = field.getIndex();
                    mergeRegion(rowIndex, maxRow, index, index, titleStyle);
                } else {
                    mergeTitle(rowIndex + 1, maxRow, cellFields);
                }
            }
        }
    }

    private void mergeRegion(int startRow, int endRow, int startCell, int endCell, XSSFCellStyle style) {
        XSSFSheet sheet = currentSheet.getSheet();
        for (int i = startRow; i <= endRow; i++) {
            XSSFRow row = sheet.getRow(i);
            for (int j = startCell; j <= endCell; j++) {
                XSSFCell cell = row.getCell(j);
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
        XSSFSheet sheet = currentSheet.getSheet();
        float valueHeight = currentSheet.getValueHeight();
        // 设置数据
        for (R data : dataList) {
            int rowIndexCopy = rowIndex;
            XSSFRow row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }
            for (CellField cellField : cellFields) {
                Object value = cellField.getValueFun().apply(data);
                List<CellField> cellFieldChi = cellField.getCellFields();
                if (cellFieldChi != null) {
                    int rowI = setExcelData((Collection<?>) value, cellFieldChi, rowIndexCopy);
                    rowIndex = Math.max(rowIndex, rowI);
                } else {
                    int cellIndex = cellField.getIndex();
                    XSSFCell cell = row.getCell(cellIndex);
                    if (cell == null) {
                        cell = row.createCell(cellIndex);
                    }
                    cell.setCellStyle(currentSheet.operateValueStyle(cellField, value));
                    cellField.setCellStyle(cell.getCellStyle());
                    CellType cellType = cellField.getCellType();
                    cell.setCellType(cellType);
                    setValue(cell, value, cellType);
                    currentSheet.operateValue(cell, data);
                }
            }
            row.setHeightInPoints(valueHeight);
            for (CellField cellField : cellFields) {
                if (rowIndexCopy < rowIndex && cellField.getCellFields() == null) {
                    int columnIndex = cellField.getIndex();
                    mergeRegion(rowIndexCopy, rowIndex, columnIndex, columnIndex, cellField.getCellStyle());
                }
            }
            rowIndex++;
        }
        return rowIndex - 1;
    }

    private void setValue(XSSFCell cell, Object value, CellType cellType) {
        if (value != null) {
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
            } else if (CellType.NUMERIC.equals(cellType)) {
                cell.setCellValue(Double.parseDouble(String.valueOf(value)));
            } else if (CellType.BOOLEAN.equals(cellType)) {
                cell.setCellValue(Boolean.parseBoolean(String.valueOf(value)));
            } else if (!CellType.BLANK.equals(cellType)) {
                cell.setCellValue(String.valueOf(value));
            }
        }
    }

}
