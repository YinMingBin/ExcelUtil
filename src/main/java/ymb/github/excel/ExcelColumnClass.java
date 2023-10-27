package ymb.github.excel;

import org.apache.poi.ss.usermodel.*;
import ymb.github.excel.annotation.AllFieldColumn;
import ymb.github.excel.annotation.ExcelColumn;

import java.lang.reflect.Field;

/**
 * @author YinMingBin
 */
@SuppressWarnings({"unused", "UnusedReturnValue"})
public class ExcelColumnClass {
    @ExcelColumn
    private static ExcelColumn defaultColumn;
    private String title;
    private short index;
    private CellType type;
    private String format;
    private int width;
    private boolean wrapText;
    private IndexedColors background;
    private IndexedColors color;
    private short size;
    private FillPatternType pattern;
    private HorizontalAlignment horizontal;
    private VerticalAlignment vertical;

    static {
        try {
            Field dataField = ExcelColumnClass.class.getDeclaredField("defaultColumn");
            defaultColumn = dataField.getAnnotation(ExcelColumn.class);
        } catch (NoSuchFieldException e) {
            System.err.println("get ExcelColumn default data fail...");
        }
    }

    private ExcelColumnClass() {}

    public static ExcelColumnClass build() {
        return getExcelColumnClass(defaultColumn);
    }

    private static ExcelColumnClass getExcelColumnClass(ExcelColumn column) {
        ExcelColumnClass columnClass = new ExcelColumnClass();
        String title = column.title();
        if (title.isEmpty()) {title = column.value();}
        columnClass.setTitle(title);
        columnClass.setIndex(column.index());
        columnClass.setType(column.type());
        columnClass.setFormat(column.format());
        columnClass.setColor(column.color());
        columnClass.setSize(column.size());
        columnClass.setWidth(column.width());
        columnClass.setWrapText(column.wrapText());
        columnClass.setBackground(column.background());
        columnClass.setPattern(column.pattern());
        columnClass.setHorizontal(column.horizontal());
        columnClass.setVertical(column.vertical());

        return columnClass;
    }

    static ExcelColumnClass getExcelColumn(AllFieldColumn fieldColumn, Field field) {
        ExcelColumn column = field.getAnnotation(ExcelColumn.class);
        if (column == null) {
            if (fieldColumn == null) {
                return null;
            }
            column = defaultColumn;
        }
        ExcelColumnClass columnClass = getExcelColumnClass(column);
        if (fieldColumn != null && column == defaultColumn) {
            columnClass.setWidth(fieldColumn.width());
            columnClass.setWrapText(fieldColumn.wrapText());
            columnClass.setBackground(fieldColumn.background());
            columnClass.setPattern(fieldColumn.pattern());
            columnClass.setHorizontal(fieldColumn.horizontal());
            columnClass.setVertical(fieldColumn.vertical());
        }
        return columnClass;
    }

    void settingStyle(CellStyle cellStyle, DataFormat dataFormat, Font font) {
        // 格式
        String format = this.getFormat();
        if (!format.isEmpty()) {
            cellStyle.setDataFormat(dataFormat.getFormat(format));
        }
        // 背景
        IndexedColors background = this.getBackground();
        if (!IndexedColors.AUTOMATIC.equals(background)) {
            cellStyle.setFillForegroundColor(background.getIndex());
            cellStyle.setFillPattern(this.getPattern());
        }
        // 字体
        font.setColor(this.getColor().getIndex());
        short size = this.getSize();
        if (size > 0) {
            font.setFontHeightInPoints(size);
        }
        // 对齐方式
        cellStyle.setAlignment(this.getHorizontal());
        cellStyle.setVerticalAlignment(this.getVertical());
        // 自动换行
        cellStyle.setWrapText(this.isWrapText());
    }

    public String getTitle() {
        return title;
    }

    public ExcelColumnClass setTitle(String title) {
        this.title = title;
        return this;
    }

    public short getIndex() {
        return index;
    }

    public ExcelColumnClass setIndex(short index) {
        this.index = index;
        return this;
    }

    public CellType getType() {
        return type;
    }

    public ExcelColumnClass setType(CellType type) {
        this.type = type;
        return this;
    }

    public String getFormat() {
        return format;
    }

    public ExcelColumnClass setFormat(String format) {
        this.format = format;
        return this;
    }

    public int getWidth() {
        return width;
    }

    public ExcelColumnClass setWidth(int width) {
        this.width = width;
        return this;
    }

    public boolean isWrapText() {
        return wrapText;
    }

    public ExcelColumnClass setWrapText(boolean wrapText) {
        this.wrapText = wrapText;
        return this;
    }

    public IndexedColors getBackground() {
        return background;
    }

    public ExcelColumnClass setBackground(IndexedColors background) {
        this.background = background;
        return this;
    }

    public IndexedColors getColor() {
        return color;
    }

    public ExcelColumnClass setColor(IndexedColors color) {
        this.color = color;
        return this;
    }

    public short getSize() {
        return size;
    }

    public ExcelColumnClass setSize(short size) {
        this.size = size;
        return this;
    }

    public FillPatternType getPattern() {
        return pattern;
    }

    public ExcelColumnClass setPattern(FillPatternType pattern) {
        this.pattern = pattern;
        return this;
    }

    public HorizontalAlignment getHorizontal() {
        return horizontal;
    }

    public ExcelColumnClass setHorizontal(HorizontalAlignment horizontal) {
        this.horizontal = horizontal;
        return this;
    }

    public VerticalAlignment getVertical() {
        return vertical;
    }

    public ExcelColumnClass setVertical(VerticalAlignment vertical) {
        this.vertical = vertical;
        return this;
    }
}