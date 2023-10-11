package ymb.github.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import ymb.github.excel.annotation.AllFieldColumn;
import ymb.github.excel.annotation.ExcelColumn;

import java.lang.reflect.Field;

/**
 * @author YinMingBin
 */
class ExcelColumnClass {
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

    static ExcelColumnClass getExcelColumn(AllFieldColumn fieldColumn, Field field) {
        ExcelColumn column = field.getAnnotation(ExcelColumn.class);
        if (column == null) {
            if (fieldColumn == null) {
                return null;
            }
            column = defaultColumn;
        }
        ExcelColumnClass columnClass = new ExcelColumnClass();
        String title = column.title();
        if (title.isEmpty()) {title = column.value();}
        columnClass.setTitle(title);
        columnClass.setIndex(column.index());
        columnClass.setType(column.type());
        columnClass.setFormat(column.format());
        columnClass.setColor(column.color());
        columnClass.setSize(column.size());
        if (fieldColumn != null) {
            columnClass.setWidth(fieldColumn.width());
            columnClass.setWrapText(fieldColumn.wrapText());
            columnClass.setBackground(fieldColumn.background());
            columnClass.setPattern(fieldColumn.pattern());
            columnClass.setHorizontal(fieldColumn.horizontal());
            columnClass.setVertical(fieldColumn.vertical());
        } else {
            columnClass.setWidth(column.width());
            columnClass.setWrapText(column.wrapText());
            columnClass.setBackground(column.background());
            columnClass.setPattern(column.pattern());
            columnClass.setHorizontal(column.horizontal());
            columnClass.setVertical(column.vertical());
        }
        return columnClass;
    }

    void settingStyle(XSSFCellStyle cellStyle, XSSFDataFormat dataFormat) {
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
        XSSFFont font = cellStyle.getFont();
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

    public void setTitle(String title) {
        this.title = title;
    }

    public short getIndex() {
        return index;
    }

    public void setIndex(short index) {
        this.index = index;
    }

    public CellType getType() {
        return type;
    }

    public void setType(CellType type) {
        this.type = type;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public boolean isWrapText() {
        return wrapText;
    }

    public void setWrapText(boolean wrapText) {
        this.wrapText = wrapText;
    }

    public IndexedColors getBackground() {
        return background;
    }

    public void setBackground(IndexedColors background) {
        this.background = background;
    }

    public IndexedColors getColor() {
        return color;
    }

    public void setColor(IndexedColors color) {
        this.color = color;
    }

    public short getSize() {
        return size;
    }

    public void setSize(short size) {
        this.size = size;
    }

    public FillPatternType getPattern() {
        return pattern;
    }

    public void setPattern(FillPatternType pattern) {
        this.pattern = pattern;
    }

    public HorizontalAlignment getHorizontal() {
        return horizontal;
    }

    public void setHorizontal(HorizontalAlignment horizontal) {
        this.horizontal = horizontal;
    }

    public VerticalAlignment getVertical() {
        return vertical;
    }

    public void setVertical(VerticalAlignment vertical) {
        this.vertical = vertical;
    }
}