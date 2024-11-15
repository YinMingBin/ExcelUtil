package ymb.github.excel;

import java.util.Collection;

/**
 * @author YinMingBin
 */
public class DataValidationItem {
    private final int firstRow;
    private final int firstCol;
    private final int endRow;
    private final int endCol;
    private final Collection<String> data;

    public DataValidationItem(int firstRow, int firstCol, int endRow, int endCol, Collection<String> data) {
        this.firstRow = firstRow;
        this.firstCol = firstCol;
        this.endRow = endRow;
        this.endCol = endCol;
        this.data = data;
    }

    public int getFirstRow() {
        return firstRow;
    }

    public int getFirstCol() {
        return firstCol;
    }

    public int getEndRow() {
        return endRow;
    }

    public int getEndCol() {
        return endCol;
    }

    public Collection<String> getData() {
        return data;
    }
}
