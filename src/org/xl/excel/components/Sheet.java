package org.xl.excel.components;

import org.apache.commons.lang3.StringUtils;

import java.util.List;


/**
 * The Sheet Object Contains all information regarding an Excel
 * File Sheet. This includes the Sheet Name, Index, Headers,
 * Row Count, List of Row Information, and also Column Types
 *
 * @author meulmees: May 16, 2013, 1:33:24 PM
 * @version $Revision:$, submitted by $Author:$
 */
public class Sheet {
    private int sheetIndex;
    private String sheetName;
    private List<String> rowList;
    private List<String> headerList;
    private List<String> columnTypes;

    public Sheet(String sheetName,
                 int sheetIndex,
                 List<String> columnTypes,
                 List<String> headerList,
                 List<String> rowList) {
        this.sheetName = sheetName;
        this.sheetIndex = sheetIndex;
        this.columnTypes = columnTypes;
        this.headerList = headerList;
        this.rowList = rowList;
    }

    /**
     * Returns the Sheet Index of the Sheet Object.
     * <p>
     * 0 Indexed
     *
     * @return SheetIndex
     */
    public int getSheetIndex() {
        return sheetIndex;
    }

    protected void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    /**
     * Returns the Sheet Name. This is whatever the User has named their
     * Sheet in Excel.
     * <p>
     * By Default Excel names them Sheet1, Sheet2, etc.
     *
     * @return SheetName
     */
    public String getSheetName() {
        return sheetName;
    }

    protected void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    /**
     * Returns the Number of Rows with Data in the Sheet.
     *
     * @return Number of Rows
     */
    public int getRowCount() {
        return rowList.size();
    }

    /**
     * Returns the list of Rows as list of Comma Separated Strings
     *
     * @return RowList
     */
    public List<String> getRowList() {
        return rowList;
    }

    protected void setRowList(List<String> valueList) {
        this.rowList = valueList;
    }

    /**
     * Returns the list of Headers in the Excel File.
     * <p>
     * This is the first row in the file which contains Data
     *
     * @return Headers
     */
    public List<String> getHeaderList() {
        return headerList;
    }

    protected void setHeaderList(List<String> headerList) {
        this.headerList = headerList;
    }

    /**
     * Returns the list of Column Types. This is a guess and should
     * not be relied on 100%.
     * <p>
     * The types are:<br>
     * - string
     * - numeric
     * - date
     *
     * @return ColumnTypes
     */
    public List<String> getColumnTypes() {
        return columnTypes;
    }

    protected void setColumnTypes(List<String> columnTypes) {
        this.columnTypes = columnTypes;
    }

    /**
     * Returns an array of the cell contents within the specified row
     * number.
     * <p>
     * 0 indexed
     *
     * @param rowNumber
     * @return String[]
     */
    public String[] getCellValues(int rowNumber, boolean withQuotes) {
        String[] cellValues = getRowList().get(rowNumber).split(",(?=([^\"]*\"[^\"]*\")*[^\"]*$)");
        if (withQuotes) {
            return cellValues;
        } else {
            String[] qCellValues = new String[cellValues.length];
            int index = 0;
            for (String cellValue : cellValues) {
                qCellValues[index++] = StringUtils.remove(cellValue, '"');
            }
            return qCellValues;
        }
    }

    /**
     * Reads the Header List for the current sheet and tries to find the
     * column
     * which contains the specified columnHeader. You can specify to
     * failSilently
     * which will return -1 if it does not find the column header or
     * setting
     * failSilently to false throws a RuntimeException if it does not find
     * the column
     * header.
     *
     * @param columnHeader
     * @param failSilently
     * @return columnIndex
     */
    public int getColumnIndex(String columnHeader, boolean failSilently) {
        for (int i = 0; i < headerList.size(); i++) {
            String header = headerList.get(i);
            if (header.equalsIgnoreCase(columnHeader))
                return i;
        }
        if (!failSilently)
            throw new RuntimeException("Column Header: '" + columnHeader +
                    "'not found on Sheet"+getSheetName()+" - index:"+getSheetIndex());
        else
        return -1;
    }
}