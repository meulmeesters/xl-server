package org.xl.excel.parser;

import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 * This class handles parsing the Excel File contents and creating the
 * Data Structures which are stored in the Sheet Component
 *
 * @author meulmees: May 16, 2013, 1:46:05 PM
 * @version $Revision:$, submitted by $Author:$
 */
public class ExcelWorkSheetHandler_CSV extends DefaultHandler {
    private static Logger LOGGER =
            LoggerFactory.getLogger(ExcelWorkSheetHandler_CSV.class);

    /**
     * The type of the data value is indicated by an attribute on
     * the cell element; the value is in a "v" element within the cell.
     */
    enum xssfDataType {
        BOOL,
        ERROR,
        FORMULA,
        INLINESTR,
        SSTINDEX,
        NUMBER,
    }

    private StylesTable stylesTable;
    private ReadOnlySharedStringsTable sharedStringsTable;
    private final PrintStream output;
    private final int minColumnCount;
    private boolean vIsOpen;
    private xssfDataType nextDataType;
    private short formatIndex;
    private String formatString;
    private final DataFormatter formatter;
    private int thisColumn = -1;
    private int lastColumnNumber = -1;
    private StringBuffer value;
    private StringBuilder objCurrentRow = new StringBuilder();
    private List<String> valueList;
    private List<String> headerList;
    private List<String> columnTypes;
    private List<String> columnFilter;
    private int maxRows;
    private int currRowNum = 0;
    private boolean settingTypes = false;
    private boolean ignoreBlankRows = true;
    private boolean useCellFormatting = true;
    private int typesSet = 0;

    /**
     * Accepts objects needed while parsing.
     *
     * @param styles  Table of styles
     * @param strings Table of shared strings
     * @param cols    Minimum number of columns to show
     * @param target  Sink for output
     */
    public ExcelWorkSheetHandler_CSV(StylesTable styles,
                                     ReadOnlySharedStringsTable strings, int cols,
                                     PrintStream target, int maxRows, List<String> columnFilter, boolean
                                             ignorBlankRows,
                                     boolean useCellFormatting) {
        this.stylesTable = styles;
        this.sharedStringsTable = strings;
        this.minColumnCount = cols;
        this.output = target;
        this.value = new StringBuffer();
        this.nextDataType = xssfDataType.NUMBER;
        this.formatter = new DataFormatter();
        this.valueList = new ArrayList<String>();
        this.headerList = new ArrayList<String>();
        this.columnTypes = new ArrayList<String>();
        this.maxRows = maxRows;
        this.columnFilter = columnFilter;
        this.ignoreBlankRows = ignorBlankRows;
        this.useCellFormatting = useCellFormatting;
    }

    public List<String> getValueList() {
        return valueList;
    }

    public List<String> getHeaderList() {
        return headerList;
    }

    public List<String> getColumnTypes() {
        return columnTypes;
    }

    /**
     * (non-Javadoc)* @see
     * org.xml.sax.helpers.DefaultHandler#startElement(java.lang.String,
     * java.lang.String,java.lang.String, org.xml.sax.Attributes)
     */
    public void startElement(String uri, String localName, String name,
                             Attributes attributes) throws SAXException {
        if ("inlineStr".equals(name) || "v".equals(name)) {
            vIsOpen = true;
// Clear contents cache
            value.setLength(0);
        }
// c => cell
        else if ("c".equals(name)) {
// Get the cell reference
            String r = attributes.getValue("r");
            int firstDigit = -1;
            for (int c = 0; c < r.length(); ++c) {
                if (Character.isDigit(r.charAt(c))) {
                    firstDigit = c;
                    break;
                }
            }
            thisColumn = nameToColumn(r.substring(0, firstDigit));
// Set up defaults.
            this.nextDataType = xssfDataType.NUMBER;
            this.formatIndex = -1;
            this.formatString = null;
            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");
            if ("b".equals(cellType))
                nextDataType = xssfDataType.BOOL;
            else if ("e".equals(cellType))
                nextDataType = xssfDataType.ERROR;
            else if ("inlineStr".equals(cellType))
                nextDataType = xssfDataType.INLINESTR;
            else if ("s".equals(cellType))
                nextDataType = xssfDataType.SSTINDEX;
            else if ("str".equals(cellType))
                nextDataType = xssfDataType.FORMULA;
            else if (cellStyleStr != null) {
/*
* It's a number, but possibly has a style
and/or special format.
* should use
org.apache.poi.ss.usermodel.BuiltinFormats,
* and I see javadoc for that at apache.org,
but it's not in the
* POI 3.5 Beta 5 jars. Scheduled to appear in
3.5 beta 6.
*/
                int styleIndex =
                        Integer.parseInt(cellStyleStr);
                XSSFCellStyle style =
                        stylesTable.getStyleAt(styleIndex);
                this.formatIndex = style.getDataFormat();
                this.formatString =
                        style.getDataFormatString();
                if (this.formatString == null)
                    this.formatString =
                            BuiltinFormats.getBuiltinFormat(this.formatIndex);
            }
        }
    }

    /**
     * (non-Javadoc)
     *
     * @see org.xml.sax.helpers.DefaultHandler#endElement(java.lang.String,
     * java.lang.String, java.lang.String)
     */
    public void endElement(String uri, String localName, String name)
            throws SAXException {
        String thisStr = "";
        boolean skipColumn = false;
// v => contents of a cell
        if ("v".equals(name)) {
            if (null != columnFilter && headerList.size() > 0) {
                if (!columnFilter.contains(headerList.get(thisColumn)))
                    skipColumn = true;
            }
            if (this.headerList.size() > 0 &&
                    this.getValueList().size() == 0) {
                settingTypes = true;
//If we're setting types and we skipped a column assume its a String column
                if ((thisColumn - 1) != lastColumnNumber) {
                    for (int i = 0; i < ((thisColumn) -
                            (lastColumnNumber + 1)); i++) {
                        columnTypes.add("String");
                        typesSet++;
                    }
                }
            } else {
                settingTypes = false;
                for (int i = 0; i < (headerList.size() - typesSet);
                     i++) {
                    columnTypes.add("String");
                    typesSet++;
                }
            }
            if (!skipColumn) {
// Process the value contents as required.
// Do now, as characters() may be called more than once
                switch (nextDataType) {
                    case BOOL:
                        char first = value.charAt(0);
                        thisStr = first == '0' ? "\"FALSE\"" :
                                "\"TRUE\"";
                        if (settingTypes) {
                            this.columnTypes.add("Boolean");
                            typesSet++;
                        }
                        break;
                    case ERROR:
                        thisStr = "\"ERROR:" + value.toString()
                                + '"';
                        if (settingTypes) {
                            this.columnTypes.add("Error");
                            typesSet++;
                        }
                        break;
                    case FORMULA:
// A formula could result in a string value,
// so always add double-quote characters.
                                        thisStr = '"' + value.toString() + '"';
                        if (settingTypes) {
                            this.columnTypes.add("String");
                            typesSet++;
                        }
                        break;
                    case INLINESTR:
// TODO: have seen an example of this, so it 's untested.
                        XSSFRichTextString rtsi = new
                                XSSFRichTextString(value.toString());
                        thisStr = '"' + rtsi.toString() + '"';
                        if (settingTypes) {
                            if (isDate(thisStr))
                                this.columnTypes.add("Date");
                            else
                                this.columnTypes.add("String");
                            typesSet++;
                        }
                        break;
                    case SSTINDEX:
                        String sstIndex = value.toString();
                        try {
                            int idx =
                                    Integer.parseInt(sstIndex);
                            XSSFRichTextString rtss = new
                                    XSSFRichTextString(sharedStringsTable.getEntryAt(idx));
                            thisStr = '"' + rtss.toString()
                                    + '"';
                            if (settingTypes) {
                                if (isDate(thisStr))
                                    this.columnTypes.add("Date");
                                else

                                    this.columnTypes.add("String");
                                typesSet++;
                            }
                        } catch (NumberFormatException ex) {
                            LOGGER.warn("Failed to parse SST index '" +
                                    sstIndex + "':" + ex.getLocalizedMessage(), ex);
                            if (null != output)
                                output.println("Failed to parse SST index '" +
                                        sstIndex + "':" + ex.toString());
                        }
                        break;
                    case NUMBER:
                        String n = value.toString();
                        if (this.formatString != null) {
                            if (useCellFormatting) {
                                thisStr = '"' +
                                        formatter.formatRawCellContents(Double.parseDouble(n), this.formatIndex,
                                                this.formatString) + '"';
                                thisStr =
                                        thisStr.replaceAll("[^-/.,\\d][\\s]", "");
                            } else {
                                thisStr = '"' + n + '"';
                            }
                            if (settingTypes) {
                                if (isDate(thisStr))
                                    this.columnTypes.add("Date");
                                else if (thisStr.contains("."))
                                    this.columnTypes.add("Number");
                                else
                                    this.columnTypes.add("Number");
                                typesSet++;
                            }
                        } else {
                            thisStr = '"' + n + '"';
                            if (settingTypes) {
                                if (thisStr.contains("\\."))
                                    this.columnTypes.add("Number");
                                else
                                    this.columnTypes.add("Number");
                                typesSet++;
                            }
                        }
                        break;
                    default:
                        thisStr = '"' + value.toString() + '"';
                        break;
                }
// Output after we've seen the string contents
// Emit commas for any fields that were missing on this row
                if (lastColumnNumber == -1) {
                    lastColumnNumber =
                            0;
                }
                for (int i = lastColumnNumber; i < thisColumn;
                     ++i) {
                    if (null != output) {
                        output.print(',');
                    } else {
                        if (null != columnFilter &&
                                headerList.size() > 0) {
                            if (columnFilter.contains(headerList.get(lastColumnNumber)))
                                objCurrentRow.append(",");
                        } else {
                            objCurrentRow.append(",");
                        }
                    }
                }
// Might be the empty string.
                if (null != output) {
                    output.print(thisStr);
                } else {
                    if (null != columnFilter &&
                            headerList.size() > 0) {
                        if (columnFilter.contains(headerList.get(thisColumn)))
                            objCurrentRow.append(thisStr);
                    } else {
                        objCurrentRow.append(thisStr);
                    }
                }
// Update column
                if (thisColumn > -1)
                    lastColumnNumber = thisColumn;
            }
        } else if ("row".equals(name)) {
            // We're onto a new row
            /**
             * Sometimes there are excel files where the first row has
             * trailing empty cells. In this case the Column Types size
             * does not match the Header Size. This next loop just adds
             * 'String' column types for the empty cells
             */int emptyTrailingColumns = this.headerList.size() -
                    this.columnTypes.size();
            for (int i = 0; i < emptyTrailingColumns; i++) {
                this.columnTypes.add("String");
            }

// Print out any missing commas if needed
            if (minColumnCount > 0) {
// Columns are 0 based
                if (lastColumnNumber == -1) {
                    lastColumnNumber =
                            0;
                }
                for (int i = lastColumnNumber;
                     i < (this.minColumnCount); i++) {
                    if (null != output) {
                        output.print(',');
                    } else {
                        objCurrentRow.append(",");
                    }
                }
            }
/**
 * This is what the logic for this section used to be
 *
 * if(headerList.size() >
 objCurrentRow.toString().split(",").length) {
 int size = objCurrentRow.toString().split(",").length;
 for(int i=0; i<(headerList.size()-size); i++) {
 objCurrentRow.append(", ");
 }
 }
 *
 * The issue with this is the simple split(",") since
 the cell contents can
 * have commas in them. We need to use the regex below
 which will do a proper
 * split. Then determine if there are empty cells which
 we have not accounted
 * for on this row and add any necessary commas
 */
            String[] cellValues =
                    objCurrentRow.toString().split(",(?=([^\"]*\"[^\"]*\")*[^\"]*$)");
            if (headerList.size() > cellValues.length) {
                int size = cellValues.length;
                for (int i = 0; i < (headerList.size() - size); i++) {
                    if (objCurrentRow.toString().equals(""))
                        objCurrentRow.append("\"\"");
                    objCurrentRow.append(",\"\"");
                }
            }

//I don't want to include empty rows
            if (!isRowBlank(objCurrentRow.toString()) && this.output
                    == null) {
                String tmp = objCurrentRow.toString();
                if (tmp.substring(tmp.length() - 1).equals("")) objCurrentRow = new
                        StringBuilder(tmp.substring(0, tmp.length() - 1));
//If both the valueList and headerList are empty then the current row
//will be considered the header row
                if (this.valueList.size() == 0 &&
                        this.headerList.size() == 0) {
                    String[] headers =
                            objCurrentRow.toString().split(",(?=([^\"]*\"[^\"]*\")*[^\"]*$)");
                    for (String header : headers) {
                        this.headerList.add(StringUtils.remove(header, '"'));
                    }
                } else {
                    this.valueList.add(objCurrentRow.toString());
                }
                objCurrentRow = new StringBuilder();
            }
            if (currRowNum++ >= maxRows && maxRows > 0) {
                throw new
                        RuntimeException(XLSXParser.MAX_ROW_CODE);
            }
            if (null != output)
                output.println();
            lastColumnNumber = -1;
        }
    }

    private boolean isRowBlank(String rowData) {
        if (!ignoreBlankRows)
            return false;

        String[] values =
                rowData.split(",(?=([^\"]*\"[^\"]*\")*[^\"]*$)");
        for (String value : values) {
            value = value.replaceAll("\"", "");
            if (value.length() > 0)
                return false;
        }
        return true;
    }

    /**
     * Captures characters only if a suitable element is open.
     * Originally was just "v"; extended for inlineStr also.
     */
    public void characters(char[] ch, int start, int length)
            throws SAXException {
        if (vIsOpen)
            value.append(ch, start, length);
    }

    /**
     * Converts an Excel column name like "C" to a zero-based index.
     *
     * @param name* @return Index corresponding to the specified name
     */
    private int nameToColumn(String name) {
        int column = -1;
        for (int i = 0; i < name.length(); ++i) {
            int c = name.charAt(i);
            column = (column + 1) * 26 + c - 'A';
        }
        return column;
    }

    private boolean isDate(String thisStr) {
        thisStr = thisStr.replace("\"", "");
        boolean result = false;
        if ((thisStr.split("-").length == 3)) {
            for (String partDate : thisStr.split("-")) {
                result = isInteger(partDate);
            }
        }
        if ((thisStr.split("/").length == 3)) {
            for (String partDate : thisStr.split("/")) {
                result = isInteger(partDate);
            }
        }
        if ((thisStr.split(".").length == 3)) {
            for (String partDate : thisStr.split(".")) {
                result = isInteger(partDate);
            }
        }
        return result;
    }

    private boolean isInteger(String value) {
        try {
            Integer.parseInt(value);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }
}
