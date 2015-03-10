package org.xl.excel.parser;

import java.io.File;
import java.io.PrintStream;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.xl.excel.components.Sheet;

/**
 * Generic ExcelParser class which is the Factory Creator for
 * the various XLSX and XLS Parsers
 *
 * @author meulmees: May 16, 2013, 1:40:50 PM
 * @version $Revision:$, submitted by $Author:$
 */
public abstract class ExcelParser {
    public static final int READ_ALL = -1;

    /**
     * Returns the List of Sheet Objects which represents the loaded
     * Excel File.
     * <p>
     * <b>NOTE:</b> You must first invoke the ExcelParser.process() method
     * to read the file into memory before the SheetList is populated
     */
    public abstract List<Sheet> getSheetList();

    /**
     * Process the specified Excel File.
     * <p>
     * This will read through the entire Excel File and load all Sheet
     * Content
     * into Memory which can be retrieved using the
     * ExcelParser.getSheetList() method
     *
     * @param ignoreBlankRows
     * @param useCellFormatting
     */
    public abstract void process(boolean ignoreBlankRows, boolean
            useCellFormatting) throws RuntimeException, InvalidFormatException;

    /**
     * Process the specified Excel File.
     * <p>
     * This will only load the specified Sheet into memory.<br>
     * 0 Indexed.
     *
     * @param ignoreBlankRows
     * @param useCellFormatting
     * @param sheetNumber
     */
    public abstract void process(boolean ignoreBlankRows, boolean
            useCellFormatting, int index) throws RuntimeException,
            InvalidFormatException;

    /**
     * Creates a Parser which will write the XLSX Excel File Contents to the
     * specified PrintStream.
     *
     * @param xlsxFile
     * @param output
     * @param minColumns
     * @param maxRows
     * @return XLSXParser
     */
    public static XLSXParser createXLSXtoCSVConverter(File xlsxFile,
                                                      PrintStream output, int minColumns, int maxRows) {
        return new XLSXParser(xlsxFile, output, minColumns, maxRows);
    }

    /**
     * Creates a Parser which will write the XLSX Excel File Contents to the
     * specified PrintStream
     *
     * @param xlsxFile
     * @param output
     * @return XLSXParser
     */
    public static XLSXParser createXLSXtoCSVConverter(File xlsxFile,
                                                      PrintStream output) {
        return new XLSXParser(xlsxFile, output, READ_ALL, READ_ALL);
    }

    /**
     * Creates a Parser which will load the XLSX Excel File Contents into
     * memory with
     * the specified maximum number of Rows
     *
     * @param xlsxFile
     * @param maxRows
     * @return XLSXParser
     */
    public static XLSXParser createXLSXtoRowArrayList(File xlsxFile, int
            maxRows) {
        return new XLSXParser(xlsxFile, READ_ALL, maxRows);
    }

    /**
     * Creates a Parser which will load the entire XLSX Excel File Contents
     * into memory
     *
     * @param xlsxFile
     * @return XLSXParser
     */
    public static XLSXParser createXLSXtoRowArrayList(File xlsxFile) {
        return new XLSXParser(xlsxFile, READ_ALL, READ_ALL);
    }

    /**
     * Creates a Parser which will load the filtered Columns in the XLSX
     * Excel File
     * into memory with the specified maximum number of Rows
     *
     * @param xlsxFile
     * @param maxRows
     * @param columnFilter
     * @return XLSXParser
     */
    public static XLSXParser createXLSXtoRowArrayList(File xlsxFile, int
            maxRows, List<String> columnFilter) {
        return new XLSXParser(xlsxFile, READ_ALL, maxRows, columnFilter);
    }

    /**
     * Creates a Parser which will load the filtered Columns in the XLSX
     * Excel File
     * into memory
     *
     * @param xlsxFile
     * @param columnFilter
     * @return XLSXParser
     */
    public static XLSXParser createXLSXtoRowArrayList(File xlsxFile,
                                                      List<String> columnFilter) {
        return new XLSXParser(xlsxFile, READ_ALL, READ_ALL, columnFilter);
    }

    /**
     * Creates a Parser which will load the entire XLS Excel File Contents
     * into memory
     *
     * @param xlsFile
     * @return XLSParser
     */
    public static XLSParser createXLSParser(File xlsFile) {
        return new XLSParser(xlsFile);
    }

    /**
     * Creates a Parser which will load the XLS Excel File Contents into
     * memory with
     * the specified maximum rows
     *
     * @param xlsFile
     * @param maxRows
     * @return XLSParser
     */
    public static XLSParser createXLSParser(File xlsFile, int maxRows) {
        return new XLSParser(xlsFile, maxRows);
    }
}
