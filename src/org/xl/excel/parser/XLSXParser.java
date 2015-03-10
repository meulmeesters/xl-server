package org.xl.excel.parser;

import java.io.File;
import java.io.InputStream;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.xl.excel.components.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

public class XLSXParser extends ExcelParser {

    private Logger LOGGER = LoggerFactory.getLogger(XLSXParser.class);

    public static final String MAX_ROW_CODE = "Reached Maximum Row";
    public static final String MAX_ROW_CODE_UNSPEC = "Reached Maximum UnSpecified Row";

    private File xlsxFile;
    private OPCPackage xlsxPackage;
    private PrintStream output;
    private int minColumns;
    private int maxRows;
    private List<Sheet> sheetList;
    private List<String> columnFilter;

        protected XLSXParser(File xlsxFile, PrintStream output,int
        minColumns,int maxRows){
        this.xlsxFile = xlsxFile;
        this.output = output;
        this.minColumns = minColumns;
        this.sheetList = new ArrayList<Sheet>();
        this.maxRows = maxRows;
    }
        protected XLSXParser(File xlsxFile, int minColumns, int maxRows){
        this.xlsxFile = xlsxFile;
        this.output = null;
        this.minColumns = minColumns;
        this.sheetList = new ArrayList<Sheet>();
        this.maxRows = maxRows;
    }
        protected XLSXParser(File xlsxFile, int minColumns, int maxRows,
        List<String> columnFilter){
        this.xlsxFile = xlsxFile;
        this.output = null;
        this.minColumns = minColumns;
        this.sheetList = new ArrayList<Sheet>();
        this.maxRows = maxRows;
        this.columnFilter = columnFilter;
    }
/**
 * Returns the List of Sheet Objects which represents the loaded
 * Excel File.
 * <p>
 * <b>NOTE:</b> You must first invoke the ExcelParser.process() method
 * to read the file into memory before the SheetList is populated
 *
 */
        @Override
        public List<Sheet> getSheetList () {
        return sheetList;
    }
/**
 * Process the specified XLSX File.
 * <p>
 * This will read through the entire Excel File and load all Sheet
 Content
 * into Memory which can be retrieved using the
 ExcelParser.getSheetList() method
 *
 */
        @Override
        public void process ( boolean ignoreBlankRows,boolean
        useCellFormatting)throws RuntimeException, InvalidFormatException {
        try {
            displayFilters();
            this.xlsxPackage = OPCPackage.open(xlsxFile.getPath(),
                    PackageAccess.READ);
            read(ignoreBlankRows, useCellFormatting, READ_ALL);
        } finally {
            if (null != this.xlsxPackage) {
                try {
                    this.xlsxPackage.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    }
/**
 * Process the specified XLSX File.
 * <p>
 * This will only load the specified Sheet into memory.<br>
 * 0 Indexed.
 *
 * @param sheetNumber
 */
        @Override
        public void process ( boolean ignoreBlankRows,boolean
        useCellFormatting,int sheetNumber)throws RuntimeException,
            InvalidFormatException {
        try {
            displayFilters();
            this.xlsxPackage = OPCPackage.open(xlsxFile.getPath(),
                    PackageAccess.READ);
            read(ignoreBlankRows, useCellFormatting, sheetNumber);
        } finally {
            if (null != this.xlsxPackage) {
                try {
                    this.xlsxPackage.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private void read(boolean ignoreBlankRows, boolean useCellFormatting,
                      int sheetNum) throws RuntimeException {
        try {
            ReadOnlySharedStringsTable strings = new
                    ReadOnlySharedStringsTable(this.xlsxPackage);
            XSSFReader xssfReader = new
                    XSSFReader(this.xlsxPackage);
            StylesTable styles = xssfReader.getStylesTable();
            XSSFReader.SheetIterator iter =
                    (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            int index = 0;
            while (iter.hasNext()) {
                InputStream stream = iter.next();
                if ((READ_ALL == sheetNum) || (index ==
                        sheetNum)) {
                    String sheetName = iter.getSheetName();
                    if (null != output) {
                        this.output.println();
                        this.output.println(sheetName +
                                " [index=" + index + "]:");
                    }
                    readSheet(ignoreBlankRows,
                            useCellFormatting, index, styles, strings, stream, sheetName);
                    stream.close();
                }
                ++index;
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private void readSheet(boolean ignoreBlankRows, boolean
            useCellFormatting, int index, StylesTable styles, ReadOnlySharedStringsTable
                                   strings,
                           InputStream sheetInputStream,
                           String sheetName) throws RuntimeException {
        ExcelWorkSheetHandler_CSV contentHandler = null;
        try {
            InputSource sheetSource = new
                    InputSource(sheetInputStream);
            SAXParserFactory saxFactory =
                    SAXParserFactory.newInstance();
            SAXParser saxParser = saxFactory.newSAXParser();
            XMLReader sheetParser = saxParser.getXMLReader();
            contentHandler = new ExcelWorkSheetHandler_CSV(styles,
                    strings, minColumns, output, maxRows, columnFilter, ignoreBlankRows,
                    useCellFormatting);
            sheetParser.setContentHandler(contentHandler);
            sheetParser.parse(sheetSource);
        } catch (RuntimeException e) {
            if (MAX_ROW_CODE.equals(e.getMessage())) {
                LOGGER.info("Reached Specified Maximum Allowed Row Count" +
                        maxRows + " on Sheet" + index + " - " + sheetName);
            } else {
                throw new RuntimeException(e);
            }
        } catch (OutOfMemoryError e) {
            LOGGER.info("Reached Maximum Allowed Memory Usage: ArrayList size" +
                    contentHandler.getValueList().size() + " on Sheet" + index + " - " + sheetName);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        if (contentHandler.getHeaderList().size() > 0) {
            if (null != columnFilter) {
                List<String> filteredHeaderList = new
                        ArrayList<String>();
                for (String header :
                        contentHandler.getHeaderList()) {
                    if (columnFilter.contains(header)) {
                        filteredHeaderList.add(header);
                    }
                }
                this.sheetList.add(new Sheet(sheetName, index,
                        contentHandler.getColumnTypes(),
                        filteredHeaderList,
                        contentHandler.getValueList()));
            } else {
                this.sheetList.add(new Sheet(sheetName, index,
                        contentHandler.getColumnTypes(),
                        contentHandler.getHeaderList(),
                        contentHandler.getValueList()));
            }
        }
    }

    private void displayFilters() {
        if (null != columnFilter) {
            StringBuffer msg = new StringBuffer("Applying Column Filters:");
            for (int j = 0; j < columnFilter.size(); j++) {
                msg.append("'" + columnFilter.get(j) + "'");
                msg.append((j < columnFilter.size() - 1) ? "," :
                        "");
            }
            LOGGER.info(msg.toString());
        }
    }
}
