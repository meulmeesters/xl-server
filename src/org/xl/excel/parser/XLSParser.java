package org.xl.excel.parser;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * This class is an ExcelParser which is built using the
 * Apache POI Java API.
 * <p>
 * It will enable the reading of Excel file contents
 *
 * @author meulmees: Dec 20, 2012, 9:09:21 AM
 * @version $Revision: #8 $, submitted by $Author: meulmees $
 */
public class XLSParser extends ExcelParser {

    private static Logger LOGGER = LoggerFactory.getLogger(XLSParser.class);

    private int maxRows = -1;
    private File xlsFile;
    private List<org.xl.excel.components.Sheet> sheetList;

    protected XLSParser(File xlsFile) {
        this.xlsFile = xlsFile;
        sheetList = new ArrayList<>();
    }

    protected XLSParser(File xlsFile, int maxRows) {
        this.xlsFile = xlsFile;
        sheetList = new ArrayList<>();
        this.maxRows = maxRows;
    }

    /**
     * Returns the List of Sheet Objects which represents the loaded
     * Excel File.
     * <p>
     * <b>NOTE:</b> You must first invoke the ExcelParser.process() method
     * to read the file into memory before the SheetList is populated
     */
    @Override
    public List<org.xl.excel.components.Sheet> getSheetList() {
        return sheetList;
    }

    /**
     * Process the specified XLS File.
     * <p>
     * This will read through the entire Excel File and load all Sheet
     * Content
     * into Memory which can be retrieved using the
     * ExcelParser.getSheetList() method
     */
    @Override
    public void process(boolean ignoreBlankRows, boolean useCellFormatting) {
        LOGGER.info("Processing XLS file: " + this.xlsFile.getName());

        InputStream fis = null;
        try {
            fis = new FileInputStream(xlsFile);
            sheetList = readContentsAsList(ignoreBlankRows, useCellFormatting, fis);
        } catch (Exception e) {
            LOGGER.warn("Failed to process workbook: " + e.getLocalizedMessage(), e);
            throw new RuntimeException("Failed to process workbook: " + e.getLocalizedMessage());
        } finally {
            if (null != fis) {
                try {
                    fis.close();
                } catch (Exception e) {
                    LOGGER.warn("Failed to close file inputstream: "+e.getLocalizedMessage(), e);
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
    public void process(boolean ignoreBlankRows, boolean useCellFormatting,
                        int index) {
        LOGGER.info("Processing XLS file: " + this.xlsFile.getName());
        InputStream fis = null;
        try {
            fis = new FileInputStream(xlsFile);
            List<org.xl.excel.components.Sheet> tmpSheets = readContentsAsList(ignoreBlankRows, useCellFormatting, fis);
            sheetList = new ArrayList<org.xl.excel.components.Sheet>();
            sheetList.add(tmpSheets.get(index));
        } catch (Exception e) {
            LOGGER.warn("Failed to process workbook: " + e.getLocalizedMessage(), e);
            throw new RuntimeException("Failed to process workbook: " + e.getLocalizedMessage());
        } finally {
            if (null != fis) {
                try {
                    fis.close();
                } catch (Exception e) {
                    LOGGER.warn("Failed to close file inputstream: " + e.getLocalizedMessage(), e);
                }
            }
        }
    }

    private List<org.xl.excel.components.Sheet> readContentsAsList(boolean ignoreBlankRows, boolean useCellFormatting, InputStream fis) {
        List<org.xl.excel.components.Sheet> sheets = new ArrayList<org.xl.excel.components.Sheet>();
        List<String> rowList = null;
        List<String> headers = null;
        List<String> columnTypes = null;
        StringBuilder currentRowObj = new StringBuilder();
        Workbook workbook = null;
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;
        try {
            workbook = WorkbookFactory.create(fis);
            for (int sheetIndex = 0; sheetIndex <
                    workbook.getNumberOfSheets(); sheetIndex++) {
                sheet = workbook.getSheetAt(sheetIndex);
                rowList = new ArrayList<String>();
                headers = new ArrayList<String>();
                columnTypes = new ArrayList<String>();

                headers = getHeaders(workbook, sheetIndex);
                columnTypes = getColumnTypes(headers);

                //If they haven't made a maxRow request it will be -1. In this case
                int len = ((maxRows == -1) ? sheet.getLastRowNum() :
                        (Math.min(40, sheet.getLastRowNum())));
                for (int i = 1; i < len + 1; i++) {
                    row = sheet.getRow(i);
//Reset the variables
                    currentRowObj.setLength(0);
                    int lastCellNum = row.getLastCellNum();
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        cell = row.getCell(j);
                        if (cell != null) {
                            currentRowObj.append("\"" + cell.toString() + "\"");
                        } else {
                            currentRowObj.append("\"\"");
                        }
                        currentRowObj.append((j < lastCellNum - 1) ? "," : "");
                    }
//Add any missing elements
                    int missingEls = (headers.size() -
                            currentRowObj.toString().split(",(?=([^\"]*\"[^\"]*\")*[^\"]*$)").length);
                    for (int j = 0; j < missingEls; j++) {
                        currentRowObj.append(",\"\"");
                    }
                    if (ignoreBlankRows) {
                        if (!isRowBlank(currentRowObj.toString()))
                            rowList.add(currentRowObj.toString());
                    } else {
                        rowList.add(currentRowObj.toString());
                    }
                }
                sheets.add(new org.xl.excel.components.Sheet(sheet.getSheetName(),
                        sheetIndex,
                        columnTypes,//ColumnTypes
                        headers,
                        rowList));
            }
        } catch (Exception e) {
            LOGGER.warn("Failed to read excel file contents: " + e.getLocalizedMessage(), e);
            throw new RuntimeException("Failed to read excel file contents: " + e.getLocalizedMessage(), e);
        }
        return sheets;
    }

    private boolean isRowBlank(String rowData) {
        String[] values = rowData.split(",");
        for (String value : values) {
            if (!value.trim().equals("\"\""))
                return false;
        }
        return true;
    }

    private List<String> getColumnTypes(List<String> headers) {
        List<String> types = new ArrayList<String>();
        for (int i = 0; i < headers.size(); i++) {
            types.add("String");
        }
        return types;
    }

    /**
     * Retrieves All header's of the given sheet index.
     * <p>
     * A header is considered to be the first row which contains information
     * within a sheet up to 5 rows down. Beyond that The Excel Parser does
     * not consider
     * the data to be a header.
     *
     * @param sheetIndex
     * @return
     */
    private List<String> getHeaders(Workbook workbook, int sheetIndex) {
        List<String> headers = new ArrayList<String>();
        Sheet sheet = null;
        Row row = null;
        int i = 0;
        try {
            sheet = workbook.getSheetAt(sheetIndex);
            //Search for the first Row which contains information
            do {
            row = sheet.getRow(i++);
        } while (row == null && i < 5) ;
        if (row != null) {
            //Retrieve All Cell's with data in them
            for (i = 0; i < row.getLastCellNum(); i++) {
                Cell cell = row.getCell(i);
                if (cell != null) {
                    String header = cell.toString();
                    if (header != null)
                        headers.add(header);
                }
            }
        }
    } catch (Exception e) {
        LOGGER.warn("Failed to retrieve headers from sheet index " +
                sheetIndex + ":" + e.getLocalizedMessage(), e);
        throw new RuntimeException("Failed to retrieve headers from sheet index" +
                sheetIndex + ":" + e.getLocalizedMessage(), e);
    }

    return headers;
}
}
