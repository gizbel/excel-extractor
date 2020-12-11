package com.gizbel.excel.factory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.sql.Date;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.gizbel.excel.annotations.ExcelBean;
import com.gizbel.excel.annotations.ExcelColumnHeader;
import com.gizbel.excel.annotations.ExcelColumnIndex;
import com.gizbel.excel.enums.ExcelFactoryType;

/**
 * Builds the excel factory on the specified class, The specified class must
 * have annotation type ExcelBean.<br>
 * Dependencies are latest version of apache POI<br>
 *
 * @author Saket Kumar
 */

public class Parser {

    /**
     * Will hold the reference to the annotated class which is being populated.
     **/
    private Class clazz;

    /**
     * Excel factory type decides whether extraction is going to be column based
     * or index based. 
     */
    private ExcelFactoryType excelFactoryType;

    /**
     * Only applicable for column based extraction.<br>
     * If set to true then first row in the excel sheet will be neglected.
     **/
    private boolean skipHeader;

    /** Will store the reference to all the fields of the annotated class. **/
    private Map<String, Field> fieldsMap;

    /**
     * Will stop the row processing whenever an empty row is encountered, be
     * default it is set to true in the constructor.
     **/
    private boolean breakAfterEmptyRow;

    /**
     * Initialize the excel parser.<br>
     * This constructor will also save the annotated class fields in to a map
     * for future use.
     *
     * @param clazz
     * @param excelFactoryType
     * @throws Exception
     */
    public Parser(Class clazz, ExcelFactoryType excelFactoryType) throws Exception {

        this.clazz = clazz;
        this.excelFactoryType = excelFactoryType;
        this.breakAfterEmptyRow = true;

        /*
         * Check whether class has ExcelBean annotation present, If not present
         * then throw exception
         */
        if (clazz.isAnnotationPresent(ExcelBean.class)) {

            /*
             * Initialize the fields map as empty hash map, this will used to
             * store the reference to the class fields
             */
            this.fieldsMap = new HashMap<String, Field>();

            /* Get all declared fields for the annotated class */
            Field[] fields = clazz.getDeclaredFields();

            for (Field field : fields) {

                /*
                 * Based on excel factory type prepare the java reflection field map
                 */
                switch (this.excelFactoryType) {
                    case COLUMN_INDEX_BASED_EXTRACTION: this.prepareColumnIndexBasedFieldMap(field);break;
                    case COLUMN_NAME_BASED_EXTRACTION:   this.prepareColumnHeaderBasedFieldMap(field);break;
                }
            }

        } else {
            throw new Exception("Provided class is not annotated with ExcelBean");
        }
    }
    
    
    /**
     * Preapares the field Map based on the column index
     * @param field
     */
    private void prepareColumnIndexBasedFieldMap(Field field){
        if (field.isAnnotationPresent(ExcelColumnIndex.class)) {

            /*
             * Make the field accessible and save it into the fields map
             */
            field.setAccessible(true);
            ExcelColumnIndex column = field.getAnnotation(ExcelColumnIndex.class);
            String key = String.valueOf(column.columnIndex());
            this.fieldsMap.put(key, field);
            
        }
    }
    
    /**
     * Prepares the field Map based on the column header
     * @param field
     */
    private void prepareColumnHeaderBasedFieldMap(Field field){
        if(field.isAnnotationPresent(ExcelColumnHeader.class)){
            field.setAccessible(true);
            ExcelColumnHeader column = field.getAnnotation(ExcelColumnHeader.class);
            String key = column.columnHeader();
            this.fieldsMap.put(key, field);
        }
    }
    
    
    /**
     * Returns the dataType specified in the field 
     * @param field
     * @return String dataType
     */
    private String getDataTypeFor(Field field){
        String dataType = null;
        switch (this.excelFactoryType) {
            case COLUMN_INDEX_BASED_EXTRACTION: ExcelColumnIndex indexColumn = field.getAnnotation(ExcelColumnIndex.class);
                                                dataType = indexColumn.dataType();
                                                break;
            case COLUMN_NAME_BASED_EXTRACTION:  ExcelColumnHeader headerColumn = field.getAnnotation(ExcelColumnHeader.class);
                                                dataType = headerColumn.dataType();
                                                break;
        }
        return dataType;
    }
    
    
    /**
     * Returns the default value specified in the field 
     * @param field
     * @return String dataType
     */
    private String getDefaultValueFor(Field field){
        String defaultValue = null;
        switch (this.excelFactoryType) {
            case COLUMN_INDEX_BASED_EXTRACTION: ExcelColumnIndex indexColumn = field.getAnnotation(ExcelColumnIndex.class);
                                                defaultValue = indexColumn.defaultValue();
                                                break;
            case COLUMN_NAME_BASED_EXTRACTION:  ExcelColumnHeader headerColumn = field.getAnnotation(ExcelColumnHeader.class);
                                                defaultValue = headerColumn.defaultValue();
                                                break;
        }
        return defaultValue;
    }
    

    /**
     * Reads and convert valid excel file into required format<br>
     * Will only process the first sheet, split multiple sheets into multiple
     * files
     * 
     * @param file
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     * @throws ParseException
     * @throws IllegalArgumentException
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    public List<Object> parse(File file) throws InvalidFormatException, IOException, InstantiationException,
            IllegalAccessException, IllegalArgumentException, ParseException {
        List<Object> result = new ArrayList<>();
        InputStream istrm = new FileInputStream(file);
        Workbook invoiceWorkbook = WorkbookFactory.create(istrm);

        // Currently processing for one sheet, we can loop here for multiple
        // sheets
        Sheet sheet = invoiceWorkbook.getSheetAt(0);

        if (excelFactoryType == ExcelFactoryType.COLUMN_NAME_BASED_EXTRACTION) {
            // Fetch the first row and save the field map with column index for
            // corresponding column headers
            Row firstRow = sheet.getRow(0);
            for (Cell column : firstRow) {

                Field field = this.fieldsMap.get(column.getStringCellValue());
                if (field != null) {
                    this.fieldsMap.remove(column.getStringCellValue());
                    this.fieldsMap.put(String.valueOf(column.getColumnIndex()), field);
                }
            }
        }

        for (Row row : sheet) {

            if (excelFactoryType == ExcelFactoryType.COLUMN_INDEX_BASED_EXTRACTION) {
                if (row.getRowNum() == 0 && skipHeader)
                    continue;
            } else if (excelFactoryType == ExcelFactoryType.COLUMN_NAME_BASED_EXTRACTION) {
                if (row.getRowNum() == 0)
                    continue;
            }

            // Process all non empty rows
            if (!isEmptyRow(row)) {
                Object beanObj = this.getBeanForARow(row);
                result.add(beanObj);
            } else {
                // If empty row found and user has opted to break whenever empty
                // row encountered then break the loop
                if (this.breakAfterEmptyRow)
                    break;
            }
        }
        istrm.close();
        return result;
    }


    /**
     * Fetches the cell details from the each row and sets its values based on
     * the instance variable defined by the annotation
     * 
     * @param row
     * @return Clazz object
     * @throws IllegalAccessException
     * @throws InstantiationException
     * @throws ParseException
     * @throws IllegalArgumentException
     */
    public Object getBeanForARow(Row row)
            throws InstantiationException, IllegalAccessException, IllegalArgumentException, ParseException {

        final Object classObj = this.clazz.newInstance();
        for (int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null) {
                if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        java.util.Date date = cell.getDateCellValue();
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                        cell.setCellValue(new SimpleDateFormat("dd-MM-yyyy").format(date));
                    } else
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                } else
                    cell.setCellType(Cell.CELL_TYPE_STRING);

                String value = cell.getStringCellValue() == null ? null : cell.getStringCellValue().trim();
                this.setCellValueBasedOnDesiredExcelFactoryType(classObj, value, i);
            } else
                this.setCellValueBasedOnDesiredExcelFactoryType(classObj, null, i);
        }

        return classObj;
    }

    
    /**
     * Parse the cell values to their specified dataType and sets into the java class field
     * @param classObj
     * @param columnValue
     * @param columnIndex
     * @throws IllegalArgumentException
     * @throws IllegalAccessException
     * @throws ParseException
     */
    private void setCellValueBasedOnDesiredExcelFactoryType(Object classObj, String columnValue, int columnIndex)
            throws IllegalArgumentException, IllegalAccessException, ParseException {

        Field field = this.fieldsMap.get(String.valueOf(columnIndex));
        if (field != null) {

            //If column value is null or empty then try to put the default value
            if( columnValue==null || columnValue.trim().isEmpty())
                columnValue = this.getDefaultValueFor(field);
            
            /**
             * Based on the dataType Specified convert it to primitive value<br>
             * But make sure that columnvalue is not null or empty
             **/
            if(columnValue!=null && !columnValue.trim().isEmpty()){
                String dataType = this.getDataTypeFor(field);
                switch (dataType) {
                case "int":
                    field.set(classObj, Integer.parseInt(columnValue));
                    break;
                case "long":
                    field.set(classObj, Long.parseLong(columnValue));
                    break;
                case "bool":
                    field.set(classObj, Boolean.parseBoolean(columnValue));
                    break;
                case "double":
                    Double data = Double.parseDouble(columnValue);
                    field.set(classObj, data);
                    break;
                case "date":
                    field.set(classObj, this.dateParser(columnValue));
                    break;
                default:
                    field.set(classObj, columnValue);
                    break;
                }
            }
        }

    }

    /**
     * Parses the date columns in dd-MM-YYYY format. Customize the format in
     * case.
     * 
     * @param value
     * @return
     */
    private Date dateParser(String value) {
        if (value != null && !value.isEmpty()) {
            String[] formats = new String[] { "dd-MM-yyyy" };
            java.util.Date date;
            try {
                date = DateUtils.parseDate(value, formats);
                return new Date(date.getTime());
            } catch (ParseException e) {
                e.printStackTrace();
            }
            return null;
        } else
            return null;
    }


    /**
     * Checks whether an encountered row is empty or not.
     * @param row
     * @return
     */
    boolean isEmptyRow(Row row) {
        boolean isEmptyRow = true;
        for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
            Cell cell = row.getCell(cellNum);
            if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK && StringUtils.isNotBlank(cell.toString())) {
                isEmptyRow = false;
            }
        }
        return isEmptyRow;
    }

    public boolean isBreakAfterEmptyRow() {
        return breakAfterEmptyRow;
    }

    public void setBreakAfterEmptyRow(boolean breakAfterEmptyRow) {
        this.breakAfterEmptyRow = breakAfterEmptyRow;
    }


    public boolean isSkipHeader() {
        return skipHeader;
    }


    public void setSkipHeader(boolean skipHeader) {
        this.skipHeader = skipHeader;
    }
}
