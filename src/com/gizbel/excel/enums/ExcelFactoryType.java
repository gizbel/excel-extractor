package com.gizbel.excel.enums;

import com.gizbel.excel.annotations.ExcelBean;
import com.gizbel.excel.annotations.ExcelColumnIndex;

/**
 * Defines the excel factory type.<br>
 * Based on this ExcelFactory, you can choose whether you want to extract cell
 * values based on (columns names) or (column index).
 * 
 * @author Saket Kumar
 * @see ExcelBean
 * @see ExcelColumnIndex
 */
public enum ExcelFactoryType {

    COLUMN_INDEX_BASED_EXTRACTION, COLUMN_NAME_BASED_EXTRACTION;
}
