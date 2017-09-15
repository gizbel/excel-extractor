package com.gizbel.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * This annotation is used to fetch the value from the Excel cell to a
 * particular java class field based on column index. <b>Not applicable for inherited fields</b>
 * 
 * @author Saket Kumar
 * @see ExcelColumnIndex#dataType()
 * @see ExcelColumnIndex#defaultValue()
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelColumnIndex {

    /**
     * Zero based column index can be used to pull cell data in the annotated
     * field.<br>
     * For instance : Any field annotated with columnIndex("2"), that field will
     * be initialized with the value contained in column number 2 starting from 0.
     * @return
     */
    String columnIndex();

    /**
     * Specify the cell data type, if dataType is provided the extracted cell
     * value is parsed and converted into specified dataType. possible values
     * are <br>
     * <ul>
     * <li>"int" : returns the Integer value</li>
     * <li>"long" : returns the long value</li>
     * <li>"bool" : returns the boolean representation</li>
     * <li>"string" returns string representation, <b>it's by default</b>
     * <li>"double" returns the double value</li>
     * <li>"date" returns the java.util.date object</li>
     * </ul>
     * 
     * @return
     */
    String dataType() default "string";

    /**
     * Specifies the default value, if data is not found in the cell ( either
     * blank or null ), then it will put the default value<br>
     * <b>Default values will also be casted in their dataType</b>
     * 
     * @return
     */
    String defaultValue() default "";
}
