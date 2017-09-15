package com.gizbel.excel.annotations;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

/**
 * Class once annotated with this, can be used to perform excel column data
 * binding to the instance variables of the annotated class.<br>
 * This annotation is required to be present on the class.<br>
 * <b>Must have zero argument constructor</b>
 * 
 * @author Saket Kumar
 */
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelBean {

}
