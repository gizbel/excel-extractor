# Excel-extractor
Read excel file and return result as list of java bean.

# Include using maven
```
<dependency>
    <groupId>com.gizbel.excel</groupId>
    <artifactId>excel-extractor</artifactId>
    <version>1.0.2</version>
</dependency>

```

# Features :
    Column index based extraction : Use column index (Starts from 0 ) to fetch column data
    Column heading based extraction : Use header name ( First row name ) to fetch the column data

# Data conversion :
    Specify the cell data type, if dataType is specified the extracted cell
    value is parsed and converted into specified dataType. 
    possible values are 
    "int" : returns the Integer value
    "long" : returns the long value
    "bool" : returns the boolean representation
    "string" returns string representation,it's by default
    "double" returns the double value
    "date" returns the java.util.date object
 
 # Default Value :
    You can also specify the default values in case if any column value is null or empty.
    
    
# Example 
Preparing your POJO
```
@ExcelBean
public class Bean {

    @ExcelColumnIndex(columnIndex = "0", dataType = "double", defaultValue = "2.356")
    private Double fee;

    @ExcelColumnIndex(columnIndex = "1", dataType = "double")
    private Double totalCost;

    @ExcelColumnIndex(columnIndex = "2", dataType = "string")
    private String reference;

    @ExcelColumnIndex(columnIndex = "3", dataType = "string")
    private String invoiceNumber;
}
```

Just few lines in your main()
```
public static void main(String[] args) throws Exception {
        Parser parser = new Parser(Bean.class, ExcelFactoryType.COLUMN_INDEX_BASED_EXTRACTION);
        parser.setSkipHeader(true);
        List<Object> result = parser.parse(new File("test/inv.xlsx"));
        for (Object obj : result) {
            Bean bean = (Bean) obj;
            System.out.println(bean.toString());
            System.out.println();
        }
}
```
