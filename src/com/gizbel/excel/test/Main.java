package com.gizbel.excel.test;

import java.io.File;
import java.util.List;

import com.gizbel.excel.enums.ExcelFactoryType;
import com.gizbel.excel.factory.Parser;

public class Main {

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

}
