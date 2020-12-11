package com.gizbel.excel.test;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;
import java.util.List;

import com.gizbel.excel.enums.ExcelFactoryType;
import com.gizbel.excel.factory.Parser;


public class Main {

    public static void main(String[] args) throws Exception {
        Parser parser = new Parser(Bean.class, ExcelFactoryType.COLUMN_INDEX_BASED_EXTRACTION);
        parser.setSkipHeader(true);
        File xlsFile = new File("test/inv.xlsx");
        String newFilePath = "target/inv.xlsx";
        List<Object> result = parser.parse(xlsFile);
        for (Object obj : result) {
            Bean bean = (Bean) obj;
            System.out.println(bean.toString());
            System.out.println();
        }
        //Files.move(Paths.get(xlsFile.toPath().toString()), Paths.get(newFilePath), REPLACE_EXISTING);
    }

}
