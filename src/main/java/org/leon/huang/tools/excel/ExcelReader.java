package org.leon.huang.tools.excel;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

/**
 * Created by azhi on 20/4/11.
 */
public class ExcelReader {
    /**
     * 用于解析xls,xlsx两张格式的excel文件
     * @param path
     */
    public static void parseExcelFromClassPath(String path){
        if (StringUtils.isBlank(path)) {
            throw new RuntimeException("path can not null");
        }
        Workbook workbook = null;
        InputStream inputStream = null;
        try {
            inputStream = ExcelReader.class.getClassLoader().getResourceAsStream(path);
            if (path.endsWith(".xls")) {
                workbook = new HSSFWorkbook(inputStream);
            } else if (path.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(inputStream);
            } else {
                throw new RuntimeException("File format not supported");
            }
            int sheetCount = workbook.getNumberOfSheets();
            System.out.println("文件共有" + sheetCount + "个sheet");
            //源码中sheet是个list
            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while (sheetIterator.hasNext()) {
                Sheet sheet = sheetIterator.next();
                Iterator<Row> rowIterator = sheet.rowIterator();
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        String value = cell.getStringCellValue();
                        System.out.print(value + "   ");
                    }
                    System.out.println();
                }
                System.out.println("---------------------");
            }
        } catch (FileNotFoundException e) {
            throw new RuntimeException("File Not Found");
        } catch (IOException e) {
            throw new RuntimeException("IOException happened", e);
        }finally {
            IOUtils.closeQuietly(inputStream);
        }


    }

    public static void main(String[] args) {
        parseExcelFromClassPath("test.xlsx");

    }
}
