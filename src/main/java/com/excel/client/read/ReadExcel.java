package com.excel.client.read;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;

/**
 * @author Amit Salvi
 */
public class ReadExcel {

    public void reader(String fileName) {
        try
        {
            HSSFWorkbook myExcelBook = new HSSFWorkbook(new FileInputStream(fileName));
            HSSFSheet myExcelSheet = myExcelBook.getSheet("FirstSheet");

            System.out.println("-> myExcelSheet.getLastRowNum() -> "+ myExcelSheet.getLastRowNum());

            for(int rowIndex = 0; rowIndex <= myExcelSheet.getLastRowNum(); rowIndex++) {

                HSSFRow row = myExcelSheet.getRow(rowIndex);
                for(int cellIndex = 0; cellIndex< row.getLastCellNum(); cellIndex++) {
                    String cellValue = row.getCell(cellIndex).getStringCellValue();
                    System.out.println("At index : " + cellIndex + " value : " + cellValue);
                }
            }

            myExcelBook.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}
