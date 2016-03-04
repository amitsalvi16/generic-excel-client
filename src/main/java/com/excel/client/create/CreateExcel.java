package com.excel.client.create;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author Amit Salvi
 */
public class CreateExcel {

    /**
     * The method would generate the excel file with one row, one column
     */
    public void generate(String fileName) {

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("FirstSheet");

        HSSFRow rowhead = sheet.createRow((short)0);
        rowhead.createCell(0).setCellValue("No.");
        rowhead.createCell(1).setCellValue("Name");
        rowhead.createCell(2).setCellValue("Address");
        rowhead.createCell(3).setCellValue("Email");

        HSSFRow row = sheet.createRow((short)1);
        row.createCell(0).setCellValue("1");
        row.createCell(1).setCellValue("Ajay");
        row.createCell(2).setCellValue("Delhi");
        row.createCell(3).setCellValue("ajaybagri15@gmail.com");

        try {
            FileOutputStream fileOut = new FileOutputStream(fileName);
            workbook.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Your excel file has been generated!");
    }
}
