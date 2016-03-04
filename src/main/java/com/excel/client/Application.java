package com.excel.client;

import com.excel.client.create.CreateExcel;
import com.excel.client.read.ReadExcel;

/**
 * Main class for initiating application
 *
 * @author Amit Salvi
 */
public class Application {

    public static void main(String[] args) {

        String fileName = "src/main/resources/excel/Ajay.xls";

        CreateExcel createExcel = new CreateExcel();
        createExcel.generate(fileName);

        ReadExcel readExcel = new ReadExcel();
        readExcel.reader(fileName);
    }

    //TODO: Add a thread pool for continuously poll the incoming file
    //TODO: read the file, write basic contents
    //TODO: To create the file, give input like number of rows, column, fields to add etc and it will create from it
}
