package com.kamble;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class Reading {

    public static void main(String[] args) {

        try{

            FileInputStream fis = new FileInputStream(new File("student.xlsx"));

            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheetAt(0); //get the first sheet

            Iterator<Row> rowIterator = sheet.iterator();

            while(rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.iterator();

                while(cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    // Checking the cell type and format
                    switch (cell.getCellType()) {

                        // Case 1
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.print((int)cell.getNumericCellValue()+ " | ");
                            break;

                        // Case 2
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(
                                    cell.getStringCellValue()
                                            + " | ");
                            break;
                    }
                }

                System.out.println("");
                fis.close();
            }

        } catch(Exception e) {
            e.printStackTrace();
        }
    }
}
