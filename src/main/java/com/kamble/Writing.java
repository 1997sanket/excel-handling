package com.kamble;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class Writing {
    public static void main(String[] args) {


        // Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        // Create a blank Excel Sheet
        XSSFSheet sheet = workbook.createSheet("student details");

        Map<String, Object[]> data = new TreeMap<>();

        //heading
        data.put("1", new Object[]{"ID", "FIRST_NAME", "LAST_NAME"});

        //actual data
        data.put("2", new Object[]{1, "John", "Doe"});
        data.put("3", new Object[]{2, "Jane", "Doe"});
        data.put("4", new Object[]{3, "James", "Doe"});

        Set<String> keySet = data.keySet();

        int rowNum = 0;

        for(String key : keySet) {

            Row row = sheet.createRow(rowNum++);

            Object[] arr = data.get(key);

            int colNum = 0;
            for(Object o : arr) {
                Cell cell = row.createCell(colNum++);

                if(o instanceof Integer) cell.setCellValue((Integer)o);

                else cell.setCellValue((String)o);
            }
        }



        //Now writing to a file
        try {
            // Writing the workbook
            FileOutputStream out = new FileOutputStream(
                    new File("student.xlsx"));
            workbook.write(out);

            // Closing file output connections
            out.close();

            // Console message for successful execution of program
            System.out.println(
                    "gfgcontribute.xlsx written successfully on disk.");
        }

        catch (Exception e) {
            e.printStackTrace();
        }
    }
}
