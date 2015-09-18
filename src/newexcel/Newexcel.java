/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package newexcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class Newexcel {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) 
    {
        //Create a new Workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Student data");
        //Create the data for the excel sheet
       Map<String, Object[]> data = new TreeMap<String, Object[]>();
       data.put("1", new Object[] {"ID", "FIRSTNAME", "LASTNAME"});
       data.put("2", new Object[] {1, "Randy", "Maven"});
       data.put("3", new Object[] {2, "Raymond", "Smith"});
        
         //Iterate over data and write it to the sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
       for (String key : keyset)
        {
            Row row = sheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
                Cell cell = row.createCell(cellnum++);
                if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
         try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("D://new_demo.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        
    }
    
}

