/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Logica;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

    
/**
 *
 * @author APRENDIZ
 */
public class LLeerExcel {

    public static void LLeerExcel (File fileName){
        
        try {
            InputStream inp = new FileInputStream(fileName);
 
            XSSFWorkbook wb =  new XSSFWorkbook(inp);
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFRow row = sheet.getRow(13);
            Cell cell = row.getCell(13);

            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue("x");
 

            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\CLAUDIA\\Documents\\porteria\\formato.xlsx");
            wb.write(fileOut);
            
            fileOut.close();
        } catch (Exception e) {
            System.out.println(e);
        }
    }
    
    public static void LLeerExcel1 (File fileName){
        
        try {
            InputStream inp = new FileInputStream(fileName);
 
            XSSFWorkbook wb =  new XSSFWorkbook(inp);
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFRow row = sheet.getRow(2);
            Cell cell = row.getCell(13);

            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue("x");
 
            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\CLAUDIA\\Documents\\porteria\\formato.xlsx");
            wb.write(fileOut);
            
            fileOut.close();
        } catch (Exception e) {
            System.out.println(e);
        }
    }

   
    
}
