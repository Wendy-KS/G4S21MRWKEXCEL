/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package g4s21mrwkexcel;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
/**
 *
 * @author sonia
 */
public class G4S21MRWKEXCEL {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        Workbook wb = new HSSFWorkbook();
        try ( OutputStream fileOut = new FileOutputStream("miarchivo.xls")) {
            Sheet sheet1 = wb.createSheet("Primer Hoja");
            Sheet sheet2 = wb.createSheet("Segunda Hoja");
            Sheet sheet = wb.createSheet("Tercer Hoja");
            Row row = sheet.createRow(0); //crear fila. se establece el indice a 0 hasta N                           
            row.createCell(0).setCellValue("Num"); // Columna A  
            row.createCell(1).setCellValue("Nombre"); // Columna B   
            row.createCell(2).setCellValue("Edad");// Columna C  
            row.createCell(3).setCellValue("Correo"); // Columna D 

            row = sheet.createRow(1); //crear fila 2
            row.createCell(0).setCellValue(1); // Columna A  
            row.createCell(1).setCellValue("Sonia Rafael"); // Columna B   
            row.createCell(2).setCellValue(32);// Columna C  
            row.createCell(3).setCellValue("sony@gmail.com"); // Columna D 
            
            row = sheet.createRow(2); //crear fila 3
            row.createCell(0).setCellValue(2); // Columna A  
            row.createCell(1).setCellValue("Oscar Rafael"); // Columna B   
            row.createCell(2).setCellValue(30);// Columna C  
            row.createCell(3).setCellValue("leafar@gmail.com"); // Columna D
            
            row = sheet.createRow(3); //crear fila 2
            row.createCell(0).setCellValue(3); // Columna A  
            row.createCell(1).setCellValue("Isaac Luna"); // Columna B   
            row.createCell(2).setCellValue(20);// Columna C  
            row.createCell(3).setCellValue("saac@gmail.com"); // Columna D
            
            
            wb.write(fileOut);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }
}
